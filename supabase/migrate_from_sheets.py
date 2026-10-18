"""
Enterprise Database Provisioning & Google Sheets Migration Script
-----------------------------------------------------------------
1. Provisions PostgreSQL tables, state-machine triggers, and views on Supabase.
2. Creates 'kasbon-attachments' storage bucket.
3. Migrates Master Store, Master User, and Historical Kasbon transactions from Google Sheets.
"""
import os
import json
import datetime
import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials

def load_env_file(filepath=".env"):
    if os.path.exists(filepath):
        with open(filepath, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))

load_env_file()

SUPABASE_URL = os.environ.get("SUPABASE_URL", "https://uvjtzkpyunfuzcbhubpg.supabase.co")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "")
SUPABASE_ACCESS_TOKEN = os.environ.get("SUPABASE_ACCESS_TOKEN", "")
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1TGsCKhBC0E0hup6RGVbGrpB6ds5Jdrp5tNlfrBORzaI")
PROJECT_REF = SUPABASE_URL.split("//")[-1].split(".")[0]

def run_sql_migration(sql_path: str):
    print(f"[*] Executing SQL migration: {sql_path} ...")
    with open(sql_path, "r", encoding="utf-8") as f:
        sql_query = f.read()

    res = requests.post(
        f"https://api.supabase.com/v1/projects/{PROJECT_REF}/database/query",
        headers={
            "Authorization": f"Bearer {SUPABASE_ACCESS_TOKEN}",
            "Content-Type": "application/json"
        },
        json={"query": sql_query},
        timeout=30
    )
    if res.status_code in (200, 201):
        print(f"[+] Successfully executed {sql_path}")
    else:
        raise RuntimeError(f"SQL Migration failed ({res.status_code}): {res.text}")

def supabase_upsert(table: str, rows: list, on_conflict: str):
    if not rows:
        return
    res = requests.post(
        f"{SUPABASE_URL}/rest/v1/{table}?on_conflict={on_conflict}",
        headers={
            "apikey": SUPABASE_KEY,
            "Authorization": f"Bearer {SUPABASE_KEY}",
            "Content-Type": "application/json",
            "Prefer": "resolution=merge-duplicates,return=minimal"
        },
        json=rows,
        timeout=30
    )
    if res.status_code not in (200, 201, 204):
        raise RuntimeError(f"Upsert into {table} failed ({res.status_code}): {res.text}")
    print(f"[+] Upserted {len(rows)} rows into {table}")

def parse_ts(val):
    if not val or str(val).strip() in ("", "-"):
        return None
    s = str(val).strip()
    try:
        dt = datetime.datetime.strptime(s, "%Y-%m-%d %H:%M:%S")
        return dt.strftime("%Y-%m-%dT%H:%M:%S+07:00")
    except ValueError:
        return None

def parse_int(val):
    if val is None or str(val).strip() in ("", "-"):
        return 0
    try:
        return int(float(str(val).replace(",", "").strip()))
    except ValueError:
        return 0

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    run_sql_migration(os.path.join(base_dir, "01_schema_tables.sql"))
    run_sql_migration(os.path.join(base_dir, "02_storage_setup.sql"))

    if not os.path.exists("service_account.json"):
        print("[!] service_account.json not found, skipping Google Sheets data import.")
        return

    print("[*] Connecting to Google Sheets for initial data migration...")
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)
    client = gspread.authorize(creds)
    sh = client.open_by_key(SPREADSHEET_ID)

    # 1. Extract Stores & Users from DATABASE_USER
    ws_user = sh.worksheet("DATABASE_USER")
    raw_users = ws_user.get_all_records()

    stores_map = {}
    users_payload = []

    for u in raw_users:
        kode_store = str(u.get("Kode_Store", "")).strip().upper()
        nama_store = str(u.get("Nama_Store", "")).strip()
        if not kode_store:
            continue
        stores_map[kode_store] = {
            "kode_store": kode_store,
            "nama_store": nama_store or kode_store,
            "is_active": True
        }
        nik_padded = str(u.get("NIK", "")).strip().zfill(6)
        role_val = str(u.get("Role", "")).strip()
        users_payload.append({
            "nik": nik_padded,
            "nama_lengkap": str(u.get("Nama Lengkap", "")).strip(),
            "email": str(u.get("Email", "")).strip().lower(),
            "role": role_val,
            "kode_store": kode_store,
            "password_pin": str(u.get("Password", "123123")).strip(),
            "is_active": True
        })

    supabase_upsert("kasbon_stores", list(stores_map.values()), "kode_store")
    supabase_upsert("kasbon_users", users_payload, "nik,kode_store,role")

    # 2. Migrate Historical Kasbon Records from DATA_KASBON_AZKO
    ws_kasbon = sh.worksheet("DATA_KASBON_AZKO")
    raw_kasbon = ws_kasbon.get_all_records()
    kasbon_payload = []

    for k in raw_kasbon:
        no_p = str(k.get("No_Pengajuan", "")).strip()
        kode_store = str(k.get("Kode_Store", "")).strip().upper()
        if not no_p or not kode_store:
            continue

        if kode_store not in stores_map:
            supabase_upsert("kasbon_stores", [{"kode_store": kode_store, "nama_store": kode_store, "is_active": True}], "kode_store")
            stores_map[kode_store] = True

        terbilang_ver_rea2 = str(k.get("Terbilang_Ver_Rea2", "")).strip()
        if terbilang_ver_rea2 == "Rupiah":
            terbilang_ver_rea2 = "Nol Rupiah"

        kasbon_payload.append({
            "no_pengajuan": no_p,
            "kode_store": kode_store,
            "timestamp_req": parse_ts(k.get("Timestamp_Req")) or datetime.datetime.now(datetime.timezone.utc).isoformat(),
            "email_request": str(k.get("Email_Request", "")).strip().lower(),
            "dibayarkan_kepada": str(k.get("Dibayarkan_Kepada", "")).strip(),
            "nip_req": str(k.get("NIP_Req", "")).strip().zfill(6),
            "departemen": str(k.get("Departemen", "")).strip(),
            "nominal": parse_int(k.get("Nominal")),
            "terbilang_req": str(k.get("Terbilang_Req", "")).strip(),
            "keperluan": str(k.get("Keperluan", "")).strip(),
            "bukti_lampiran_req": str(k.get("Bukti_Lampiran_Req", "-") or "-").strip(),
            "janji_penyelesaian": str(k.get("Janji_Penyelesaian", "")).strip(),
            "senior_cashier": str(k.get("Senior_Cashier", "")).strip(),
            "manager_incharge": str(k.get("Manager_Incharge", "")).strip(),
            "timestamp_app": parse_ts(k.get("Timestamp_App")),
            "nip_app": str(k.get("NIP_App", "")).strip().zfill(6) if k.get("NIP_App") else None,
            "status_approval": str(k.get("Status_Approval", "")).strip() or "Pending",
            "reason_reject_app": str(k.get("Reason_Reject_App", "-") or "-").strip(),
            "timestamp_ver": parse_ts(k.get("Timestamp_Ver")),
            "nip_ver": str(k.get("NIP_Ver", "")).strip().zfill(6) if k.get("NIP_Ver") else None,
            "status_verifikasi": str(k.get("Status_Verifikasi", "")).strip() or "Pending",
            "reason_reject_ver": str(k.get("Reason_Reject_Ver", "-") or "-").strip(),
            "timestamp_rec": parse_ts(k.get("Timestamp_Rec")),
            "nip_rec": str(k.get("NIP_Rec", "")).strip().zfill(6) if k.get("NIP_Rec") else None,
            "status_uang": str(k.get("Status_Uang", "")).strip() or "Pending",
            "timestamp_rea": parse_ts(k.get("Timestamp_Rea")),
            "nip_rea": str(k.get("NIP_Rea", "")).strip().zfill(6) if k.get("NIP_Rea") else None,
            "nominal_pembelanjaan_rea": parse_int(k.get("Nominal_Pembelanjaan_Rea")),
            "terbilang_rea1": str(k.get("Terbilang_Rea1", "")).strip(),
            "uang_dikembalikan_rea": parse_int(k.get("Uang_Dikembalikan_Rea")),
            "terbilang_rea2": str(k.get("Terbilang_Rea2", "")).strip(),
            "uang_diterima_rea": parse_int(k.get("Uang_Diterima_Rea")),
            "terbilang_rea3": str(k.get("Terbilang_Rea3", "")).strip(),
            "bukti_lampiran_rea": str(k.get("Bukti_Lampiran_Rea", "-") or "-").strip(),
            "status_realisasi": str(k.get("Status_Realisasi", "")).strip() or "Pending",
            "timestamp_ver_rea": parse_ts(k.get("Timestamp_Ver_Rea")),
            "nip_ver_rea": str(k.get("NIP_Ver_Rea", "")).strip().zfill(6) if k.get("NIP_Ver_Rea") else None,
            "uang_dikembalikan_ver_rea": parse_int(k.get("Uang_Dikembalikan_Ver_Rea")),
            "terbilang_ver_rea1": str(k.get("Terbilang_Ver_Rea1", "")).strip(),
            "uang_diterima_ver_rea": parse_int(k.get("Uang_Diterima_Ver_Rea")),
            "terbilang_ver_rea2": terbilang_ver_rea2,
            "status_verifikasi_rea": str(k.get("Status_Verifikasi_Rea", "")).strip() or "Pending",
            "reason_ver_rea": str(k.get("Reason", "-") or "-").strip(),
            "timestamp_fin_cek": parse_ts(k.get("Timestamp_Fin_Cek")),
            "nip_fin_cek": str(k.get("NIP_Fin_Cek", "")).strip().zfill(6) if k.get("NIP_Fin_Cek") else None,
            "status_qty_value_nota": str(k.get("Status_Qty_Value_Nota", "")).strip(),
            "reason_fin_cek_1": str(k.get("Reason_Fin_Cek_1", "-") or "-").strip(),
            "status_item": str(k.get("Status_Item", "")).strip(),
            "reason_fin_cek_2": str(k.get("Reason_Fin_Cek_2", "-") or "-").strip(),
            "revisi_bukti_realisasi": str(k.get("Revisi_Bukti_Realisasi", "-") or "-").strip()
        })

    supabase_upsert("kasbon_requests", kasbon_payload, "no_pengajuan")
    print("[OK] Database Provisioning & Migration Complete!")

if __name__ == "__main__":
    main()
