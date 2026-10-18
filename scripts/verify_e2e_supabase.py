"""
End-to-End Corporate Verification Test Suite for Digital Kasbon (Supabase + Storage)
------------------------------------------------------------------------------------
Simulates the complete 7-step Kasbon lifecycle against live Supabase:
  Step 1: Create Kasbon Request + Upload Attachment to 'kasbon-attachments'
  Step 2: Manager Approval -> Verifies state transitions to WAITING_CASHIER_VERIF (Step 3)
  Step 3: Senior Cashier Verification -> Verifies state transitions to WAITING_MONEY_RECEIPT (Step 4)
  Step 4: Requester Money Receipt -> Verifies state transitions to WAITING_REALIZATION_INPUT (Step 5)
  Step 5: Requester Realization -> Verifies state transitions to WAITING_CASHIER_REA_VERIF (Step 6)
  Step 6: Cashier Realization Verif -> Verifies state transitions to WAITING_MGR_FINAL_CHECK (Step 7)
  Step 7: Manager Final Check -> Verifies state transitions to COMPLETED (Step 7)
  Cleanup: Removes temporary test record & storage object after verification.
"""
import os
import sys
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from supabase.migrate_from_sheets import SUPABASE_URL, SUPABASE_KEY

HEADERS = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=representation"
}

TEST_ID = "KBA644-TEST-999"

def get_dashboard_row(no_p):
    res = requests.get(
        f"{SUPABASE_URL}/rest/v1/v_kasbon_dashboard?no_pengajuan=eq.{no_p}",
        headers=HEADERS,
        timeout=15
    )
    res.raise_for_status()
    rows = res.json()
    return rows[0] if rows else None

def patch_kasbon(no_p, payload):
    res = requests.patch(
        f"{SUPABASE_URL}/rest/v1/kasbon_requests?no_pengajuan=eq.{no_p}",
        headers=HEADERS,
        json=payload,
        timeout=15
    )
    res.raise_for_status()
    return res.json()[0]

def main():
    print("[1/8] Testing Supabase Storage upload to 'kasbon-attachments'...")
    dummy_png = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\nIDATx\x9cc\x00\x01\x00\x00\x05\x00\x01\r\n-\xb4\x00\x00\x00\x00IEND\xaeB`\x82"
    storage_path = f"pengajuan/{TEST_ID}.png"
    up_res = requests.post(
        f"{SUPABASE_URL}/storage/v1/object/kasbon-attachments/{storage_path}",
        headers={
            "apikey": SUPABASE_KEY,
            "Authorization": f"Bearer {SUPABASE_KEY}",
            "Content-Type": "image/png",
            "x-upsert": "true"
        },
        data=dummy_png,
        timeout=15
    )
    assert up_res.status_code in (200, 201), f"Storage upload failed: {up_res.text}"
    public_url = f"{SUPABASE_URL}/storage/v1/object/public/kasbon-attachments/{storage_path}"
    print(f"  [OK] Uploaded attachment: {public_url}")

    print("[2/8] Testing Step 1 (Submit Pengajuan Baru)...")
    requests.delete(f"{SUPABASE_URL}/rest/v1/kasbon_requests?no_pengajuan=eq.{TEST_ID}", headers=HEADERS)
    ins_res = requests.post(
        f"{SUPABASE_URL}/rest/v1/kasbon_requests",
        headers=HEADERS,
        json={
            "no_pengajuan": TEST_ID,
            "kode_store": "A644",
            "email_request": "requestertrial@gmail.com",
            "dibayarkan_kepada": "User E2E Test",
            "nip_req": "121014",
            "departemen": "Operational",
            "nominal": 250000,
            "terbilang_req": "Dua Ratus Lima Puluh Ribu Rupiah",
            "keperluan": "Pembelian ATK Operasional Test",
            "bukti_lampiran_req": public_url,
            "janji_penyelesaian": "20/10/2026",
            "senior_cashier": "131356 - Onasih (requestertrial@gmail.com)",
            "manager_incharge": "003264 - Dadang Iskandar (fajarachmad15@gmail.com)"
        },
        timeout=15
    )
    assert ins_res.status_code == 201, f"Insert failed: {ins_res.text}"
    row = get_dashboard_row(TEST_ID)
    assert row["overall_status"] == "WAITING_MGR_APPROVAL" and row["current_step"] == 2
    assert row["manager_incharge_nik"] == "003264" and row["senior_cashier_nik"] == "131356"
    print(f"  [OK] Step 1 -> State: {row['overall_status']} (Step {row['current_step']})")

    print("[3/8] Testing Step 2 (Manager Approval)...")
    patch_kasbon(TEST_ID, {
        "timestamp_app": "2026-10-18T11:00:00+07:00",
        "nip_app": "003264",
        "status_approval": "Approved",
        "reason_reject_app": "-"
    })
    row = get_dashboard_row(TEST_ID)
    assert row["overall_status"] == "WAITING_CASHIER_VERIF" and row["current_step"] == 3
    print(f"  [OK] Step 2 -> State: {row['overall_status']} (Step {row['current_step']})")

    print("[4/8] Testing Step 3 (Senior Cashier Verification)...")
    patch_kasbon(TEST_ID, {
        "timestamp_ver": "2026-10-18T11:05:00+07:00",
        "nip_ver": "131356",
        "status_verifikasi": "Verifikasi Approved",
        "reason_reject_ver": "-"
    })
    row = get_dashboard_row(TEST_ID)
    assert row["overall_status"] == "WAITING_MONEY_RECEIPT" and row["current_step"] == 4
    print(f"  [OK] Step 3 -> State: {row['overall_status']} (Step {row['current_step']})")

    print("[5/8] Testing Step 4 (Requester Money Receipt Confirmation)...")
    patch_kasbon(TEST_ID, {
        "timestamp_rec": "2026-10-18T11:10:00+07:00",
        "nip_rec": "121014",
        "status_uang": "Sudah diterima"
    })
    row = get_dashboard_row(TEST_ID)
    assert row["overall_status"] == "WAITING_REALIZATION_INPUT" and row["current_step"] == 5
    print(f"  [OK] Step 4 -> State: {row['overall_status']} (Step {row['current_step']})")

    print("[6/8] Testing Step 5 (Requester Realization Input)...")
    patch_kasbon(TEST_ID, {
        "timestamp_rea": "2026-10-18T11:15:00+07:00",
        "nip_rea": "121014",
        "nominal_pembelanjaan_rea": 200000,
        "terbilang_rea1": "Dua Ratus Ribu Rupiah",
        "uang_dikembalikan_rea": 50000,
        "terbilang_rea2": "Lima Puluh Ribu Rupiah",
        "uang_diterima_rea": 0,
        "terbilang_rea3": "Nol Rupiah",
        "bukti_lampiran_rea": public_url,
        "status_realisasi": "Terrealisasi"
    })
    row = get_dashboard_row(TEST_ID)
    assert row["overall_status"] == "WAITING_CASHIER_REA_VERIF" and row["current_step"] == 6
    print(f"  [OK] Step 5 -> State: {row['overall_status']} (Step {row['current_step']})")

    print("[7/8] Testing Step 6 (Cashier Realization Verification)...")
    patch_kasbon(TEST_ID, {
        "timestamp_ver_rea": "2026-10-18T11:20:00+07:00",
        "nip_ver_rea": "131356",
        "uang_dikembalikan_ver_rea": 50000,
        "terbilang_ver_rea1": "Lima Puluh Ribu Rupiah",
        "uang_diterima_ver_rea": 0,
        "terbilang_ver_rea2": "Nol Rupiah",
        "status_verifikasi_rea": "Ya, Sesuai",
        "reason_ver_rea": "-"
    })
    row = get_dashboard_row(TEST_ID)
    assert row["overall_status"] == "WAITING_MGR_FINAL_CHECK" and row["current_step"] == 7
    print(f"  [OK] Step 6 -> State: {row['overall_status']} (Step {row['current_step']})")

    print("[8/8] Testing Step 7 (Manager Final Check -> COMPLETED) & Cleanup...")
    patch_kasbon(TEST_ID, {
        "timestamp_fin_cek": "2026-10-18T11:25:00+07:00",
        "nip_fin_cek": "003264",
        "status_qty_value_nota": "Ya, Sesuai",
        "reason_fin_cek_1": "-",
        "status_item": "Ya, Sesuai",
        "reason_fin_cek_2": "-",
        "revisi_bukti_realisasi": "-"
    })
    row = get_dashboard_row(TEST_ID)
    assert row["overall_status"] == "COMPLETED" and row["is_completed"] is True
    print(f"  [OK] Step 7 -> State: {row['overall_status']} (is_completed={row['is_completed']})")

    # Cleanup test record and test storage object
    requests.delete(f"{SUPABASE_URL}/rest/v1/kasbon_requests?no_pengajuan=eq.{TEST_ID}", headers=HEADERS)
    requests.delete(
        f"{SUPABASE_URL}/storage/v1/object/kasbon-attachments/{storage_path}",
        headers={"apikey": SUPABASE_KEY, "Authorization": f"Bearer {SUPABASE_KEY}"}
    )
    print("[SUCCESS] All 7 Lifecycle Steps, Triggers, Views, and Storage verified 100%!")

if __name__ == "__main__":
    main()
