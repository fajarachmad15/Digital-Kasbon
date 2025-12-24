import streamlit as st
import datetime
import gspread
import io
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formataddr
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from oauth2client.service_account import ServiceAccountCredentials

# --- 1. SETTING HALAMAN & CSS ---
st.set_page_config(page_title="Kasbon Digital Petty Cash", layout="centered")

st.markdown("""
    <style>
    [data-testid="InputInstructions"] { display: none !important; }
    .label-container {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: -10px;
        padding-top: 10px;
    }
    .label-text { font-weight: 600; color: inherit; font-size: 15px; }
    .error-tag { color: #FF0000 !important; font-size: 13px; font-weight: bold; }
    div[data-testid="stWidgetLabel"] { display: none; }
    
    .store-header {
        color: #FF0000 !important;
        font-size: 1.25rem;
        font-weight: 600;
        margin-top: 1rem;
        margin-bottom: 1rem;
        display: block;
    }

    /* STYLING TOMBOL APPROVAL TRANSPARAN (Image 707050) */
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        color: white !important;
    }
    /* Approve (Hijau Transparan) */
    div.stColumn:nth-of-type(1) [data-testid="stButton"] button {
        background-color: rgba(40, 167, 69, 0.4) !important;
        border: 1px solid rgba(40, 167, 69, 0.6) !important;
    }
    /* Reject (Merah Transparan) */
    div.stColumn:nth-of-type(2) [data-testid="stButton"] button {
        background-color: rgba(220, 53, 69, 0.4) !important;
        border: 1px solid rgba(220, 53, 69, 0.6) !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- 2. KONFIGURASI ---
SENDER_EMAIL = "achmad.setiawan@kawanlamacorp.com"
APP_PASSWORD = st.secrets["APP_PASSWORD"] 
BASE_URL = "https://digital-kasbon-ahi.streamlit.app" 
SPREADSHEET_ID = "1TGsCKhBC0E0hup6RGVbGrpB6ds5Jdrp5tNlfrBORzaI"

WIB = datetime.timezone(datetime.timedelta(hours=7))

def get_creds():
    creds_dict = st.secrets["gcp_service_account"]
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    return ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)

def send_email_with_attachment(to_email, subject, message_body, attachment_file=None):
    try:
        msg = MIMEMultipart()
        msg['Subject'] = subject
        msg['From'] = formataddr(("Bot_KasbonPC_Digital <No-Reply>", SENDER_EMAIL))
        msg['To'] = to_email
        msg.attach(MIMEText(message_body, 'html'))
        if attachment_file is not None:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment_file.getvalue())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={attachment_file.name}')
            msg.attach(part)
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(SENDER_EMAIL, APP_PASSWORD)
            server.sendmail(SENDER_EMAIL, to_email, msg.as_string())
        return True
    except: return False

def terbilang(n):
    units = ["", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas"]
    if n == 0: return "" 
    if n < 12: res = units[n]
    elif n < 20: res = terbilang(n - 10) + " Belas"
    elif n < 100: res = units[n // 10] + " Puluh " + terbilang(n % 10)
    elif n < 200: res = "Seratus " + terbilang(n - 100)
    elif n < 1000: res = units[n // 100] + " Ratus " + terbilang(n % 100)
    elif n < 2000: res = "Seribu " + terbilang(n - 1000)
    elif n < 1000000: res = terbilang(n // 1000) + " Ribu " + terbilang(n % 1000)
    elif n < 1000000000: res = terbilang(n // 1000000) + " Juta " + terbilang(n % 1000000)
    return res.strip()

# --- 3. SESSION STATE ---
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'user_data' not in st.session_state: st.session_state.user_data = {}
if 'submitted' not in st.session_state: st.session_state.submitted = False
if 'data_ringkasan' not in st.session_state: st.session_state.data_ringkasan = {}
if 'show_errors' not in st.session_state: st.session_state.show_errors = False

# --- 4. LOGIKA LOGIN ---
if not st.session_state.logged_in and not st.query_params.get("id"):
    st.subheader("🔐 Login User")
    login_nip = st.text_input("NIK (Wajib 6 Digit)", max_chars=6)
    login_pass = st.text_input("Password", type="password")
    
    if st.button("Masuk", type="primary", use_container_width=True):
        if len(login_nip) == 6 and len(login_pass) >= 6:
            try:
                creds = get_creds()
                client = gspread.authorize(creds)
                db_sheet = client.open_by_key(SPREADSHEET_ID).worksheet("DATABASE_USER")
                records = db_sheet.get_all_records()
                
                # Validasi NIK dan Password (Case Sensitive)
                user = next((r for r in records if str(r['NIK']) == login_nip and str(r['Password']) == login_pass), None)
                
                if user:
                    st.session_state.logged_in = True
                    st.session_state.user_data = user
                    st.rerun()
                else:
                    st.error("NIK atau Password salah.")
            except Exception as e: st.error(f"Gagal akses database: {e}")
        else:
            st.warning("NIK wajib 6 digit & Password minimal 6 karakter.")
    st.stop()

# --- 5. TAMPILAN APPROVAL MANAGER ---
query_id = st.query_params.get("id")
if query_id:
    st.markdown('<span class="store-header">Portal Approval Manager</span>', unsafe_allow_html=True)
    try:
        creds = get_creds()
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID).worksheet("DATA_KASBON_AZKO")
        cell = sheet.find(query_id)
        row_data = sheet.row_values(cell.row)
        
        st.info(f"### Rincian Pengajuan: {query_id}")
        col_app1, col_app2 = st.columns(2)
        with col_app1:
            st.write(f"**Tgl Pengajuan:** {row_data[0]}")
            st.write(f"**Dibayarkan:** {row_data[4]} / {row_data[5]}")
            st.write(f"**Departemen:** {row_data[6]}")
        with col_app2:
            st.write(f"**Nominal:** Rp {int(row_data[7]):,}")
            # PERBAIKAN: Tambah Terbilang di Portal Approval
            st.write(f"**Terbilang:** {row_data[8]}")
            st.write(f"**Keperluan:** {row_data[9]}")
            st.write(f"**Janji Selesai:** {row_data[11]}")
            
        st.write(f"**Status:** `{row_data[14]}`")
        st.divider()
        
        if row_data[14] == "Pending":
            c1, c2 = st.columns(2)
            if c1.button("✓ APPROVE", use_container_width=True):
                sheet.update_cell(cell.row, 15, "APPROVED")
                st.success("Berhasil di-Approve!"); st.balloons()
                st.rerun()
            if c2.button("✕ REJECT", use_container_width=True):
                sheet.update_cell(cell.row, 15, "REJECTED")
                st.error("Pengajuan telah di-Reject.")
                st.rerun()
        else:
            st.warning(f"Sudah diproses: {row_data[14]}")
    except: pass

# --- 6. TAMPILAN INPUT USER ---
elif st.session_state.logged_in:
    if st.session_state.submitted:
        d = st.session_state.data_ringkasan
        st.success("## ✅ PENGAJUAN TELAH TERKIRIM")
        st.write("---")
        st.subheader("Ringkasan Pengajuan")
        col_res1, col_res2 = st.columns(2)
        with col_res1:
            st.write(f"**1. No. Pengajuan:** {d['no_pengajuan']}")
            st.write(f"**2. Kode Store:** {d['kode_store']}")
            st.write(f"**3. Dibayarkan:** {d['nama']} / {d['nip']}")
        with col_res2:
            st.write(f"**4. Nominal:** Rp {int(d['nominal']):,}")
            st.write(f"**5. Terbilang:** {d['terbilang']}")
            st.write(f"**6. Janji:** {d['janji']}")
        
        if st.button("Buat Pengajuan Baru"):
            st.session_state.submitted = False
            st.rerun()

    else:
        st.subheader("📍 Identifikasi Lokasi")
        kode_store = st.text_input("Masukkan Kode Store", placeholder="Contoh: A644").upper()

        if kode_store != "":
            try:
                creds = get_creds()
                client = gspread.authorize(creds)
                db_sheet = client.open_by_key(SPREADSHEET_ID).worksheet("DATABASE_USER")
                user_records = db_sheet.get_all_records()
                
                store_info = next((u for u in user_records if str(u['Kode_Store']) == kode_store), None)
                if not store_info:
                    st.error("⚠️ Kode store tidak ada atau belum terdaftar")
                    st.stop()
                
                nama_store_display = store_info.get('Nama_Store', kode_store)
                managers_db = [f"{u['NIK']} - {u['Nama Lengkap']} ({u['Email']})" for u in user_records if str(u['Kode_Store']) == kode_store and u['Role'] == 'Manager']
                cashiers_db = [f"{u['NIK']} - {u['Nama Lengkap']} ({u['Email']})" for u in user_records if str(u['Kode_Store']) == kode_store and u['Role'] == 'Senior Cashier']
                manager_email_map = {f"{u['NIK']} - {u['Nama Lengkap']} ({u['Email']})": u['Email'] for u in user_records if str(u['Kode_Store']) == kode_store and u['Role'] == 'Manager'}
                
                if not managers_db or not cashiers_db:
                    st.error(f"⚠️ Data Manager/Cashier untuk {nama_store_display} tidak lengkap.")
                    st.stop()
            except:
                st.error("⚠️ Kode store tidak ada atau belum terdaftar")
                st.stop()

            tgl_obj = datetime.datetime.now(WIB)
            st.markdown(f'<span class="store-header">Unit Bisnis Store: {nama_store_display}</span>', unsafe_allow_html=True)
            
            st.markdown('<div class="label-container"><span class="label-text">Email Request</span></div>', unsafe_allow_html=True)
            email_req = st.text_input("", value=SENDER_EMAIL, disabled=True)
            
            nama_penerima = st.text_input("Dibayarkan Kepada", key="nama_val")
            nip = st.text_input("NIP (Wajib 6 Digit)", max_chars=6, key="nip_val")
            dept = st.selectbox("Departemen", ["-", "Operational", "Sales", "Inventory", "HR", "Other"], key="dept_val")
            nominal_raw = st.text_input("Nominal", key="nom_val")
            if nominal_raw and nominal_raw.isdigit():
                st.caption(f"**Terbilang:** {terbilang(int(nominal_raw))} Rupiah")

            keperluan = st.text_input("Keperluan", key="kep_val")
            opsi = st.radio("Metode Lampiran:", ["Upload File", "Kamera"])
            bukti_file = st.file_uploader("Pilih file") if opsi == "Upload File" else st.camera_input("Ambil Foto")
            janji_tgl = st.date_input("Janji Penyelesaian", min_value=datetime.date.today(), key="janji_val")
            senior_cashier = st.selectbox("Senior Cashier Incharge", ["-"] + cashiers_db, key="sc_val")
            mgr_name_full = st.selectbox("Manager Incharge", ["-"] + managers_db, key="mgr_val")

            if st.button("Kirim Pengajuan", type="primary"):
                if nama_penerima and len(nip)==6 and nominal_raw.isdigit() and dept!="-" and senior_cashier!="-" and mgr_name_full!="-":
                    try:
                        with st.spinner("Processing..."):
                            creds = get_creds()
                            client = gspread.authorize(creds)
                            sheet = client.open_by_key(SPREADSHEET_ID).worksheet("DATA_KASBON_AZKO")
                            
                            # LOGIKA NOMOR KASBON (TANPA DD, RESET HARIAN)
                            tgl_hari_ini = tgl_obj.strftime("%d/%m/%Y")
                            prefix = f"KB{kode_store}-{tgl_obj.strftime('%m%y')}-"
                            all_rows = sheet.get_all_values()
                            last_num = 0
                            for row in all_rows:
                                if len(row) > 1 and row[0].startswith(tgl_obj.strftime("%Y-%m-%d")): # Cek tgl yg sama di Timestamp
                                    num_part = row[1].split("-")[-1]
                                    try: last_num = max(last_num, int(num_part[-3:])) # Ambil 3 digit terakhir
                                    except: pass
                            
                            # ID Unik Database (Tetap pakai Day di dalam urutan agar link unik)
                            no_pengajuan = f"{prefix}{tgl_obj.strftime('%d')}{str(last_num + 1).zfill(3)}"
                            
                            final_terbilang = terbilang(int(nominal_raw)) + " Rupiah"
                            data_final = [
                                datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S"),
                                no_pengajuan, kode_store, email_req, nama_penerima, nip, dept, 
                                nominal_raw, final_terbilang, keperluan, "Terlampir", 
                                janji_tgl.strftime("%d/%m/%Y"), senior_cashier, mgr_name_full, "Pending"
                            ]
                            sheet.append_row(data_final)

                            # Email Manager
                            app_link = f"{BASE_URL}?id={no_pengajuan}"
                            send_email_with_attachment(manager_email_map[mgr_name_full], f"Kasbon {no_pengajuan}", f"Mohon Approval: <a href='{app_link}'>Link</a>", bukti_file)
                        
                        st.session_state.data_ringkasan = {'no_pengajuan': no_pengajuan, 'kode_store': kode_store, 'nama': nama_penerima, 'nip': nip, 'nominal': nominal_raw, 'terbilang': final_terbilang, 'janji': janji_tgl.strftime("%d/%m/%Y")}
                        st.session_state.submitted = True
                        st.rerun()
                    except Exception as e: st.error(f"Sistem Sibuk: {e}")
                else: st.error("Lengkapi data!")