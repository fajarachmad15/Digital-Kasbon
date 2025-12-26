import streamlit as st
import datetime
import gspread
import io
import smtplib
import requests  # LIBRARY WAJIB
import base64    # LIBRARY WAJIB
import time      # LIBRARY WAJIB UNTUK DELAY
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formataddr
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

    /* Styling Tombol Transparan & Simbol Putih */
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        color: white !important;
    }
    div.stColumn:nth-of-type(1) [data-testid="stButton"] button {
        background-color: rgba(40, 167, 69, 0.3) !important;
        border: 1px solid rgba(40, 167, 69, 0.5) !important;
    }
    div.stColumn:nth-of-type(2) [data-testid="stButton"] button {
        background-color: rgba(220, 53, 69, 0.3) !important;
        border: 1px solid rgba(220, 53, 69, 0.5) !important;
    }

    /* HILANGKAN TOMBOL +/- PADA NUMBER INPUT */
    [data-testid="stNumberInput"] button {
        display: none !important;
    }
    /* Sembunyikan container step jika ada */
    div[data-testid="stInputNumberStepContainer"] {
        display: none !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- LOGIKA LOGIN GOOGLE (WAJIB) ---
if not st.experimental_user.is_logged_in:
    st.markdown("## 🌐 Kasbon Digital Petty Cash")
    st.info("Silakan login menggunakan akun Google Anda untuk melanjutkan.")
    if st.button("Sign in with Google", type="primary", use_container_width=True):
        st.login()
    st.stop()

# Simpan email google yang sedang login
pic_email = st.experimental_user.email

# --- 2. KONFIGURASI ---
SENDER_EMAIL = "achmad.setiawan@kawanlamacorp.com"
APP_PASSWORD = st.secrets["APP_PASSWORD"] 
BASE_URL = "https://digital-kasbon-ahi.streamlit.app" 
SPREADSHEET_ID = "1TGsCKhBC0E0hup6RGVbGrpB6ds5Jdrp5tNlfrBORzaI"

# --- CONFIG BARU UNTUK UPLOAD ---
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbylCNQsYQCIvO2qWtEUIq7gPufCgx4U5sbPasGVMGTIbaZhRFZBnpcMiHMlB2CpsEpj/exec" 
DRIVE_FOLDER_ID = "1H6aZbRbJ7Kw7zdTqkIED1tQUrBR43dBr"
# -------------------------------------------------------------------------------

WIB = datetime.timezone(datetime.timedelta(hours=7))

def get_creds():
    creds_dict = st.secrets["gcp_service_account"]
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    return ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)

def send_email_with_attachment(to_email, subject, message_body):
    try:
        msg = MIMEText(message_body, 'html')
        msg['Subject'] = subject
        msg['To'] = to_email
        msg['From'] = formataddr(("Bot_KasbonPC_Digital <No-Reply>", SENDER_EMAIL))
        
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
if 'submitted' not in st.session_state: st.session_state.submitted = False
if 'data_ringkasan' not in st.session_state: st.session_state.data_ringkasan = {}
if 'show_errors' not in st.session_state: st.session_state.show_errors = False
if 'mgr_logged_in' not in st.session_state: st.session_state.mgr_logged_in = False
if 'cashier_real_logged_in' not in st.session_state: st.session_state.cashier_real_logged_in = False
if 'mgr_final_logged_in' not in st.session_state: st.session_state.mgr_final_logged_in = False
if 'user_role' not in st.session_state: st.session_state.user_role = ""
if 'user_nik' not in st.session_state: st.session_state.user_nik = ""
if 'user_store_code' not in st.session_state: st.session_state.user_store_code = ""
if 'portal_verified' not in st.session_state: st.session_state.portal_verified = False

# --- 4. TAMPILAN PORTAL (Unified Routing) ---
query_id = st.query_params.get("id")

if query_id:
    try:
        creds = get_creds()
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID).worksheet("DATA_KASBON_AZKO")
        cell = sheet.find(query_id)
        row_data = sheet.row_values(cell.row)
        
        # Mapping Data Dasar
        r_store_code = row_data[2]
        r_no = row_data[1]
        r_req_email = row_data[3]
        r_nama = row_data[4]
        r_nip = str(row_data[5])
        r_dept = row_data[6]
        r_nominal_awal = int(row_data[7])
        r_terbilang_awal = row_data[8]
        r_keperluan = row_data[9]
        r_link_lampiran = row_data[10]
        r_janji = row_data[11]
        
        # --- DATA PENUGASAN (Assignment) ---
        # Format string di GSheet: "NIK - Nama (Email)"
        assigned_cashier_str = row_data[12] # Kolom M (Index 12)
        assigned_manager_str = row_data[13] # Kolom N (Index 13)
        
        # Parse NIK dan Email Petugas
        try:
            target_cashier_nik = assigned_cashier_str.split(" - ")[0].strip()
            target_cashier_email = assigned_cashier_str.split("(")[1].split(")")[0].strip()
            
            target_manager_nik = assigned_manager_str.split(" - ")[0].strip()
            target_manager_email = assigned_manager_str.split("(")[1].split(")")[0].strip()
        except:
            target_cashier_nik = ""; target_cashier_email = ""
            target_manager_nik = ""; target_manager_email = ""

        # Approval MGR: O=14, P=15, Q=16, R=17
        status_mgr = row_data[16] if len(row_data) > 16 else "Pending"
        reason_mgr = row_data[17] if len(row_data) > 17 else ""

        # Verifikasi Cashier: S=18, T=19, U=20, V=21
        status_cashier = row_data[20] if len(row_data) > 20 else "Pending"
        reason_csr = row_data[21] if len(row_data) > 21 else ""

        # Uang Diterima: W=22, X=23, Y=24
        status_terima = row_data[24] if len(row_data) > 24 else "Pending"

        # Realisasi: Z=25 ... AI=34
        status_real = row_data[34] if len(row_data) > 34 else "Pending"
        link_real_lampiran = row_data[33] if len(row_data) > 33 else "-"
        
        # Verifikasi Realisasi Cashier: AJ=35 s/d AQ=42
        status_verif_real = row_data[41] if len(row_data) > 41 else "Pending" # AP=41
        
        # Final Cek Manager: AR=43
        final_cek_timestamp = row_data[43] if len(row_data) > 43 else ""

        # =========================================================================
        #                               ROUTING SYSTEM
        # =========================================================================

        # --- KONDISI 1: APPROVAL MANAGER ---
        if status_mgr == "Pending":
            judul_portal = "Portal Approval Manager"
            display_status = "Status Kasbon: Waiting Manager Approval"
            
            st.markdown(f'<span class="store-header">{judul_portal}</span>', unsafe_allow_html=True)
            
            # User Recognition & Task Completion Check
            if pic_email == target_manager_email and status_mgr != "Pending":
                 st.info(f"ℹ️ Anda telah menyelesaikan bagian Anda untuk pengajuan ini. Status saat ini: {status_mgr}")
                 st.stop()

            # LOGIN MGR
            if not st.session_state.mgr_logged_in:
                st.subheader("🔐 Verifikasi Manager")
                # Hide NIK in caption for privacy
                mgr_clean_name = assigned_manager_str.split(" - ")[1] if " - " in assigned_manager_str else assigned_manager_str
                st.caption(f"Verifikasi untuk: {mgr_clean_name}")
                
                v_nik = st.text_input("NIK (6 Digit)", max_chars=6)
                v_pass = st.text_input("Password", type="password")
                
                if st.button("Masuk & Verifikasi", type="primary", use_container_width=True):
                    # STRICT LOGIN VALIDATION
                    if v_nik != target_manager_nik:
                        st.error("⛔ Anda tidak berwenang memproses pengajuan ini (NIK tidak sesuai penugasan).")
                        st.stop()
                        
                    if len(v_nik) == 6 and len(v_pass) >= 6:
                        records = client.open_by_key(SPREADSHEET_ID).worksheet("DATABASE_USER").get_all_records()
                        user = next((r for r in records if str(r['NIK']).zfill(6) == v_nik and str(r['Password']) == v_pass), None)
                        if user: 
                            st.session_state.mgr_logged_in = True
                            st.session_state.user_role = user['Role']
                            st.session_state.user_nik = str(user['NIK']).zfill(6)
                            st.session_state.user_store_code = str(user['Kode_Store'])
                            st.rerun()
                        else: st.error("NIK atau Password salah.")
                    else: st.warning("Cek kembali NIK & Password.")
                st.stop()
            
            store_pengajuan = row_data[2]
            if st.session_state.user_store_code != store_pengajuan:
                st.error("⛔ AKSES DITOLAK! Store tidak sesuai.")
                st.stop()

            # Tampilan Data (Bullet Point Standard)
            st.info(f"### Rincian Pengajuan")
            st.markdown(f"""
            * **Nomor Pengajuan Kasbon** : {query_id}
            * **Tgl dan Jam Pengajuan** : {row_data[0]}
            * **Dibayarkan Kepada** : {r_nama} / {r_nip}
            * **Departement** : {r_dept}
            * **Senilai** : Rp {r_nominal_awal:,} ({r_terbilang_awal})
            * **Untuk Keperluan** : {r_keperluan}
            * **Approval Pendukung** : [{r_link_lampiran}]({r_link_lampiran})
            * **Janji Penyelesaian** : {r_janji}
            """)
            st.write(f"**Status Saat Ini:** `{display_status}`")
            st.divider()

            if st.session_state.user_role == "Manager":
                alasan = st.text_area("Alasan Reject (Wajib diisi jika Reject)", placeholder="Contoh: Nominal terlalu besar...")
                b1, b2 = st.columns(2)
                if b1.button("✓ APPROVE", use_container_width=True):
                    tgl = datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S")
                    sheet.update_cell(cell.row, 15, tgl)
                    sheet.update_cell(cell.row, 16, st.session_state.user_nik)
                    sheet.update_cell(cell.row, 17, "APPROVED")
                    sheet.update_cell(cell.row, 18, "-")
                    
                    try:
                        cashier_name = assigned_cashier_str.split(" - ")[1].split(" (")[0]
                        link_verif_kasbon = f"{BASE_URL}?id={query_id}"
                        
                        # Email HTML Professional - Strict Wording
                        email_msg = f"""
                        <html><body style='font-family: Arial, sans-serif; font-size: 14px; color: #000000;'>
                            <div style='margin-bottom: 10px;'>Dear Bapak / Ibu {cashier_name}</div>
                            <div style='margin-bottom: 10px;'>Pengajuan kasbon dengan data dibawah ini telah di-<b>APPROVED</b> oleh Manager:</div>
                            <table style='border: none; border-collapse: collapse; width: 100%; max-width: 600px;'>
                                <tr><td style='width: 200px; padding: 2px 0;'>Nomor Pengajuan Kasbon</td><td>: {query_id}</td></tr>
                                <tr><td style='padding: 2px 0;'>Tgl dan Jam Pengajuan</td><td>: {row_data[0]}</td></tr>
                                <tr><td style='padding: 2px 0;'>Dibayarkan Kepada</td><td>: {r_nama} / {r_nip}</td></tr>
                                <tr><td style='padding: 2px 0;'>Departement</td><td>: {r_dept}</td></tr>
                                <tr><td style='padding: 2px 0;'>Senilai</td><td>: Rp {r_nominal_awal:,} ({r_terbilang_awal})</td></tr>
                                <tr><td style='padding: 2px 0;'>Untuk Keperluan</td><td>: {r_keperluan}</td></tr>
                                <tr><td style='padding: 2px 0;'>Approval Pendukung</td><td>: <a href="{r_link_lampiran}">{r_link_lampiran}</a></td></tr>
                                <tr><td style='padding: 2px 0;'>Janji Penyelesaian</td><td>: {r_janji}</td></tr>
                            </table>
                            <div style='margin-top: 15px; margin-bottom: 5px;'>
                                Silahkan klik <a href='{link_verif_kasbon}' style='text-decoration: none; color: #0000EE; font-weight: bold;'>Link Verifikasi Kasbon</a> untuk melanjutkan prosesnya.
                            </div>
                            <div style='margin-bottom: 10px;'>
                                Kemudian klik <a href='{link_verif_kasbon}' style='text-decoration: none; color: #0000EE; font-weight: bold;'>Link Verifikasi Realisasi Kasbon</a> (link yang sama) setelah pemohon melakukan realisasi.
                            </div>
                            <div>Terima Kasih</div>
                        </body></html>
                        """
                        send_email_with_attachment(target_cashier_email, f"Verifikasi Kasbon {query_id}", email_msg)
                    except: pass
                    
                    st.success("✅ Berhasil! Tugas Anda untuk ID Kasbon ini telah selesai.")
                    st.balloons()
                    time.sleep(2)
                    st.rerun() # Refresh to hit User Recognition block

                if b2.button("✕ REJECT", use_container_width=True):
                    if not alasan: st.error("Harap isi alasan reject!"); st.stop()
                    tgl = datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S")
                    sheet.update_cell(cell.row, 15, tgl)
                    sheet.update_cell(cell.row, 16, st.session_state.user_nik)
                    sheet.update_cell(cell.row, 17, "REJECTED")
                    sheet.update_cell(cell.row, 18, alasan)
                    st.error("Pengajuan telah di-Reject.")
                    time.sleep(2)
                    st.rerun()
            else:
                st.info("Anda bukan Manager untuk pengajuan ini.")


        # --- KONDISI 2: VERIFIKASI CASHIER ---
        elif status_mgr == "APPROVED" and status_cashier == "Pending":
            judul_portal = "Portal Verifikasi Cashier"
            display_status = "Status Kasbon: Waiting Cashier Verification"
            
            st.markdown(f'<span class="store-header">{judul_portal}</span>', unsafe_allow_html=True)

            # User Recognition
            if pic_email == target_cashier_email and status_cashier != "Pending":
                st.info(f"ℹ️ Anda telah menyelesaikan bagian Anda untuk pengajuan ini. Status saat ini: {status_cashier}")
                st.stop()

            # LOGIN CSR
            if not st.session_state.mgr_logged_in:
                st.subheader("🔐 Verifikasi Cashier")
                # Hide NIK in caption
                csr_clean_name = assigned_cashier_str.split(" - ")[1] if " - " in assigned_cashier_str else assigned_cashier_str
                st.caption(f"Verifikasi untuk: {csr_clean_name}")
                
                v_nik = st.text_input("NIK (6 Digit)", max_chars=6)
                v_pass = st.text_input("Password", type="password")
                
                if st.button("Masuk & Verifikasi", type="primary", use_container_width=True):
                    # STRICT LOGIN VALIDATION
                    if v_nik != target_cashier_nik:
                        st.error("⛔ Anda tidak berwenang memproses pengajuan ini (NIK tidak sesuai penugasan).")
                        st.stop()

                    if len(v_nik) == 6 and len(v_pass) >= 6:
                        records = client.open_by_key(SPREADSHEET_ID).worksheet("DATABASE_USER").get_all_records()
                        user = next((r for r in records if str(r['NIK']).zfill(6) == v_nik and str(r['Password']) == v_pass), None)
                        if user: 
                            st.session_state.mgr_logged_in = True
                            st.session_state.user_role = user['Role']
                            st.session_state.user_nik = str(user['NIK']).zfill(6)
                            st.session_state.user_store_code = str(user['Kode_Store'])
                            st.rerun()
                        else: st.error("NIK atau Password salah.")
                    else: st.warning("Cek kembali NIK & Password.")
                st.stop()

            store_pengajuan = row_data[2]
            if st.session_state.user_store_code != store_pengajuan:
                st.error("⛔ AKSES DITOLAK! Store tidak sesuai.")
                st.stop()

            # Tampilan Data (Bullet Point)
            st.info(f"### Rincian Pengajuan")
            st.markdown(f"""
            * **Nomor Pengajuan Kasbon** : {query_id}
            * **Tgl dan Jam Pengajuan** : {row_data[0]}
            * **Dibayarkan Kepada** : {r_nama} / {r_nip}
            * **Departement** : {r_dept}
            * **Senilai** : Rp {r_nominal_awal:,} ({r_terbilang_awal})
            * **Untuk Keperluan** : {r_keperluan}
            * **Approval Pendukung** : [{r_link_lampiran}]({r_link_lampiran})
            * **Janji Penyelesaian** : {r_janji}
            """)
            st.write(f"**Status Saat Ini:** `{display_status}`")
            st.divider()

            if st.session_state.user_role == "Senior Cashier":
                alasan_c = st.text_area("Alasan Reject (Wajib diisi jika Reject)", placeholder="Contoh: Saldo fisik tidak cukup...")
                k1, k2 = st.columns(2)
                
                if k1.button("✓ VERIFIKASI APPROVE", use_container_width=True):
                    tgl = datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S")
                    sheet.update_cell(cell.row, 19, tgl)
                    sheet.update_cell(cell.row, 20, st.session_state.user_nik)
                    sheet.update_cell(cell.row, 21, "APPROVED")
                    sheet.update_cell(cell.row, 22, "-")
                    
                    try:
                        link_portal = f"{BASE_URL}?id={query_id}"
                        # Email HTML Professional - Strict Wording
                        email_req_body = f"""
                        <html><body style='font-family: Arial, sans-serif; font-size: 14px; color: #000000;'>
                            <div style='margin-bottom: 10px;'>Dear Bapak / Ibu {r_nama}</div>
                            <div style='margin-bottom: 10px;'>Pengajuan kasbon dengan data dibawah ini telah di-<b>APPROVE</b> oleh Manager dan di-<b>VERIFIKASI</b> oleh Cashier :</div>
                            <table style='border: none; border-collapse: collapse; width: 100%; max-width: 600px;'>
                                <tr><td style='width: 200px; padding: 2px 0;'>Nomor Pengajuan Kasbon</td><td>: {query_id}</td></tr>
                                <tr><td style='padding: 2px 0;'>Tgl dan Jam Pengajuan</td><td>: {row_data[0]}</td></tr>
                                <tr><td style='padding: 2px 0;'>Dibayarkan Kepada</td><td>: {r_nama} / {r_nip}</td></tr>
                                <tr><td style='padding: 2px 0;'>Departement</td><td>: {r_dept}</td></tr>
                                <tr><td style='padding: 2px 0;'>Senilai</td><td>: Rp {r_nominal_awal:,} ({r_terbilang_awal})</td></tr>
                                <tr><td style='padding: 2px 0;'>Untuk Keperluan</td><td>: {r_keperluan}</td></tr>
                                <tr><td style='padding: 2px 0;'>Approval Pendukung</td><td>: <a href="{r_link_lampiran}">{r_link_lampiran}</a></td></tr>
                                <tr><td style='padding: 2px 0;'>Janji Penyelesaian</td><td>: {r_janji}</td></tr>
                            </table>
                            <div style='margin-top: 20px; margin-bottom: 5px;'>
                                Klik <a href='{link_portal}' style='text-decoration: none; color: #0000EE; font-weight:bold;'>Link Diterima</a> sebagai konfirmasi uang telah diterima.
                            </div>
                            <div style='margin-bottom: 20px;'>
                                Dan Klik <a href='{link_portal}' style='text-decoration: none; color: #0000EE; font-weight:bold;'>Link Realisasi</a> ketika uang sudah selesai digunakan.
                            </div>
                            <div>Terima Kasih</div>
                        </body></html>
                        """
                        send_email_with_attachment(r_req_email, f"Kasbon Disetujui {query_id}", email_req_body)
                    except: pass
                    
                    st.success("✅ Berhasil! Tugas Anda untuk ID Kasbon ini telah selesai.")
                    st.balloons()
                    time.sleep(2)
                    st.rerun()
                
                if k2.button("✕ VERIFIKASI REJECT", use_container_width=True):
                    if not alasan_c: st.error("Harap isi alasan reject!"); st.stop()
                    tgl = datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S")
                    sheet.update_cell(cell.row, 19, tgl)
                    sheet.update_cell(cell.row, 20, st.session_state.user_nik)
                    sheet.update_cell(cell.row, 21, "REJECTED")
                    sheet.update_cell(cell.row, 22, alasan_c)
                    st.error("Verifikasi Ditolak.")
                    time.sleep(2)
                    st.rerun()
            else:
                st.info("Anda bukan Senior Cashier untuk pengajuan ini.")


        # --- KONDISI 3: KONFIRMASI UANG DITERIMA ---
        elif status_cashier == "APPROVED" and status_terima == "Pending":
            judul_portal = "Portal Konfirmasi Uang Diterima"
            display_status = "Status Kasbon: Waiting Requester Confirmation"
            
            st.markdown(f'<span class="store-header">{judul_portal}</span>', unsafe_allow_html=True)
            
            # User Recognition
            # Note: Pemohon biasanya pakai email pic_email, kalau dia sudah klik terima, status_terima != pending, jadi blok ini auto skip.
            
            # LOGIN NIP + DOUBLE AUTH
            if not st.session_state.portal_verified:
                st.info("🔒 Untuk keamanan, masukkan NIP Anda dan Password (6 karakter awal email login).")
                c_a, c_b = st.columns(2)
                nip_input = c_a.text_input("NIP Pemohon", max_chars=6)
                # Requester Password Logic
                pass_input = c_b.text_input("Password (6 char awal email)", type="password", max_chars=6)
                
                if st.button("Masuk Portal"):
                    correct_password = pic_email[:6]
                    if nip_input == r_nip and pass_input == correct_password:
                        st.session_state.portal_verified = True
                        st.rerun()
                    else: st.error("⛔ Validasi Gagal! Pastikan NIP benar dan Password adalah 6 huruf pertama email Anda.")
                st.stop()

            # Tampilan Data
            st.info(f"### Rincian Pengajuan")
            st.markdown(f"""
            * **Nomor Pengajuan Kasbon** : {query_id}
            * **Tgl dan Jam Pengajuan** : {row_data[0]}
            * **Dibayarkan Kepada** : {r_nama} / {r_nip}
            * **Departement** : {r_dept}
            * **Senilai** : Rp {r_nominal_awal:,} ({r_terbilang_awal})
            * **Untuk Keperluan** : {r_keperluan}
            * **Approval Pendukung** : [{r_link_lampiran}]({r_link_lampiran})
            * **Janji Penyelesaian** : {r_janji}
            """)
            st.write(f"**Status Saat Ini:** `{display_status}`")
            st.divider()

            st.write("Silakan klik tombol di bawah jika uang kasbon fisik telah Anda terima.")
            if st.button("Konfirmasi uang sudah diterima dan sesuai", type="primary", use_container_width=True):
                tgl_terima = datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S")
                sheet.update_cell(cell.row, 23, tgl_terima) 
                sheet.update_cell(cell.row, 24, r_nip)      
                sheet.update_cell(cell.row, 25, "Sudah diterima") 
                st.success("✅ Berhasil! Tugas Anda untuk ID Kasbon ini telah selesai.")
                st.balloons()
                time.sleep(2)
                st.rerun()


        # --- KONDISI 4: INPUT REALISASI ---
        elif status_terima == "Sudah diterima" and status_real == "Pending":
            judul_portal = "Portal Realisasi Kasbon"
            display_status = "Status Kasbon: Waiting Realization Input"
            
            st.markdown(f'<span class="store-header">{judul_portal}</span>', unsafe_allow_html=True)
            
            # LOGIN NIP + DOUBLE AUTH
            if not st.session_state.portal_verified:
                st.info("🔒 Untuk keamanan, masukkan NIP Anda dan Password (6 karakter awal email login).")
                c_a, c_b = st.columns(2)
                nip_input = c_a.text_input("NIP Pemohon", max_chars=6)
                # Requester Password Logic
                pass_input = c_b.text_input("Password (6 char awal email)", type="password", max_chars=6)
                
                if st.button("Masuk Portal"):
                    correct_password = pic_email[:6]
                    if nip_input == r_nip and pass_input == correct_password:
                        st.session_state.portal_verified = True
                        st.rerun()
                    else: st.error("⛔ Validasi Gagal! Pastikan NIP benar dan Password adalah 6 huruf pertama email Anda.")
                st.stop()

            # Tampilan Data
            st.info(f"### Rincian Pengajuan")
            st.markdown(f"""
            * **Nomor Pengajuan Kasbon** : {query_id}
            * **Tgl dan Jam Pengajuan** : {row_data[0]}
            * **Dibayarkan Kepada** : {r_nama} / {r_nip}
            * **Departement** : {r_dept}
            * **Senilai** : Rp {r_nominal_awal:,} ({r_terbilang_awal})
            * **Untuk Keperluan** : {r_keperluan}
            * **Approval Pendukung** : [{r_link_lampiran}]({r_link_lampiran})
            * **Janji Penyelesaian** : {r_janji}
            """)
            st.write(f"**Status Saat Ini:** `{display_status}`")
            st.divider()

            st.subheader("📝 Laporan Pertanggung Jawaban")
            st.markdown("**Total Uang Digunakan (Rp)**")
            uang_digunakan = st.number_input("", min_value=0, step=1, label_visibility="collapsed")
            
            terbilang_guna = ""
            if uang_digunakan > 0:
                terbilang_guna = terbilang(uang_digunakan).title() + " Rupiah"
                st.caption(f"*{terbilang_guna}*")
            else:
                st.caption("*Nol Rupiah*")

            selisih = r_nominal_awal - uang_digunakan
            
            col_kiri, col_kanan = st.columns(2)
            val_kembali = 0
            val_terima = 0
            txt_kembali = "Nol Rupiah"
            txt_terima = "Nol Rupiah"

            if uang_digunakan < r_nominal_awal:
                val_kembali = selisih
                txt_kembali = terbilang(val_kembali).title() + " Rupiah"
            elif uang_digunakan > r_nominal_awal:
                val_terima = abs(selisih)
                txt_terima = terbilang(val_terima).title() + " Rupiah"

            with col_kiri:
                st.markdown(f"**Uang yg dikembalikan ke perusahaan:**")
                st.text_input("Nominal Kembali", value=f"Rp {val_kembali:,}", disabled=True)
                st.caption(f"*{txt_kembali}*")
            
            with col_kanan:
                st.markdown(f"**Uang yg diterima (Reimburse):**")
                st.text_input("Nominal Terima", value=f"Rp {val_terima:,}", disabled=True)
                st.caption(f"*{txt_terima}*")
            
            st.write("---")
            st.write("📸 **Bukti Lampiran (Wajib: Foto Nota & Item)**")
            opsi_b = st.radio("Metode Lampiran:", ["Upload File", "Kamera"])
            bukti_real = st.file_uploader("Upload Foto", type=['png','jpg','jpeg','pdf']) if opsi_b == "Upload File" else st.camera_input("Ambil Foto")

            if st.button("Kirim Laporan Realisasi", type="primary", use_container_width=True):
                if bukti_real and uang_digunakan >= 0:
                    try:
                        link_bukti = "Lampiran Ada (File/Foto)"
                        if bukti_real:
                            try:
                                f_type = bukti_real.type
                                f_ext = f_type.split("/")[-1]
                                f_name = f"Lampiran_Realisasi_{query_id}.{f_ext}"
                                f_content = bukti_real.getvalue()
                                f_b64 = base64.b64encode(f_content).decode('utf-8')
                                pl = {"filename": f_name, "filedata": f_b64, "mimetype": f_type, "folderId": DRIVE_FOLDER_ID}
                                with st.spinner("Mengupload Bukti Realisasi..."):
                                    res = requests.post(APPS_SCRIPT_URL, json=pl)
                                    rj = res.json()
                                    if rj.get("status") == "success":
                                        link_bukti = rj.get("url")
                                    else: st.warning(f"Gagal upload ke Drive: {rj.get('message')}")
                            except Exception as e: st.warning(f"Error upload drive: {e}")
                        
                        tgl_real = datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S")
                        sheet.update_cell(cell.row, 26, tgl_real)
                        sheet.update_cell(cell.row, 27, r_nip)
                        sheet.update_cell(cell.row, 28, uang_digunakan)
                        sheet.update_cell(cell.row, 29, terbilang_guna)
                        sheet.update_cell(cell.row, 30, val_kembali)
                        sheet.update_cell(cell.row, 31, txt_kembali)
                        sheet.update_cell(cell.row, 32, val_terima)
                        sheet.update_cell(cell.row, 33, txt_terima)
                        sheet.update_cell(cell.row, 34, link_bukti)
                        sheet.update_cell(cell.row, 35, "Terrealisasi")
                        st.success("✅ Berhasil! Tugas Anda untuk ID Kasbon ini telah selesai.")
                        st.balloons()
                        time.sleep(2)
                        st.rerun()
                    except Exception as e: st.error(f"Gagal menyimpan: {e}")
                else: st.error("⚠️ Harap lengkapi bukti lampiran.")


        # --- KONDISI 5: VERIFIKASI REALISASI CASHIER ---
        elif status_real == "Terrealisasi" and status_verif_real == "Pending":
            judul_portal = "Portal Verifikasi Realisasi Kasbon"
            display_status = "Status Kasbon: Waiting Cashier Verification (Realisasi)"
            
            st.markdown(f'<span class="store-header">{judul_portal}</span>', unsafe_allow_html=True)
            
            # User Recognition
            if pic_email == target_cashier_email and status_verif_real != "Pending":
                 st.info(f"ℹ️ Anda telah menyelesaikan bagian Anda untuk pengajuan ini. Status saat ini: {status_verif_real}")
                 st.stop()

            # LOGIN CASHIER
            if not st.session_state.cashier_real_logged_in:
                st.subheader("🔐 Login Cashier")
                # Hide NIK in caption
                csr_clean_name = assigned_cashier_str.split(" - ")[1] if " - " in assigned_cashier_str else assigned_cashier_str
                st.caption(f"Verifikasi untuk: {csr_clean_name}")
                
                v_nik = st.text_input("NIK (6 Digit)", max_chars=6)
                v_pass = st.text_input("Password", type="password")
                if st.button("Masuk"):
                    # STRICT LOGIN VALIDATION
                    if v_nik != target_cashier_nik:
                        st.error("⛔ Anda tidak berwenang memproses pengajuan ini (NIK tidak sesuai penugasan).")
                        st.stop()

                    records = client.open_by_key(SPREADSHEET_ID).worksheet("DATABASE_USER").get_all_records()
                    user = next((r for r in records if str(r['NIK']).zfill(6) == v_nik and str(r['Password']) == v_pass), None)
                    if user and (user['Role'] == 'Senior Cashier'):
                        st.session_state.cashier_real_logged_in = True
                        st.session_state.user_nik = str(user['NIK']).zfill(6)
                        st.rerun()
                    else: st.error("Login gagal atau akses ditolak (Bukan Cashier).")
                st.stop()

            # Tampilan Data
            st.info(f"### Rincian Pengajuan")
            st.markdown(f"""
            * **Nomor Pengajuan Kasbon** : {query_id}
            * **Tgl dan Jam Pengajuan** : {row_data[0]}
            * **Dibayarkan Kepada** : {r_nama} / {r_nip}
            * **Departement** : {r_dept}
            * **Senilai** : Rp {r_nominal_awal:,} ({r_terbilang_awal})
            * **Untuk Keperluan** : {r_keperluan}
            * **Approval Pendukung** : [{r_link_lampiran}]({r_link_lampiran})
            * **Lampiran Realisasi** : [{link_real_lampiran}]({link_real_lampiran})
            * **Janji Penyelesaian** : {r_janji}
            """)
            st.write(f"**Status Saat Ini:** `{display_status}`")
            st.divider()

            st.markdown("### Verifikasi Data")
            status_pilihan = st.radio("Apakah status realisasi sesuai?", ["Ya, Sesuai", "Tidak Sesuai"])
            
            # 5. Reactive Terbilang Logic
            is_disabled = True if status_pilihan == "Ya, Sesuai" else False
            
            reason_verif = "-"
            if status_pilihan == "Tidak Sesuai":
                reason_verif = st.text_area("Reason (Wajib diisi)", placeholder="Jelaskan alasan ketidaksesuaian...")

            # Ambil data read-only awal
            u_kembali_db = int(row_data[29]) if len(row_data) > 29 and row_data[29] else 0
            u_terima_db = int(row_data[31]) if len(row_data) > 31 and row_data[31] else 0
            
            c_al, c_am = st.columns(2)
            with c_al:
                # Reactive: When user types, streamlit reruns, and st.caption below updates immediately
                u_kembali_input = st.number_input("Uang Dikembalikan", value=u_kembali_db, disabled=is_disabled, step=1)
                st.caption(terbilang(u_kembali_input).title() + " Rupiah")
                
            with c_am:
                u_terima_input = st.number_input("Uang Diterima", value=u_terima_db, disabled=is_disabled, step=1)
                st.caption(terbilang(u_terima_input).title() + " Rupiah")
            
            if st.button("Submit Verifikasi", type="primary"):
                if status_pilihan == "Tidak Sesuai" and not reason_verif:
                    st.error("Harap isi reason jika tidak sesuai.")
                    st.stop()
                
                final_u_kembali = u_kembali_input
                final_u_terima = u_terima_input
                
                txt_kembali_final = terbilang(final_u_kembali).title() + " Rupiah"
                txt_terima_final = terbilang(final_u_terima).title() + " Rupiah"
                
                tgl_verif = datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S")
                
                sheet.update_cell(cell.row, 36, tgl_verif)                # AJ
                sheet.update_cell(cell.row, 37, st.session_state.user_nik) # AK
                sheet.update_cell(cell.row, 38, final_u_kembali)          # AL
                sheet.update_cell(cell.row, 39, txt_kembali_final)        # AM 
                sheet.update_cell(cell.row, 40, final_u_terima)           # AN
                sheet.update_cell(cell.row, 41, txt_terima_final)         # AO
                sheet.update_cell(cell.row, 42, status_pilihan)           # AP
                sheet.update_cell(cell.row, 43, reason_verif)             # AQ

                try:
                    mgr_email = target_manager_email if target_manager_email else SENDER_EMAIL
                    mgr_nama = assigned_manager_str.split(" - ")[1].split(" (")[0]
                    
                    link_final = f"{BASE_URL}?id={query_id}"
                    
                    email_mgr_body = f"""
                    <html><body style='font-family: Arial, sans-serif; font-size: 14px; color: #000000;'>
                        <div style='margin-bottom: 10px;'>Dear Bapak / Ibu {mgr_nama}</div>
                        <div style='margin-bottom: 10px;'>Mohon melakukan Final Cek untuk Laporan realisasi kasbon dengan data di bawah ini :</div>
                        <table style='border: none; border-collapse: collapse; width: 100%; max-width: 600px;'>
                            <tr><td style='width: 200px; padding: 2px 0;'>Nomor Pengajuan Kasbon</td><td>: {query_id}</td></tr>
                            <tr><td style='padding: 2px 0;'>Tgl dan Jam Pengajuan</td><td>: {row_data[0]}</td></tr>
                            <tr><td style='padding: 2px 0;'>Dibayarkan Kepada</td><td>: {r_nama} / {r_nip}</td></tr>
                            <tr><td style='padding: 2px 0;'>Departement</td><td>: {r_dept}</td></tr>
                            <tr><td style='padding: 2px 0;'>Senilai</td><td>: Rp {r_nominal_awal:,} ({r_terbilang_awal})</td></tr>
                            <tr><td style='padding: 2px 0;'>Untuk Keperluan</td><td>: {r_keperluan}</td></tr>
                            <tr><td style='padding: 2px 0;'>Approval Pendukung</td><td>: <a href="{r_link_lampiran}">{r_link_lampiran}</a></td></tr>
                            <tr><td style='padding: 2px 0;'>Lampiran Realisasi</td><td>: <a href="{link_real_lampiran}">{link_real_lampiran}</a></td></tr>
                        </table>
                        <div style='margin-top: 15px; margin-bottom: 10px;'>
                            Silahkan klik <a href='{link_final}' style='text-decoration: none; color: #0000EE; font-weight:bold;'>Link Final Cek</a> untuk melanjutkan prosesnya
                        </div>
                        <div>Terima Kasih</div>
                    </body></html>
                    """
                    send_email_with_attachment(mgr_email, f"Final Cek Laporan Realisasi Kasbon {query_id}", email_mgr_body)
                except: pass

                st.success("✅ Berhasil! Tugas Anda untuk ID Kasbon ini telah selesai.")
                st.balloons()
                time.sleep(2)
                st.rerun()


        # --- KONDISI 6: FINAL CEK MANAGER ---
        elif (status_verif_real == "Ya, Sesuai" or status_verif_real == "Tidak Sesuai") and final_cek_timestamp == "":
            judul_portal = "Portal Final Cek Realisasi Kasbon"
            display_status = "Status Kasbon: Waiting Manager Final Check"
            
            st.markdown(f'<span class="store-header">{judul_portal}</span>', unsafe_allow_html=True)
            
            # User Recognition
            if pic_email == target_manager_email:
                pass # Manager harus login ulang untuk final approval keamanan

            # LOGIN MANAGER
            if not st.session_state.mgr_final_logged_in:
                st.subheader("🔐 Login Manager")
                mgr_clean_name = assigned_manager_str.split(" - ")[1] if " - " in assigned_manager_str else assigned_manager_str
                st.caption(f"Verifikasi untuk: {mgr_clean_name}")
                v_nik = st.text_input("NIK (6 Digit)", max_chars=6)
                v_pass = st.text_input("Password", type="password")
                if st.button("Masuk"):
                    # STRICT LOGIN VALIDATION
                    if v_nik != target_manager_nik:
                        st.error("⛔ Anda tidak berwenang memproses pengajuan ini (NIK tidak sesuai penugasan).")
                        st.stop()
                        
                    records = client.open_by_key(SPREADSHEET_ID).worksheet("DATABASE_USER").get_all_records()
                    user = next((r for r in records if str(r['NIK']).zfill(6) == v_nik and str(r['Password']) == v_pass), None)
                    if user and (user['Role'] == 'Manager'):
                        st.session_state.mgr_final_logged_in = True
                        st.session_state.user_nik = str(user['NIK']).zfill(6)
                        st.rerun()
                    else: st.error("Login gagal atau akses ditolak (Bukan Manager).")
                st.stop()

            # Tampilan Data
            st.info(f"### Rincian Pengajuan")
            st.markdown(f"""
            * **Nomor Pengajuan Kasbon** : {query_id}
            * **Tgl dan Jam Pengajuan** : {row_data[0]}
            * **Dibayarkan Kepada** : {r_nama} / {r_nip}
            * **Departement** : {r_dept}
            * **Senilai** : Rp {r_nominal_awal:,} ({r_terbilang_awal})
            * **Untuk Keperluan** : {r_keperluan}
            * **Approval Pendukung** : [{r_link_lampiran}]({r_link_lampiran})
            * **Lampiran Realisasi** : [{link_real_lampiran}]({link_real_lampiran})
            * **Janji Penyelesaian** : {r_janji}
            """)
            st.write(f"**Status Saat Ini:** `{display_status}`")
            st.divider()

            st.markdown("**1. Apakah foto nota sesuai dengan data pembelian yg diajukan baik qty maupun nominal pembelanjaan?**")
            q1_ans = st.radio("Jawaban Q1", ["Ya, Sesuai", "Tidak Sesuai"], key="q1")
            q1_reason = "-"
            if q1_ans == "Tidak Sesuai":
                q1_reason = st.text_input("Reason Q1 (Wajib)", key="r1")
            
            st.markdown("**2. Apakah foto item sesuai dengan kebutuhan pembelian yg diajukan?**")
            q2_ans = st.radio("Jawaban Q2", ["Ya, Sesuai", "Tidak Sesuai"], key="q2")
            q2_reason = "-"
            if q2_ans == "Tidak Sesuai":
                q2_reason = st.text_input("Reason Q2 (Wajib)", key="r2")
            
            if st.button("Posting", type="primary"):
                if (q1_ans == "Tidak Sesuai" and not q1_reason) or (q2_ans == "Tidak Sesuai" and not q2_reason):
                    st.error("Harap isi reason jika memilih Tidak Sesuai.")
                    st.stop()
                
                tgl_final = datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S")
                
                # AR=44 (idx 43)
                sheet.update_cell(cell.row, 44, tgl_final)                  # AR
                sheet.update_cell(cell.row, 45, st.session_state.user_nik)  # AS
                sheet.update_cell(cell.row, 46, q1_ans)                     # AT
                sheet.update_cell(cell.row, 47, q1_reason)                  # AU
                sheet.update_cell(cell.row, 48, q2_ans)                     # AV
                sheet.update_cell(cell.row, 49, q2_reason)                  # AW
                
                # 7. Alur Status Final
                st.success("✅ Berhasil! Tugas Anda telah selesai. Status Kasbon Completed."); 
                st.balloons()
                time.sleep(2)
                st.rerun()


        # --- KONDISI 7: COMPLETED ---
        elif final_cek_timestamp != "":
            st.markdown('<span class="store-header">Portal Informasi Pengajuan</span>', unsafe_allow_html=True)
            display_status = "Status Kasbon Completed"
            
            st.info(f"### Rincian Pengajuan")
            st.markdown(f"""
            * **Nomor Pengajuan Kasbon** : {query_id}
            * **Tgl dan Jam Pengajuan** : {row_data[0]}
            * **Dibayarkan Kepada** : {r_nama} / {r_nip}
            * **Departement** : {r_dept}
            * **Senilai** : Rp {r_nominal_awal:,} ({r_terbilang_awal})
            * **Untuk Keperluan** : {r_keperluan}
            * **Approval Pendukung** : [{r_link_lampiran}]({r_link_lampiran})
            * **Lampiran Realisasi** : [{link_real_lampiran}]({link_real_lampiran})
            * **Janji Penyelesaian** : {r_janji}
            """)
            st.success(f"**{display_status}**")
            st.divider()

        # --- KONDISI REJECTED (Fallback) ---
        elif status_mgr == "REJECTED" or status_cashier == "REJECTED":
             st.error("Status Kasbon: REJECTED")
             if status_mgr == "REJECTED": st.write(f"Reason Manager: {reason_mgr}")
             if status_cashier == "REJECTED": st.write(f"Reason Cashier: {reason_csr}")

    except Exception as e: st.error(f"Error Database/System: {e}")
    st.stop()

# --- 5. TAMPILAN INPUT USER (FORMULIR PENGAJUAN) ---
if st.session_state.submitted:
    d = st.session_state.data_ringkasan
    st.success("## ✅ PENGAJUAN TELAH TERKIRIM")
    st.info(f"Email notifikasi telah dikirim ke Manager.")
    st.write("---")
    st.subheader("Ringkasan Pengajuan")
    
    # 6. Standardization Detail (Bullet Point)
    st.markdown(f"""
    * **Nomor Pengajuan Kasbon** : {d['no_pengajuan']}
    * **Tgl dan Jam Pengajuan** : {d.get('tgl_jam', '-')}
    * **Dibayarkan Kepada** : {d['nama']} / {d['nip']}
    * **Departement** : {d['dept']}
    * **Senilai** : Rp {int(d['nominal']):,} ({d['terbilang']})
    * **Untuk Keperluan** : {d['keperluan']}
    * **Approval Pendukung** : {d.get('link_pendukung', '-')}
    * **Janji Penyelesaian** : {d['janji']}
    """)
    
    c1, c2 = st.columns(2)
    if c1.button("Buat Pengajuan Baru", use_container_width=True):
        st.session_state.submitted = False; st.session_state.show_errors = False; st.rerun()
    if c2.button("Logout Google", use_container_width=True):
        st.logout()

else:
    st.caption(f"Logged in as: **{pic_email}**")
    st.subheader("📍 Identifikasi Lokasi")
    kode_store = st.text_input("Masukkan Kode Store", placeholder="Contoh: A644").upper()

    if kode_store:
        try:
            creds = get_creds()
            client = gspread.authorize(creds)
            db_sheet = client.open_by_key(SPREADSHEET_ID).worksheet("DATABASE_USER")
            user_records = db_sheet.get_all_records()
            store_info = next((u for u in user_records if str(u['Kode_Store']) == kode_store), None)
            
            if not store_info: st.error("⚠️ Kode store tidak ada atau belum terdaftar"); st.stop()
            
            st.markdown(f'<span class="store-header">Unit Bisnis Store: {store_info["Nama_Store"]}</span>', unsafe_allow_html=True)
            st.markdown('<div class="label-container"><span class="label-text">Email Request</span></div>', unsafe_allow_html=True)
            st.text_input("", value=pic_email, disabled=True)
            
            err_nama = '<span class="error-tag">Harap dilengkapi</span>' if st.session_state.show_errors and not st.session_state.get('nama_val') else ''
            st.markdown(f'<div class="label-container"><span class="label-text">Dibayarkan Kepada (Nama Lengkap)</span>{err_nama}</div>', unsafe_allow_html=True)
            nama_p = st.text_input("", key="nama_val")

            err_nip = '<span class="error-tag">Isi sesuai format</span>' if st.session_state.show_errors and (not st.session_state.get('nip_val') or len(st.session_state.get('nip_val')) != 6) else ''
            st.markdown(f'<div class="label-container"><span class="label-text">NIP (Wajib 6 Digit)</span>{err_nip}</div>', unsafe_allow_html=True)
            nip = st.text_input("", max_chars=6, key="nip_val")

            err_dept = '<span class="error-tag">Harap dipilih</span>' if st.session_state.show_errors and st.session_state.get('dept_val') == "-" else ''
            st.markdown(f'<div class="label-container"><span class="label-text">Departemen</span>{err_dept}</div>', unsafe_allow_html=True)
            dept = st.selectbox("", ["-", "Operational", "Sales", "Inventory", "HR", "Other"], key="dept_val")

            err_nom = '<span class="error-tag">Hanya angka</span>' if st.session_state.show_errors and (not st.session_state.get('nom_val') or not st.session_state.get('nom_val').isdigit()) else ''
            st.markdown(f'<div class="label-container"><span class="label-text">Nominal (Hanya Angka)</span>{err_nom}</div>', unsafe_allow_html=True)
            nom_r = st.text_input("", key="nom_val")
            if nom_r.isdigit(): 
                teks_terbilang = terbilang(int(nom_r)).title() + " Rupiah"
                st.markdown(f"<span style='font-size:1.5em; font-weight:bold;'>Terbilang: {teks_terbilang}</span>", unsafe_allow_html=True)

            err_kep = '<span class="error-tag">Harap dilengkapi</span>' if st.session_state.show_errors and not st.session_state.get('kep_val') else ''
            st.markdown(f'<div class="label-container"><span class="label-text">Untuk Keperluan</span>{err_kep}</div>', unsafe_allow_html=True)
            kep = st.text_input("", key="kep_val")

            st.write("📸 **Bukti Lampiran (Maks 5MB)**")
            opsi = st.radio("Metode Lampiran:", ["Upload File", "Kamera"])
            bukti = st.file_uploader("Pilih file") if opsi == "Upload File" else st.camera_input("Ambil Foto")
            
            st.markdown('<div class="label-container"><span class="label-text">Janji Penyelesaian</span></div>', unsafe_allow_html=True)
            janji = st.date_input("", min_value=datetime.date.today(), key="janji_val")

            managers = [f"{u['NIK']} - {u['Nama Lengkap']} ({u['Email']})" for u in user_records if str(u['Kode_Store']) == kode_store and u['Role'] == 'Manager']
            cashiers = [f"{u['NIK']} - {u['Nama Lengkap']} ({u['Email']})" for u in user_records if str(u['Kode_Store']) == kode_store and u['Role'] == 'Senior Cashier']
            mgr_map = {f"{u['NIK']} - {u['Nama Lengkap']} ({u['Email']})": u['Email'] for u in user_records if str(u['Kode_Store']) == kode_store and u['Role'] == 'Manager'}

            err_mgr = '<span class="error-tag">Pilih Manager</span>' if st.session_state.show_errors and st.session_state.get('mgr_val') == "-" else ''
            st.markdown(f'<div class="label-container"><span class="label-text">Manager Incharge</span>{err_mgr}</div>', unsafe_allow_html=True)
            mgr_f = st.selectbox("", ["-"] + managers, key="mgr_val")

            err_sc = '<span class="error-tag">Pilih Cashier</span>' if st.session_state.show_errors and st.session_state.get('sc_val') == "-" else ''
            st.markdown(f'<div class="label-container"><span class="label-text">Senior Cashier Incharge</span>{err_sc}</div>', unsafe_allow_html=True)
            sc_f = st.selectbox("", ["-"] + cashiers, key="sc_val")

            st.divider()
            
            if st.session_state.show_errors: st.error("⚠️ Lengkapi kolom bertanda merah.")

            if st.button("Kirim Pengajuan", type="primary"):
                if nama_p and len(nip)==6 and nom_r.isdigit() and kep and dept!="-" and mgr_f!="-" and sc_f!="-":
                    try:
                        with st.spinner("Processing..."):
                            sheet = client.open_by_key(SPREADSHEET_ID).worksheet("DATA_KASBON_AZKO")
                            tgl_now = datetime.datetime.now(WIB)
                            
                            prefix = f"KB{kode_store}-{tgl_now.strftime('%m%y')}-" 
                            all_rows = sheet.get_all_values()
                            last_n = 0
                            for r in all_rows:
                                if len(r)>1 and r[1].startswith(prefix):
                                    try: last_n = max(last_n, int(r[1].split("-")[-1]))
                                    except: continue
                            no_p = f"{prefix}{str(last_n + 1).zfill(3)}"
                            
                            final_t = terbilang(int(nom_r)).title() + " Rupiah"
                            
                            link_drive = "-"
                            if bukti:
                                try:
                                    file_type = bukti.type
                                    ext = file_type.split("/")[-1]
                                    file_name = f"Lampiran_Kasbon_{no_p}.{ext}"
                                    
                                    file_content = bukti.getvalue()
                                    b64_data = base64.b64encode(file_content).decode('utf-8')
                                    
                                    payload = {
                                        "filename": file_name,
                                        "filedata": b64_data,
                                        "mimetype": file_type,
                                        "folderId": DRIVE_FOLDER_ID
                                    }
                                    
                                    with st.spinner("Mengupload ke Drive..."):
                                        response = requests.post(APPS_SCRIPT_URL, json=payload)
                                        try:
                                            res_json = response.json()
                                            if res_json.get("status") == "success":
                                                link_drive = res_json.get("url")
                                            else:
                                                st.error(f"Gagal Upload: {res_json.get('message')}")
                                                st.stop()
                                        except: st.error("Error respon drive.")
                                except Exception as e:
                                    st.error(f"Error Koneksi Upload: {e}")
                                    st.stop()
                            
                            sheet.append_row([
                                tgl_now.strftime("%Y-%m-%d %H:%M:%S"), no_p, kode_store, pic_email, 
                                nama_p, nip, dept, nom_r, final_t, kep, link_drive, 
                                janji.strftime("%d/%m/%Y"), sc_f, mgr_f, 
                                "", "", "Pending", "", 
                                "", "", "Pending", "",
                                "", "", "Pending",
                                "", "", "", "", "", "", "", "", "", "Pending",
                                "", "", "", "", "", "", "Pending", "", # AJ-AQ
                                "", "", "", "", "", "" # AR-AW
                            ])
                            
                            mgr_clean = mgr_f.split(" - ")[1].split(" (")[0]
                            tgl_full = tgl_now.strftime("%d/%m/%Y %H:%M")
                            app_link = f"{BASE_URL}?id={no_p}"
                            
                            email_body = f"""
                            <html><body style='font-family: Arial, sans-serif; font-size: 14px; color: #000000;'>
                                <div style='margin-bottom: 10px;'>Dear Bapak / Ibu {mgr_clean}</div>
                                <div style='margin-bottom: 10px;'>Mohon approvalnya untuk pengajuan kasbon dengan data di bawah ini :</div>
                                <table style='border: none; border-collapse: collapse; width: 100%; max-width: 600px;'>
                                    <tr><td style='width: 200px; padding: 2px 0;'>Nomor Pengajuan Kasbon</td><td>: {no_p}</td></tr>
                                    <tr><td style='padding: 2px 0;'>Tgl dan Jam Pengajuan</td><td>: {tgl_full}</td></tr>
                                    <tr><td style='padding: 2px 0;'>Dibayarkan Kepada</td><td>: {nama_p} / {nip}</td></tr>
                                    <tr><td style='padding: 2px 0;'>Departement</td><td>: {dept}</td></tr>
                                    <tr><td style='padding: 2px 0;'>Senilai</td><td>: Rp {int(nom_r):,} ({final_t})</td></tr>
                                    <tr><td style='padding: 2px 0;'>Untuk Keperluan</td><td>: {kep}</td></tr>
                                    <tr><td style='padding: 2px 0;'>Approval Pendukung</td><td>: <a href="{link_drive}">{link_drive}</a></td></tr>
                                    <tr><td style='padding: 2px 0;'>Janji Penyelesaian</td><td>: {janji.strftime("%d/%m/%Y")}</td></tr>
                                </table>
                                <div style='margin-top: 15px; margin-bottom: 10px;'>
                                    Silahkan klik <a href='{app_link}' style='text-decoration: none; color: #0000EE; font-weight: bold;'>Link Approval</a> untuk melanjutkan prosesnya
                                </div>
                                <div>Terima Kasih</div>
                            </body></html>
                            """
                            send_email_with_attachment(mgr_map[mgr_f], f"Pengajuan Kasbon {no_p}", email_body)
                        
                            st.session_state.data_ringkasan = {
                                'no_pengajuan': no_p, 
                                'kode_store': kode_store, 
                                'nama': nama_p, 
                                'nip': nip, 
                                'dept': dept, 
                                'nominal': nom_r, 
                                'terbilang': final_t, 
                                'keperluan': kep, 
                                'janji': janji.strftime("%d/%m/%Y"),
                                'tgl_jam': tgl_full,     
                                'link_pendukung': link_drive 
                            }
                            st.session_state.submitted = True; st.session_state.show_errors = False; st.rerun()
                    except Exception as e: st.error(f"Error Sistem: {e}")
                else:
                    st.session_state.show_errors = True; st.rerun()
        except Exception as e: st.error(f"Database Error: {e}")