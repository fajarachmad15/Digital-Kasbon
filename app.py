import streamlit as st
import datetime
import gspread
import smtplib
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

    /* STYLING TOMBOL APPROVAL TRANSPARAN */
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        color: white !important;
    }
    div.stColumn:nth-of-type(1) [data-testid="stButton"] button {
        background-color: rgba(40, 167, 69, 0.4) !important;
        border: 1px solid rgba(40, 167, 69, 0.6) !important;
    }
    div.stColumn:nth-of-type(2) [data-testid="stButton"] button {
        background-color: rgba(220, 53, 69, 0.4) !important;
        border: 1px solid rgba(220, 53, 69, 0.6) !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- 2. LOGIKA LOGIN GOOGLE (NATIVE & WAJIB) ---
if not st.experimental_user.is_logged_in:
    st.markdown("## 🌐 Kasbon Digital Petty Cash")
    st.info("Silakan login menggunakan akun Google Anda untuk melanjutkan.")
    if st.button("Sign in with Google", type="primary", use_container_width=True):
        st.login()
    st.stop()

# Ambil email asli dari Google
pic_email = st.experimental_user.email

# --- 3. KONFIGURASI ---
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
        msg = MIMEMultipart(); msg['Subject'] = subject; msg['To'] = to_email
        msg['From'] = formataddr(("Bot_KasbonPC_Digital <No-Reply>", SENDER_EMAIL))
        msg.attach(MIMEText(message_body, 'html'))
        if attachment_file:
            part = MIMEBase('application', 'octet-stream'); part.set_payload(attachment_file.getvalue())
            encoders.encode_base64(part); part.add_header('Content-Disposition', f'attachment; filename={attachment_file.name}')
            msg.attach(part)
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(SENDER_EMAIL, APP_PASSWORD); server.sendmail(SENDER_EMAIL, to_email, msg.as_string())
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

# --- 4. SESSION STATE ---
if 'submitted' not in st.session_state: st.session_state.submitted = False
if 'data_ringkasan' not in st.session_state: st.session_state.data_ringkasan = {}
if 'show_errors' not in st.session_state: st.session_state.show_errors = False
if 'mgr_logged_in' not in st.session_state: st.session_state.mgr_logged_in = False

# --- 5. TAMPILAN APPROVAL MANAGER ---
query_id = st.query_params.get("id")
if query_id:
    st.markdown('<span class="store-header">Portal Approval Manager</span>', unsafe_allow_html=True)
    
    # LOGIN MANAGER (NIK & PASSWORD)
    if not st.session_state.mgr_logged_in:
        st.subheader("🔐 Verifikasi Manager/Cashier")
        v_nik = st.text_input("NIK (6 Digit)", max_chars=6)
        v_pass = st.text_input("Password", type="password")
        if st.button("Masuk & Verifikasi", type="primary", use_container_width=True):
            if len(v_nik) == 6 and len(v_pass) >= 6:
                records = gspread.authorize(get_creds()).open_by_key(SPREADSHEET_ID).worksheet("DATABASE_USER").get_all_records()
                user = next((r for r in records if str(r['NIK']) == v_nik and str(r['Password']) == v_pass), None)
                if user: st.session_state.mgr_logged_in = True; st.rerun()
                else: st.error("NIK atau Password salah.")
            else: st.warning("Cek kembali NIK & Password.")
        st.stop()

    try:
        sheet = gspread.authorize(get_creds()).open_by_key(SPREADSHEET_ID).worksheet("DATA_KASBON_AZKO")
        cell = sheet.find(query_id); row_data = sheet.row_values(cell.row)
        st.info(f"### Rincian Pengajuan: {query_id}")
        c1, c2 = st.columns(2)
        with c1:
            st.write(f"**Tgl:** {row_data[0]}"); st.write(f"**Dibayarkan:** {row_data[4]} / {row_data[5]}")
            st.write(f"**Dept:** {row_data[6]}")
        with c2:
            st.write(f"**Nominal:** Rp {int(row_data[7]):,}")
            st.write(f"**Terbilang:** {row_data[8]}") # FITUR TERBILANG
            st.write(f"**Janji:** {row_data[11]}")
        st.write(f"**Keperluan:** {row_data[9]} | **Status:** `{row_data[14]}`")
        st.divider()

        if row_data[14] == "Pending":
            b1, b2 = st.columns(2)
            if b1.button("✓ APPROVE", use_container_width=True):
                sheet.update_cell(cell.row, 15, "APPROVED"); st.balloons(); st.rerun()
            if b2.button("✕ REJECT", use_container_width=True):
                sheet.update_cell(cell.row, 15, "REJECTED"); st.rerun()
        else: st.warning(f"Sudah diproses: {row_data[14]}")
    except: pass
    st.stop()

# --- 6. TAMPILAN INPUT USER (PIC) ---
if st.session_state.submitted:
    d = st.session_state.data_ringkasan
    st.success("## ✅ PENGAJUAN TELAH TERKIRIM")
    st.write(f"**No:** {d['no_pengajuan']} | **Nominal:** Rp {int(d['nominal']):,} ({d['terbilang']})")
    c1, c2 = st.columns(2)
    if c1.button("Buat Baru", use_container_width=True): st.session_state.submitted = False; st.rerun()
    if c2.button("Logout Google", use_container_width=True): st.logout()

else:
    # Tampilkan email asli di atas
    st.caption(f"Logged in as: **{pic_email}**") 
    
    st.subheader("📍 Identifikasi Lokasi")
    kode_store = st.text_input("Masukkan Kode Store", placeholder="Contoh: A644").upper()

    if kode_store:
        try:
            creds = get_creds()
            client = gspread.authorize(creds)
            records = client.open_by_key(SPREADSHEET_ID).worksheet("DATABASE_USER").get_all_records()
            store_info = next((u for u in records if str(u['Kode_Store']) == kode_store), None)
            
            if not store_info:
                st.error("⚠️ Kode store tidak ada atau belum terdaftar"); st.stop()
            
            st.markdown(f'<span class="store-header">Unit Bisnis Store: {store_info["Nama_Store"]}</span>', unsafe_allow_html=True)
            
            # EMAIL REQUEST OTOMATIS (Non-editable, diambil dari Google)
            st.markdown('<div class="label-container"><span class="label-text">Email Request</span></div>', unsafe_allow_html=True)
            st.text_input("", value=pic_email, disabled=True)

            # FORM INPUT LENGKAP DENGAN VALIDASI MERAH
            err_nama = '<span class="error-tag">Harap dilengkapi</span>' if st.session_state.show_errors and not st.session_state.get('nama_val') else ''
            st.markdown(f'<div class="label-container"><span class="label-text">Dibayarkan Kepada</span>{err_nama}</div>', unsafe_allow_html=True)
            nama_p = st.text_input("", key="nama_val")

            err_nip = '<span class="error-tag">Wajib 6 digit</span>' if st.session_state.show_errors and (not st.session_state.get('nip_val') or len(st.session_state.get('nip_val')) != 6) else ''
            st.markdown(f'<div class="label-container"><span class="label-text">NIP (Wajib 6 Digit)</span>{err_nip}</div>', unsafe_allow_html=True)
            nip_p = st.text_input("", max_chars=6, key="nip_val")

            err_dept = '<span class="error-tag">Pilih satu</span>' if st.session_state.show_errors and st.session_state.get('dept_val') == "-" else ''
            st.markdown(f'<div class="label-container"><span class="label-text">Departemen</span>{err_dept}</div>', unsafe_allow_html=True)
            dept = st.selectbox("", ["-", "Operational", "Sales", "Inventory", "HR", "Other"], key="dept_val")

            err_nom = '<span class="error-tag">Hanya angka</span>' if st.session_state.show_errors and (not st.session_state.get('nom_val') or not st.session_state.get('nom_val').isdigit()) else ''
            st.markdown(f'<div class="label-container"><span class="label-text">Nominal (Angka)</span>{err_nom}</div>', unsafe_allow_html=True)
            nom_r = st.text_input("", key="nom_val")
            if nom_r.isdigit(): st.caption(f"**Terbilang:** {terbilang(int(nom_r))} Rupiah")

            err_kep = '<span class="error-tag">Harap dilengkapi</span>' if st.session_state.show_errors and not st.session_state.get('kep_val') else ''
            st.markdown(f'<div class="label-container"><span class="label-text">Keperluan</span>{err_kep}</div>', unsafe_allow_html=True)
            kep = st.text_input("", key="kep_val")

            opsi = st.radio("Metode Lampiran:", ["Upload File", "Kamera"])
            bukti = st.file_uploader("Pilih file") if opsi == "Upload File" else st.camera_input("Ambil Foto")
            
            st.markdown('<div class="label-container"><span class="label-text">Janji Penyelesaian</span></div>', unsafe_allow_html=True)
            janji = st.date_input("", min_value=datetime.date.today())

            managers_db = [f"{u['NIK']} - {u['Nama Lengkap']} ({u['Email']})" for u in records if str(u['Kode_Store']) == kode_store and u['Role'] == 'Manager']
            cashiers_db = [f"{u['NIK']} - {u['Nama Lengkap']} ({u['Email']})" for u in records if str(u['Kode_Store']) == kode_store and u['Role'] == 'Senior Cashier']
            
            # MAPPING EMAIL MANAGER UNTUK PENGIRIMAN
            mgr_email_map = {f"{u['NIK']} - {u['Nama Lengkap']} ({u['Email']})": u['Email'] for u in records if str(u['Kode_Store']) == kode_store and u['Role'] == 'Manager'}

            err_mgr = '<span class="error-tag">Pilih Manager</span>' if st.session_state.show_errors and st.session_state.get('mgr_val') == "-" else ''
            st.markdown(f'<div class="label-container"><span class="label-text">Manager Incharge</span>{err_mgr}</div>', unsafe_allow_html=True)
            mgr_f = st.selectbox("", ["-"] + managers_db, key="mgr_val")

            err_sc = '<span class="error-tag">Pilih Cashier</span>' if st.session_state.show_errors and st.session_state.get('sc_val') == "-" else ''
            st.markdown(f'<div class="label-container"><span class="label-text">Senior Cashier Incharge</span>{err_sc}</div>', unsafe_allow_html=True)
            sc_f = st.selectbox("", ["-"] + cashiers_db, key="sc_val")

            st.divider()

            if st.button("Kirim Pengajuan", type="primary", use_container_width=True):
                # Validasi Lengkap (Super App Logic)
                is_valid = nama_p and len(nip_p)==6 and nom_r.isdigit() and kep and dept!="-" and mgr_f!="-" and sc_f!="-"
                
                if is_valid:
                    try:
                        sheet = client.open_by_key(SPREADSHEET_ID).worksheet("DATA_KASBON_AZKO")
                        tgl_n = datetime.datetime.now(WIB)
                        
                        # NOMOR KASBON (Reset Harian)
                        prefix = f"KB{kode_store}-{tgl_n.strftime('%m%y')}-"
                        count_today = sum(1 for row in sheet.get_all_values() if row[0].startswith(tgl_n.strftime("%Y-%m-%d")))
                        no_p = f"{prefix}{str(count_today + 1).zfill(3)}"
                        
                        final_t = terbilang(int(nom_r)) + " Rupiah"
                        janji_str = janji.strftime("%d/%m/%Y")
                        
                        # RECORD KE DATABASE (Email Asli)
                        sheet.append_row([
                            tgl_n.strftime("%Y-%m-%d %H:%M:%S"), no_p, kode_store, pic_email, 
                            nama_p, nip_p, dept, nom_r, final_t, kep, "Terlampir", 
                            janji_str, sc_f, mgr_f, "Pending"
                        ])
                        
                        # KIRIM EMAIL KE MANAGER
                        app_link = f"{BASE_URL}?id={no_p}"
                        mgr_clean = mgr_f.split(" - ")[1].split(" (")[0]
                        target_email = mgr_email_map.get(mgr_f)
                        
                        if target_email:
                            email_body = f"""
                            <p>Dear {mgr_clean},</p>
                            <p>Mohon approval pengajuan kasbon <b>{no_p}</b> dari <b>{nama_p}</b> senilai <b>Rp {int(nom_r):,}</b>.</p>
                            <p>Keperluan: {kep}</p>
                            <p>Silakan klik <a href='{app_link}'><b>LINK INI</b></a> untuk memproses.</p>
                            """
                            send_email_with_attachment(target_email, f"Approval Kasbon {no_p}", email_body, bukti)

                        st.session_state.data_ringkasan = {'no_pengajuan': no_p, 'nominal': nom_r, 'terbilang': final_t}
                        st.session_state.submitted = True
                        st.session_state.show_errors = False
                        st.rerun()

                    except Exception as e: st.error(f"Error Sistem: {e}")
                else:
                    st.session_state.show_errors = True; st.rerun()
        
        except Exception as e: st.error("Gagal memuat database store/user.")