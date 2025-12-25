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
from oauth2client.service_account import ServiceAccountCredentials

# --- 1. SETTING HALAMAN & CSS ---
st.set_page_config(page_title="Kasbon Digital Petty Cash", layout="centered")

st.markdown("""
    <style>
    [data-testid="InputInstructions"] { display: none !important; }
    .label-container { display: flex; justify-content: space-between; align-items: center; margin-bottom: -10px; padding-top: 10px; }
    .label-text { font-weight: 600; color: inherit; font-size: 15px; }
    div[data-testid="stWidgetLabel"] { display: none; }
    .store-header { color: #FF0000 !important; font-size: 1.25rem; font-weight: 600; margin-top: 1rem; margin-bottom: 1rem; display: block; }

    /* STYLING TOMBOL APPROVAL TRANSPARAN (Image 707050) */
    .stButton > button { border-radius: 8px; font-weight: 600; color: white !important; }
    div.stColumn:nth-of-type(1) [data-testid="stButton"] button { background-color: rgba(40, 167, 69, 0.4) !important; border: 1px solid rgba(40, 167, 69, 0.6) !important; }
    div.stColumn:nth-of-type(2) [data-testid="stButton"] button { background-color: rgba(220, 53, 69, 0.4) !important; border: 1px solid rgba(220, 53, 69, 0.6) !important; }
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

# --- 3. SESSION STATE ---
if 'submitted' not in st.session_state: st.session_state.submitted = False
if 'mgr_logged_in' not in st.session_state: st.session_state.mgr_logged_in = False

# --- 4. TAMPILAN APPROVAL MANAGER (NIK & PASSWORD) ---
query_id = st.query_params.get("id")
if query_id:
    st.markdown('<span class="store-header">Portal Approval Manager</span>', unsafe_allow_html=True)
    
    # Login Manager Wajib (Image 70ec0a)
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
            else: st.warning("NIK wajib 6 digit & Password min 6 char.")
        st.stop()

    try:
        sheet = gspread.authorize(get_creds()).open_by_key(SPREADSHEET_ID).worksheet("DATA_KASBON_AZKO")
        cell = sheet.find(query_id); row_data = sheet.row_values(cell.row)
        st.info(f"### Rincian Pengajuan: {query_id}")
        c1, c2 = st.columns(2)
        with c1:
            st.write(f"**Tgl:** {row_data[0]}"); st.write(f"**Dibayarkan:** {row_data[4]} / {row_data[5]}")
        with c2:
            st.write(f"**Nominal:** Rp {int(row_data[7]):,}")
            st.write(f"**Terbilang:** {row_data[8]}") # TERBILANG UNTUK MANAGER
        st.write(f"**Keperluan:** {row_data[9]} | **Status:** `{row_data[14]}`")
        if row_data[14] == "Pending":
            b1, b2 = st.columns(2)
            if b1.button("✓ APPROVE", use_container_width=True): sheet.update_cell(cell.row, 15, "APPROVED"); st.balloons(); st.rerun()
            if b2.button("✕ REJECT", use_container_width=True): sheet.update_cell(cell.row, 15, "REJECTED"); st.rerun()
        else: st.warning(f"Sudah diproses: {row_data[14]}")
    except: pass
    st.stop()

# --- 5. LOGIKA LOGIN GOOGLE (PENGADAN) ---
# Memastikan Pengaju Wajib Klik (image_71650b.png)
if not st.user.get("email"):
    st.subheader("🌐 Kasbon Digital Petty Cash")
    st.info("Klik tombol di bawah untuk masuk menggunakan akun Gmail aktif di perangkat Anda (Tanpa Ketik Manual).")
    if st.button("Sign in with Google", type="primary", use_container_width=True):
        st.login()
    st.stop()

# --- 6. TAMPILAN INPUT USER (PIC) ---
pic_email = st.user.email # MENANGKAP EMAIL ASLI DARI GOOGLE

if st.session_state.submitted:
    st.success("## ✅ PENGAJUAN TELAH TERKIRIM"); st.write("---")
    if st.button("Buat Baru"): st.session_state.submitted = False; st.rerun()
    if st.button("Logout"): st.logout()

else:
    st.caption(f"Logged in as: **{pic_email}**") # Verifikasi Visual
    st.subheader("📍 Identifikasi Lokasi")
    kode_store = st.text_input("Masukkan Kode Store", placeholder="Contoh: A644").upper()
    if kode_store:
        try:
            records = gspread.authorize(get_creds()).open_by_key(SPREADSHEET_ID).worksheet("DATABASE_USER").get_all_records()
            store_info = next((u for u in records if str(u['Kode_Store']) == kode_store), None)
            if not store_info: st.error("⚠️ Kode store tidak ada"); st.stop()
            st.markdown(f'<span class="store-header">Unit Bisnis Store: {store_info["Nama_Store"]}</span>', unsafe_allow_html=True)
            st.text_input("Email Request", value=pic_email, disabled=True)
            
            nama_p = st.text_input("Dibayarkan Kepada"); nip_p = st.text_input("NIP (Wajib 6 Digit)", max_chars=6)
            dept = st.selectbox("Departemen", ["-", "Operational", "Sales", "Inventory", "HR", "Other"])
            nom_r = st.text_input("Nominal (Angka)")
            if nom_r.isdigit(): st.caption(f"**Terbilang:** {terbilang(int(nom_r))} Rupiah")
            kep = st.text_input("Keperluan")
            bukti = st.file_uploader("Lampiran") if st.radio("Metode", ["File", "Kamera"]) == "File" else st.camera_input("Foto")
            janji = st.date_input("Janji Penyelesaian", min_value=datetime.date.today())
            
            mgrs = [f"{u['NIK']} - {u['Nama Lengkap']} ({u['Email']})" for u in records if str(u['Kode_Store']) == kode_store and u['Role'] == 'Manager']
            scs = [f"{u['NIK']} - {u['Nama Lengkap']} ({u['Email']})" for u in records if str(u['Kode_Store']) == kode_store and u['Role'] == 'Senior Cashier']
            mgr_f = st.selectbox("Manager", ["-"] + mgrs); sc_f = st.selectbox("Senior Cashier", ["-"] + scs)

            if st.button("Kirim Pengajuan", type="primary"):
                if nama_p and len(nip_p)==6 and nom_r.isdigit() and mgr_f!="-" and sc_f!="-":
                    sheet = gspread.authorize(get_creds()).open_by_key(SPREADSHEET_ID).worksheet("DATA_KASBON_AZKO")
                    tgl_n = datetime.datetime.now(WIB)
                    # ID Kasbon Tanpa Hari (Reset Harian)
                    prefix = f"KB{kode_store}-{tgl_n.strftime('%m%y')}-"
                    count_today = sum(1 for r in sheet.get_all_values() if r[0].startswith(tgl_n.strftime("%Y-%m-%d")))
                    no_p = f"{prefix}{str(count_today + 1).zfill(3)}"
                    
                    final_t = terbilang(int(nom_r)) + " Rupiah"
                    # RECORD EMAIL ASLI DI KOLOM D (Image de50c1)
                    sheet.append_row([tgl_n.strftime("%Y-%m-%d %H:%M:%S"), no_p, kode_store, pic_email, nama_p, nip_p, dept, nom_r, final_t, kep, "Terlampir", janji.strftime("%d/%m/%Y"), sc_f, mgr_f, "Pending"])
                    app_link = f"{BASE_URL}?id={no_p}"
                    send_email_with_attachment(mgr_f.split("(")[1][:-1], f"Approval {no_p}", f"Data {nama_p} Rp {int(nom_r):,}. <a href='{app_link}'>Klik</a>", bukti)
                    st.session_state.submitted = True; st.rerun()
                else: st.error("Lengkapi data!")
        except: st.error("Gagal muat database.")