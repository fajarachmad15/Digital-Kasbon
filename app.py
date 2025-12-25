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
    </style>
""", unsafe_allow_html=True)

# --- LOGIKA LOGIN GOOGLE (WAJIB) ---
# Bagian ini memastikan user harus login dulu sebelum melihat isi aplikasi
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
if 'data_ringkasan' not in st.session_state: st.session_state.data_ringkasan = {}
if 'show_errors' not in st.session_state: st.session_state.show_errors = False
if 'mgr_logged_in' not in st.session_state: st.session_state.mgr_logged_in = False

# --- 4. TAMPILAN APPROVAL (MANAGER & CASHIER) ---
query_id = st.query_params.get("id")
if query_id:
    # Cek Status untuk Judul Portal
    try:
        creds = get_creds()
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID).worksheet("DATA_KASBON_AZKO")
        cell = sheet.find(query_id)
        row_data = sheet.row_values(cell.row)
        status_now = row_data[14]

        # Judul Portal Dinamis
        if status_now == "Pending":
            judul_portal = "Portal Approval Manager"
        elif status_now == "APPROVED":
            judul_portal = "Portal Verifikasi Cashier"
        else:
            judul_portal = "Portal Informasi Pengajuan"

        st.markdown(f'<span class="store-header">{judul_portal}</span>', unsafe_allow_html=True)

        # LOGIN CREDENTIAL (MANAGER / CASHIER)
        if not st.session_state.mgr_logged_in:
            st.subheader("🔐 Verifikasi Manager/Cashier")
            v_nik = st.text_input("NIK (6 Digit)", max_chars=6)
            v_pass = st.text_input("Password", type="password")
            if st.button("Masuk & Verifikasi", type="primary", use_container_width=True):
                if len(v_nik) == 6 and len(v_pass) >= 6:
                    records = client.open_by_key(SPREADSHEET_ID).worksheet("DATABASE_USER").get_all_records()
                    user = next((r for r in records if str(r['NIK']) == v_nik and str(r['Password']) == v_pass), None)
                    if user: 
                        st.session_state.mgr_logged_in = True
                        st.rerun()
                    else: st.error("NIK atau Password salah.")
                else: st.warning("Cek kembali NIK & Password.")
            st.stop()

        # TAMPILAN DATA
        st.info(f"### Rincian Pengajuan: {query_id}")
        c1, c2 = st.columns(2)
        with c1:
            st.write(f"**Tgl:** {row_data[0]}")
            st.write(f"**Dibayarkan:** {row_data[4]} / {row_data[5]}")
            st.write(f"**Dept:** {row_data[6]}")
        with c2:
            st.write(f"**Nominal:** Rp {int(row_data[7]):,}")
            st.write(f"**Terbilang:** {row_data[8]}")
            st.write(f"**Janji:** {row_data[11]}")
        
        st.write(f"**Keperluan:** {row_data[9]}")
        st.write(f"**Status Saat Ini:** `{status_now}`")
        st.divider()

        # --- LOGIKA MANAGER (PENDING) ---
        if status_now == "Pending":
            alasan = st.text_area("Alasan Reject (Wajib diisi jika Reject)", placeholder="Contoh: Nominal terlalu besar...")
            
            b1, b2 = st.columns(2)
            if b1.button("✓ APPROVE", use_container_width=True):
                sheet.update_cell(cell.row, 15, "APPROVED")
                # KIRIM EMAIL KE CASHIER (APPROVE)
                try:
                    cashier_info = row_data[12] # NIK - Nama (Email)
                    cashier_email = cashier_info.split("(")[1].split(")")[0]
                    cashier_name = cashier_info.split(" - ")[1].split(" (")[0]
                    
                    email_msg = f"""
                    Dear Bapak / Ibu {cashier_name}
                    <br><br>
                    Pengajuan kasbon dengan data dibawah ini telah di-<b>APPROVE</b> oleh Manager:
                    <br><br>
                    Nomor Pengajuan Kasbon : {query_id}<br>
                    Tgl dan Jam Pengajuan : {row_data[0]}<br>
                    Dibayarkan Kepada : {row_data[4]} / {row_data[5]}<br>
                    Departement : {row_data[6]}<br>
                    Senilai : Rp {int(row_data[7]):,} ({row_data[8]})<br>
                    Untuk Keperluan : {row_data[9]}<br>
                    Approval Pendukung : {row_data[10]}<br>
                    Janji Penyelesaian : {row_data[11]}
                    <br><br>
                    Silahkan klik <a href='{BASE_URL}?id={query_id}'>link berikut</a> untuk melanjutkan prosesnya.
                    """
                    send_email_with_attachment(cashier_email, f"Verifikasi Kasbon {query_id}", email_msg)
                except: pass
                
                st.success("Approved & Notifikasi Cashier terkirim!"); st.balloons(); st.rerun()

            if b2.button("✕ REJECT", use_container_width=True):
                if not alasan: st.error("Harap isi alasan reject!"); st.stop()
                sheet.update_cell(cell.row, 15, "REJECTED")
                # KIRIM EMAIL KE CASHIER (REJECT)
                try:
                    cashier_info = row_data[12]
                    cashier_email = cashier_info.split("(")[1].split(")")[0]
                    cashier_name = cashier_info.split(" - ")[1].split(" (")[0]
                    
                    email_msg = f"""
                    Dear Bapak / Ibu {cashier_name}
                    <br><br>
                    Pengajuan kasbon dengan data dibawah ini telah di-<b>REJECT</b> oleh Manager.<br>
                    <b>Reason Reject:</b> {alasan}
                    <br><br>
                    Nomor Pengajuan Kasbon : {query_id}<br>
                    Dibayarkan Kepada : {row_data[4]} / {row_data[5]}<br>
                    Senilai : Rp {int(row_data[7]):,}
                    <br><br>
                    Silahkan cek sistem.
                    """
                    send_email_with_attachment(cashier_email, f"Penolakan Kasbon {query_id}", email_msg)
                except: pass
                
                st.error("Pengajuan telah di-Reject."); st.rerun()

        # --- LOGIKA CASHIER (APPROVED) -> STEP 2 ---
        elif status_now == "APPROVED":
            st.info("Menunggu Verifikasi Pencairan Dana oleh Cashier.")
            if st.button("✓ KONFIRMASI PENCAIRAN DANA", use_container_width=True):
                sheet.update_cell(cell.row, 15, "COMPLETED")
                st.success("Dana Telah Dicairkan. Status Selesai.")
                st.balloons()
                st.rerun()
        
        else:
            st.warning(f"Pengajuan ini sudah selesai/ditolak: {status_now}")

    except Exception as e: st.error(f"Error Database: {e}")
    st.stop()

# --- 5. TAMPILAN INPUT USER (SAMA PERSIS DENGAN YANG LAMA) ---
if st.session_state.submitted:
    d = st.session_state.data_ringkasan
    st.success("## ✅ PENGAJUAN TELAH TERKIRIM")
    st.info(f"Email notifikasi telah dikirim ke Manager.")
    st.write("---")
    st.subheader("Ringkasan Pengajuan")
    col_res1, col_res2 = st.columns(2)
    with col_res1:
        st.write(f"**1. No. Pengajuan:** {d['no_pengajuan']}")
        st.write(f"**2. Kode Store:** {d['kode_store']}")
        st.write(f"**3. Dibayarkan Kepada:** {d['nama']} / {d['nip']}")
        st.write(f"**4. Departemen:** {d['dept']}")
    with col_res2:
        st.write(f"**6. Nominal:** Rp {int(d['nominal']):,}")
        st.write(f"**7. Terbilang:** {d['terbilang']}")
        st.write(f"**8. Keperluan:** {d['keperluan']}")
        st.write(f"**9. Janji Penyelesaian:** {d['janji']}")
    
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
                # REQ 1: TERBILANG 1.5x, BOLD, TITLE CASE
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
                            
                            # REQ 2: ID KASBON MMYY (Tanpa DD)
                            prefix = f"KB{kode_store}-{tgl_now.strftime('%m%y')}-" 
                            all_rows = sheet.get_all_values()
                            last_n = 0
                            for r in all_rows:
                                if len(r)>1 and r[1].startswith(prefix):
                                    try: last_n = max(last_n, int(r[1].split("-")[-1]))
                                    except: continue
                            no_p = f"{prefix}{str(last_n + 1).zfill(3)}"
                            
                            final_t = terbilang(int(nom_r)).title() + " Rupiah"
                            link_db = "Terlampir di Email" if bukti else "-"
                            
                            sheet.append_row([
                                tgl_now.strftime("%Y-%m-%d %H:%M:%S"), no_p, kode_store, pic_email, 
                                nama_p, nip, dept, nom_r, final_t, kep, link_db, 
                                janji.strftime("%d/%m/%Y"), sc_f, mgr_f, "Pending", ""
                            ])
                            
                            mgr_clean = mgr_f.split(" - ")[1].split(" (")[0]
                            tgl_full = tgl_now.strftime("%d/%m/%Y %H:%M")
                            app_link = f"{BASE_URL}?id={no_p}"
                            link_html = "Lihat Lampiran di bawah" if bukti else "-"
                            
                            # --- PERUBAHAN TAMPILAN EMAIL (MENYATU & SESUAI REQUEST) ---
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
                                    <tr><td style='padding: 2px 0;'>Approval Pendukung</td><td>: {link_html}</td></tr>
                                    <tr><td style='padding: 2px 0;'>Janji Penyelesaian</td><td>: {janji.strftime("%d/%m/%Y")}</td></tr>
                                </table>
                                
                                <div style='margin-top: 15px; margin-bottom: 10px;'>
                                    Silahkan klik <a href='{app_link}' style='text-decoration: none; color: #0000EE;'>link berikut</a> untuk melanjutkan prosesnya
                                </div>
                                
                                <div>Terima Kasih</div>
                            </body></html>
                            """
                            send_email_with_attachment(mgr_map[mgr_f], f"Pengajuan Kasbon {no_p}", email_body, bukti)
                        
                        st.session_state.data_ringkasan = {'no_pengajuan': no_p, 'kode_store': kode_store, 'nama': nama_p, 'nip': nip, 'dept': dept, 'nominal': nom_r, 'terbilang': final_t, 'keperluan': kep, 'janji': janji.strftime("%d/%m/%Y")}
                        st.session_state.submitted = True; st.session_state.show_errors = False; st.rerun()
                    except Exception as e: st.error(f"Error Sistem: {e}")
                else:
                    st.session_state.show_errors = True; st.rerun()
        except Exception as e: st.error(f"Database Error: {e}")