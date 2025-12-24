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
# PERBAIKAN IMAGE 707f4f: Nama aplikasi diganti
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

    /* PERBAIKAN IMAGE 707050: Styling Tombol Transparan & Simbol Putih */
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        color: white !important;
    }
    /* Tombol Approve (Indeks Pertama di kolom) */
    div.stColumn:nth-of-type(1) [data-testid="stButton"] button {
        background-color: rgba(40, 167, 69, 0.3) !important;
        border: 1px solid rgba(40, 167, 69, 0.5) !important;
    }
    /* Tombol Reject (Indeks Kedua di kolom) */
    div.stColumn:nth-of-type(2) [data-testid="stButton"] button {
        background-color: rgba(220, 53, 69, 0.3) !important;
        border: 1px solid rgba(220, 53, 69, 0.5) !important;
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
    except Exception as e:
        st.error(f"Gagal kirim email: {e}")
        return False

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

# --- 4. TAMPILAN APPROVAL MANAGER ---
query_id = st.query_params.get("id")
if query_id:
    # PERBAIKAN IMAGE 707f4f: Nama Header
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
            st.write(f"**Dibayarkan Kepada:** {row_data[4]} / {row_data[5]}")
            st.write(f"**Departemen:** {row_data[6]}")
        with col_app2:
            st.write(f"**Nominal:** Rp {int(row_data[7]):,}")
            st.write(f"**Keperluan:** {row_data[9]}")
            st.write(f"**Janji Penyelesaian:** {row_data[11]}")
            
        st.write(f"**Status Saat Ini:** `{row_data[14]}`")
        st.divider()
        
        if row_data[14] == "Pending":
            c1, c2 = st.columns(2)
            # PERBAIKAN IMAGE 707050: Simbol checklist dan silang putih (✓ dan ✕)
            if c1.button("✓ APPROVE", use_container_width=True):
                sheet.update_cell(cell.row, 15, "APPROVED")
                st.success("Berhasil di-Approve!"); st.balloons()
                st.rerun()
            if c2.button("✕ REJECT", use_container_width=True):
                sheet.update_cell(cell.row, 15, "REJECTED")
                st.error("Pengajuan telah di-Reject.")
                st.rerun()
        else:
            # IMAGE 707be9: Sudah bagus, tidak diubah
            st.warning(f"Pengajuan ini sudah diproses dengan status: {row_data[14]}")
    except: 
        # PERBAIKAN IMAGE 70d18a: Menghapus notifikasi error "Data Kasbon tidak ditemukan"
        pass

# --- 5. TAMPILAN INPUT USER ---
else:
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
        
        if st.button("Buat Pengajuan Baru"):
            st.session_state.submitted = False
            st.session_state.show_errors = False
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

                managers_db = [f"{u['NIK']} - {u['Nama Lengkap']} ({u['Email']})" 
                               for u in user_records if str(u['Kode_Store']) == kode_store and u['Role'] == 'Manager']
                cashiers_db = [f"{u['NIK']} - {u['Nama Lengkap']} ({u['Email']})" 
                               for u in user_records if str(u['Kode_Store']) == kode_store and u['Role'] == 'Senior Cashier']
                
                manager_email_map = {f"{u['NIK']} - {u['Nama Lengkap']} ({u['Email']})": u['Email'] 
                                     for u in user_records if str(u['Kode_Store']) == kode_store and u['Role'] == 'Manager'}
                
                if not managers_db or not cashiers_db:
                    st.error(f"⚠️ Data Manager/Cashier untuk {nama_store_display} tidak lengkap di DATABASE_USER.")
                    st.stop()
            except Exception as e:
                if "Kode_Store" in str(e) or "KeyError" in str(type(e).__name__):
                    st.error("⚠️ Kode store tidak ada atau belum terdaftar")
                else:
                    st.error(f"Gagal memuat database user: {e}")
                st.stop()

            tgl_obj = datetime.datetime.now(WIB)
            # PERBAIKAN IMAGE 707f4f: Header Petty Cash
            st.markdown(f'<span class="store-header">Unit Bisnis Store: {nama_store_display}</span>', unsafe_allow_html=True)
            
            st.markdown('<div class="label-container"><span class="label-text">Email Request</span></div>', unsafe_allow_html=True)
            email_req = st.text_input("", value=SENDER_EMAIL, disabled=True)
            
            err_nama = '<span class="error-tag">Harap dilengkapi</span>' if st.session_state.show_errors and not st.session_state.get('nama_val') else ''
            st.markdown(f'<div class="label-container"><span class="label-text">Dibayarkan Kepada (Nama Lengkap)</span>{err_nama}</div>', unsafe_allow_html=True)
            nama_penerima = st.text_input("", key="nama_val")

            err_nip = '<span class="error-tag">Isi sesuai format</span>' if st.session_state.show_errors and (not st.session_state.get('nip_val') or len(st.session_state.get('nip_val')) != 6) else ''
            st.markdown(f'<div class="label-container"><span class="label-text">NIP (Wajib 6 Digit)</span>{err_nip}</div>', unsafe_allow_html=True)
            nip = st.text_input("", max_chars=6, key="nip_val")

            err_dept = '<span class="error-tag">Harap dipilih</span>' if st.session_state.show_errors and st.session_state.get('dept_val') == "-" else ''
            st.markdown(f'<div class="label-container"><span class="label-text">Departemen</span>{err_dept}</div>', unsafe_allow_html=True)
            dept = st.selectbox("", ["-", "Operational", "Sales", "Inventory", "HR", "Other"], key="dept_val")

            err_nom = '<span class="error-tag">Hanya angka</span>' if st.session_state.show_errors and (not st.session_state.get('nom_val') or not st.session_state.get('nom_val').isdigit()) else ''
            st.markdown(f'<div class="label-container"><span class="label-text">Nominal (Hanya Angka)</span>{err_nom}</div>', unsafe_allow_html=True)
            nominal_raw = st.text_input("", key="nom_val")
            if nominal_raw and nominal_raw.isdigit():
                st.caption(f"**Terbilang:** {terbilang(int(nominal_raw)) if int(nominal_raw)>0 else 'Nol'} Rupiah")

            err_kep = '<span class="error-tag">Harap dilengkapi</span>' if st.session_state.show_errors and not st.session_state.get('kep_val') else ''
            st.markdown(f'<div class="label-container"><span class="label-text">Untuk Keperluan</span>{err_kep}</div>', unsafe_allow_html=True)
            keperluan = st.text_input("", key="kep_val")
            
            st.write("📸 **Bukti Lampiran (Maks 5MB)**")
            opsi = st.radio("Metode Lampiran:", ["Upload File", "Kamera"])
            bukti_file = st.file_uploader("Pilih file") if opsi == "Upload File" else st.camera_input("Ambil Foto")
            
            st.markdown('<div class="label-container"><span class="label-text">Janji Penyelesaian</span></div>', unsafe_allow_html=True)
            janji_tgl = st.date_input("", min_value=datetime.date.today(), key="janji_val")
            
            err_sc = '<span class="error-tag">Harap dipilih</span>' if st.session_state.show_errors and st.session_state.get('sc_val') == "-" else ''
            st.markdown(f'<div class="label-container"><span class="label-text">Senior Cashier Incharge</span>{err_sc}</div>', unsafe_allow_html=True)
            senior_cashier = st.selectbox("", ["-"] + cashiers_db, key="sc_val")
            
            err_mgr = '<span class="error-tag">Harap dipilih</span>' if st.session_state.show_errors and st.session_state.get('mgr_val') == "-" else ''
            st.markdown(f'<div class="label-container"><span class="label-text">Manager Incharge</span>{err_mgr}</div>', unsafe_allow_html=True)
            mgr_name_full = st.selectbox("", ["-"] + managers_db, key="mgr_val")

            st.divider()
            
            if st.session_state.show_errors:
                st.error("⚠️ Mohon lengkapi semua kolom yang bertanda merah di atas sebelum mengirim.")

            if st.button("Kirim Pengajuan", type="primary"):
                is_valid = nama_penerima and len(nip)==6 and nominal_raw.isdigit() and keperluan and dept!="-" and senior_cashier!="-" and mgr_name_full!="-"
                
                if is_valid:
                    try:
                        with st.spinner("Processing..."):
                            creds = get_creds()
                            client = gspread.authorize(creds)
                            sheet = client.open_by_key(SPREADSHEET_ID).worksheet("DATA_KASBON_AZKO")
                            
                            tgl_str = tgl_obj.strftime("%d%m%y")
                            all_records = sheet.get_all_values()
                            prefix = f"KB{kode_store}-{tgl_str}-"
                            last_num = 0
                            for row in all_records:
                                if len(row) > 1 and row[1].startswith(prefix):
                                    try:
                                        num = int(row[1].split("-")[-1])
                                        last_num = max(last_num, num)
                                    except: continue
                            no_pengajuan = f"{prefix}{str(last_num + 1).zfill(3)}"

                            link_database = "Terlampir di Email" if bukti_file else "-"
                            final_terbilang = (terbilang(int(nominal_raw)) if int(nominal_raw) > 0 else "Nol") + " Rupiah"
                            tgl_full = tgl_obj.strftime("%d/%m/%Y %H:%M")
                            janji_str = janji_tgl.strftime("%d/%m/%Y")
                            
                            data_final = [
                                datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S"), 
                                no_pengajuan, kode_store, email_req,
                                nama_penerima, nip, dept, nominal_raw, final_terbilang, keperluan,
                                link_database, janji_str, senior_cashier, mgr_name_full, "Pending"
                            ]
                            sheet.append_row(data_final)

                            app_link = f"{BASE_URL}?id={no_pengajuan}"
                            mgr_name_clean = mgr_name_full.split(" - ")[1].split(" (")[0]
                            link_html = "Lihat Lampiran di bawah" if bukti_file else "-"
                            
                            subject_email = f"Pengajuan Kasbon {no_pengajuan}"
                            email_body = f"""
                            <html>
                            <body style='font-family: Arial, sans-serif; line-height: 1.6;'>
                                <p>Dear Bapak / Ibu <b>{mgr_name_clean}</b></p>
                                <p>Mohon approvalnya untuk pengajuan kasbon dengan data di bawah ini :</p>
                                <table style='border-collapse: collapse; width: 100%;'>
                                    <tr><td style='width: 200px;'><b>Nomor Pengajuan Kasbon</b></td><td>: {no_pengajuan}</td></tr>
                                    <tr><td><b>Tgl dan Jam Pengajuan</b></td><td>: {tgl_full}</td></tr>
                                    <tr><td><b>Dibayarkan Kepada</b></td><td>: {nama_penerima} / {nip}</td></tr>
                                    <tr><td><b>Departement</b></td><td>: {dept}</td></tr>
                                    <tr><td><b>Senilai</b></td><td>: Rp {int(nominal_raw):,} ({final_terbilang})</td></tr>
                                    <tr><td><b>Untuk Keperluan</b></td><td>: {keperluan}</td></tr>
                                    <tr><td><b>Approval Pendukung</b></td><td>: {link_html}</td></tr>
                                    <tr><td><b>Janji Penyelesaian</b></td><td>: {janji_str}</td></tr>
                                </table>
                                <p>Silahkan klik <a href='{app_link}'><b>link berikut</b></a> untuk melanjutkan prosesnya</p>
                                <p>Terima Kasih</p>
                            </body>
                            </html>
                            """
                            target_email = manager_email_map[mgr_name_full]
                            send_email_with_attachment(target_email, subject_email, email_body, bukti_file)
                        
                        st.session_state.data_ringkasan = {
                            'no_pengajuan': no_pengajuan, 'kode_store': kode_store, 'nama': nama_penerima, 'nip': nip,
                            'dept': dept, 'nominal': nominal_raw, 'terbilang': final_terbilang, 
                            'keperluan': keperluan, 'janji': janji_str, 'manager': mgr_name_clean
                        }
                        st.session_state.submitted = True
                        st.session_state.show_errors = False
                        st.rerun()
                    except Exception as e: st.error(f"Error Sistem: {e}")
                else:
                    st.session_state.show_errors = True
                    st.rerun()