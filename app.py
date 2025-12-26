import streamlit as st
import datetime, gspread, smtplib, requests, base64, time
from email.mime.text import MIMEText
from email.utils import formataddr
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIG & CONSTANTS ---
st.set_page_config(page_title="Kasbon Digital Petty Cash", layout="centered")
SENDER_EMAIL = "achmad.setiawan@kawanlamacorp.com"
APP_PASSWORD = st.secrets["APP_PASSWORD"]
BASE_URL = "https://digital-kasbon-ahi.streamlit.app"
SPREADSHEET_ID = "1TGsCKhBC0E0hup6RGVbGrpB6ds5Jdrp5tNlfrBORzaI"
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbylCNQsYQCIvO2qWtEUIq7gPufCgx4U5sbPasGVMGTIbaZhRFZBnpcMiHMlB2CpsEpj/exec"
DRIVE_FOLDER_ID = "1H6aZbRbJ7Kw7zdTqkIED1tQUrBR43dBr"
WIB = datetime.timezone(datetime.timedelta(hours=7))

# --- CSS STYLING ---
st.markdown("""
    <style>
    [data-testid="InputInstructions"], div[data-testid="stWidgetLabel"], div[data-testid="stInputNumberStepContainer"] { display: none !important; }
    .label-container { display: flex; justify-content: space-between; align-items: center; margin-bottom: -10px; padding-top: 10px; }
    .label-text { font-weight: 600; font-size: 15px; }
    .error-tag { color: #FF0000 !important; font-size: 13px; font-weight: bold; }
    .store-header { color: #FF0000 !important; font-size: 1.25rem; font-weight: 600; margin: 1rem 0; display: block; }
    .stButton > button { border-radius: 8px; font-weight: 600; color: white !important; }
    div.stColumn:nth-of-type(1) [data-testid="stButton"] button { background-color: rgba(40, 167, 69, 0.3) !important; border: 1px solid rgba(40, 167, 69, 0.5) !important; }
    div.stColumn:nth-of-type(2) [data-testid="stButton"] button { background-color: rgba(220, 53, 69, 0.3) !important; border: 1px solid rgba(220, 53, 69, 0.5) !important; }
    [data-testid="stNumberInput"] button { display: none !important; }
    </style>
""", unsafe_allow_html=True)

# --- AUTH CHECK ---
if not st.experimental_user.is_logged_in:
    st.markdown("## 🌐 Kasbon Digital Petty Cash")
    st.info("Silakan login menggunakan akun Google Anda untuk melanjutkan.")
    if st.button("Sign in with Google", type="primary", use_container_width=True): st.login()
    st.stop()

pic_email = st.experimental_user.email

# --- HELPER FUNCTIONS ---
def get_client():
    creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"])
    return gspread.authorize(creds)

def send_email(to_email, subject, html_body):
    try:
        msg = MIMEText(html_body, 'html')
        msg['Subject'], msg['To'], msg['From'] = subject, to_email, formataddr(("Bot_KasbonPC_Digital <No-Reply>", SENDER_EMAIL))
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(SENDER_EMAIL, APP_PASSWORD)
            server.sendmail(SENDER_EMAIL, to_email, msg.as_string())
        return True
    except: return False

def upload_file(file_obj, filename):
    try:
        payload = {
            "filename": filename, 
            "filedata": base64.b64encode(file_obj.getvalue()).decode('utf-8'), 
            "mimetype": file_obj.type, 
            "folderId": DRIVE_FOLDER_ID
        }
        res = requests.post(APPS_SCRIPT_URL, json=payload).json()
        return res.get("url") if res.get("status") == "success" else None
    except: return None

def terbilang(n):
    units = ["", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas"]
    if n == 0: return ""
    if n < 12: return units[n]
    if n < 20: return terbilang(n - 10) + " Belas"
    if n < 100: return units[n // 10] + " Puluh " + terbilang(n % 10)
    if n < 200: return "Seratus " + terbilang(n - 100)
    if n < 1000: return units[n // 100] + " Ratus " + terbilang(n % 100)
    if n < 1000000: return terbilang(n // 1000) + " Ribu " + terbilang(n % 1000)
    return terbilang(n // 1000000) + " Juta " + terbilang(n % 1000000)

def ui_rincian_pengajuan(d, status_text):
    st.info(f"### Rincian Pengajuan")
    st.markdown(f"""
    * **Nomor Pengajuan** : {d['id']}
    * **Tgl Pengajuan** : {d['tgl']}
    * **Dibayarkan Kepada** : {d['nama']} / {d['nip']}
    * **Departement** : {d['dept']}
    * **Senilai** : Rp {d['nom']:,} ({d['terbilang']})
    * **Keperluan** : {d['ket']}
    * **Lampiran** : [{d['link']}]({d['link']})
    * **Link Realisasi** : [{d['link_real']}]({d['link_real']})
    * **Janji Penyelesaian** : {d['janji']}
    """)
    st.write(f"**Status Saat Ini:** `{status_text}`")
    st.divider()

def check_login_portal(role_required, target_nik, assigned_str, store_code, session_key):
    if st.session_state.get(session_key): return True
    st.subheader(f"🔐 Verifikasi {role_required}")
    clean_name = assigned_str.split(" - ")[1] if " - " in assigned_str else assigned_str
    st.caption(f"Verifikasi untuk: {clean_name}")
    v_nik = st.text_input("NIK (6 Digit)", max_chars=6)
    v_pass = st.text_input("Password", type="password")
    
    if st.button("Masuk & Verifikasi", type="primary", use_container_width=True):
        if v_nik != target_nik: st.error("⛔ NIK tidak sesuai penugasan."); st.stop()
        client = get_client()
        records = client.open_by_key(SPREADSHEET_ID).worksheet("DATABASE_USER").get_all_records()
        user = next((r for r in records if str(r['NIK']).zfill(6) == v_nik and str(r['Password']) == v_pass), None)
        
        if user and (user['Role'] == role_required or role_required == "Manager"): # Flexible for double auth logic
            if str(user['Kode_Store']) != store_code: st.error("⛔ Store tidak sesuai."); st.stop()
            st.session_state[session_key] = True
            st.session_state.user_nik = str(user['NIK']).zfill(6)
            st.rerun()
        else: st.error("Login Gagal.")
    st.stop()

def success_action(msg):
    st.success("✅ Berhasil! Tugas Anda selesai.")
    st.balloons()
    time.sleep(2)
    st.rerun()

# --- INITIALIZE STATE ---
for k in ['submitted', 'data_ringkasan', 'show_errors', 'mgr_logged_in', 'cashier_real_logged_in', 'mgr_final_logged_in', 'portal_verified']:
    if k not in st.session_state: st.session_state[k] = False

# --- MAIN ROUTING ---
query_id = st.query_params.get("id")

if query_id:
    try:
        client = get_client()
        sheet = client.open_by_key(SPREADSHEET_ID).worksheet("DATA_KASBON_AZKO")
        cell = sheet.find(query_id)
        row = sheet.row_values(cell.row)
        
        # Mapping Data for Easy Access
        d = {
            'id': query_id, 'tgl': row[0], 'no': row[1], 'store': row[2], 'email_req': row[3],
            'nama': row[4], 'nip': str(row[5]), 'dept': row[6], 'nom': int(row[7]), 'terbilang': row[8],
            'ket': row[9], 'link': row[10], 'janji': row[11],
            'ass_cashier': row[12], 'ass_mgr': row[13],
            'mgr_stat': row[16] if len(row)>16 else "Pending",
            'csr_stat': row[20] if len(row)>20 else "Pending",
            'money_stat': row[24] if len(row)>24 else "Pending",
            'real_stat': row[34] if len(row)>34 else "Pending",
            'verif_real_stat': row[41] if len(row)>41 else "Pending",
            'final_stat': row[43] if len(row)>43 else "",
            'link_real': row[33] if len(row)>33 else "-"
        }
        
        # Parse Targets
        try:
            tgt_csr_nik = d['ass_cashier'].split(" - ")[0].strip()
            tgt_csr_email = d['ass_cashier'].split("(")[1].split(")")[0].strip()
            tgt_mgr_nik = d['ass_mgr'].split(" - ")[0].strip()
            tgt_mgr_email = d['ass_mgr'].split("(")[1].split(")")[0].strip()
        except: tgt_csr_nik = tgt_csr_email = tgt_mgr_nik = tgt_mgr_email = ""

        # --- LOGIC BLOCKS ---
        
        # 1. MANAGER APPROVAL
        if d['mgr_stat'] == "Pending":
            st.markdown(f'<span class="store-header">Portal Approval Manager</span>', unsafe_allow_html=True)
            if pic_email == tgt_mgr_email and d['mgr_stat'] != "Pending": st.info("Selesai."); st.stop()
            check_login_portal("Manager", tgt_mgr_nik, d['ass_mgr'], d['store'], 'mgr_logged_in')
            ui_rincian_pengajuan(d, "Waiting Manager Approval")
            
            reason = st.text_area("Alasan Reject (Jika Reject)")
            c1, c2 = st.columns(2)
            if c1.button("✓ APPROVE", use_container_width=True):
                sheet.update_cell(cell.row, 15, datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S"))
                sheet.update_cell(cell.row, 16, st.session_state.user_nik)
                sheet.update_cell(cell.row, 17, "APPROVED")
                sheet.update_cell(cell.row, 18, "-")
                
                email_body = f"""<html><body>Pengajuan Kasbon {d['id']} disetujui.<br>Klik <a href='{BASE_URL}?id={d['id']}'>Link Verifikasi</a></body></html>"""
                send_email(tgt_csr_email, f"Verifikasi Kasbon {d['id']}", email_body)
                success_action("Approved")
                
            if c2.button("✕ REJECT", use_container_width=True):
                if not reason: st.error("Isi alasan!"); st.stop()
                sheet.update_cell(cell.row, 17, "REJECTED"); sheet.update_cell(cell.row, 18, reason)
                success_action("Rejected")

        # 2. CASHIER VERIFICATION
        elif d['mgr_stat'] == "APPROVED" and d['csr_stat'] == "Pending":
            st.markdown(f'<span class="store-header">Portal Verifikasi Cashier</span>', unsafe_allow_html=True)
            if pic_email == tgt_csr_email and d['csr_stat'] != "Pending": st.info("Selesai."); st.stop()
            check_login_portal("Senior Cashier", tgt_csr_nik, d['ass_cashier'], d['store'], 'mgr_logged_in') # Re-use session key for simplicity in flow
            ui_rincian_pengajuan(d, "Waiting Cashier Verification")
            
            reason = st.text_area("Alasan Reject (Jika Reject)")
            c1, c2 = st.columns(2)
            if c1.button("✓ VERIFIKASI APPROVE", use_container_width=True):
                sheet.update_cell(cell.row, 21, "APPROVED")
                email_body = f"""<html><body>Kasbon {d['id']} disetujui & diverifikasi.<br>Klik <a href='{BASE_URL}?id={d['id']}'>Link Konfirmasi</a></body></html>"""
                send_email(d['email_req'], f"Kasbon Disetujui {d['id']}", email_body)
                success_action("Verified")
            if c2.button("✕ REJECT", use_container_width=True):
                if not reason: st.error("Isi alasan!"); st.stop()
                sheet.update_cell(cell.row, 21, "REJECTED"); sheet.update_cell(cell.row, 22, reason)
                success_action("Rejected")

        # 3. REQUESTER CONFIRMATION
        elif d['csr_stat'] == "APPROVED" and d['money_stat'] == "Pending":
            st.markdown(f'<span class="store-header">Portal Konfirmasi Uang Diterima</span>', unsafe_allow_html=True)
            if not st.session_state.portal_verified:
                st.info("🔒 Masukkan NIP & Password (6 huruf awal email login).")
                c1, c2 = st.columns(2)
                if st.button("Masuk"):
                    if c1.text_input("NIP") == d['nip'] and c2.text_input("Pass", type="password") == pic_email[:6]:
                        st.session_state.portal_verified = True; st.rerun()
                    else: st.error("Gagal.")
                st.stop()
            
            ui_rincian_pengajuan(d, "Waiting Requester Confirmation")
            if st.button("Konfirmasi uang sudah diterima", type="primary"):
                sheet.update_cell(cell.row, 23, datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S"))
                sheet.update_cell(cell.row, 24, d['nip']); sheet.update_cell(cell.row, 25, "Sudah diterima")
                success_action("Confirmed")

        # 4. REALIZATION INPUT
        elif d['money_stat'] == "Sudah diterima" and d['real_stat'] == "Pending":
            st.markdown(f'<span class="store-header">Portal Realisasi Kasbon</span>', unsafe_allow_html=True)
            # (Reuse Requester Auth Logic - Simplified for brevity in this view, assumes logic same as above)
            if not st.session_state.portal_verified:
                 # ... same auth block as step 3 ...
                 st.info("🔒 Masukkan NIP & Password."); c1, c2 = st.columns(2); nip_i = c1.text_input("NIP"); pass_i = c2.text_input("Pass", type="password")
                 if st.button("Masuk"): 
                     if nip_i == d['nip'] and pass_i == pic_email[:6]: st.session_state.portal_verified=True; st.rerun()
                     else: st.error("Gagal")
                 st.stop()

            ui_rincian_pengajuan(d, "Waiting Realization Input")
            used = st.number_input("Total Uang Digunakan", min_value=0, step=1)
            st.caption(terbilang(used).title() + " Rupiah")
            
            selisih = d['nom'] - used
            c1, c2 = st.columns(2)
            kembali = selisih if used < d['nom'] else 0
            terima = abs(selisih) if used > d['nom'] else 0
            
            c1.text_input("Kembali ke Perusahaan", value=f"Rp {kembali:,}", disabled=True)
            c2.text_input("Reimburse ke Karyawan", value=f"Rp {terima:,}", disabled=True)
            
            opsi = st.radio("Bukti:", ["Upload", "Kamera"])
            bukti = st.file_uploader("File") if opsi == "Upload" else st.camera_input("Foto")
            
            if st.button("Kirim Laporan"):
                if bukti and used >= 0:
                    link_b = upload_file(bukti, f"Realisasi_{d['id']}.{bukti.type.split('/')[-1]}")
                    if not link_b: st.error("Gagal Upload"); st.stop()
                    
                    # Update Cells Z(26) to AI(35)
                    vals = [datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S"), d['nip'], used, terbilang(used), kembali, terbilang(kembali), terima, terbilang(terima), link_b, "Terrealisasi"]
                    # Gspread batch update or single cell update loop
                    for i, val in enumerate(vals): sheet.update_cell(cell.row, 26+i, val)
                    success_action("Realization Submitted")
                else: st.error("Lengkapi data.")

        # 5. CASHIER REALIZATION VERIF
        elif d['real_stat'] == "Terrealisasi" and d['verif_real_stat'] == "Pending":
            st.markdown(f'<span class="store-header">Portal Verif Realisasi</span>', unsafe_allow_html=True)
            check_login_portal("Senior Cashier", tgt_csr_nik, d['ass_cashier'], d['store'], 'cashier_real_logged_in')
            ui_rincian_pengajuan(d, "Waiting Cashier Realization Verif")
            
            status_pilihan = st.radio("Status Realisasi", ["Ya, Sesuai", "Tidak Sesuai"])
            reason = st.text_area("Reason") if status_pilihan == "Tidak Sesuai" else "-"
            
            # Read DB values
            db_kembali = int(row[29]) if len(row)>29 and row[29] else 0
            db_terima = int(row[31]) if len(row)>31 and row[31] else 0
            
            c1, c2 = st.columns(2)
            u_kembali = c1.number_input("Verif Kembali", value=db_kembali, disabled=(status_pilihan=="Ya, Sesuai"))
            u_terima = c2.number_input("Verif Terima", value=db_terima, disabled=(status_pilihan=="Ya, Sesuai"))
            
            if st.button("Submit"):
                if status_pilihan == "Tidak Sesuai" and not reason: st.error("Isi Reason"); st.stop()
                # Update AJ(36) to AQ(43)
                vals = [datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S"), st.session_state.user_nik, u_kembali, terbilang(u_kembali), u_terima, terbilang(u_terima), status_pilihan, reason]
                for i, val in enumerate(vals): sheet.update_cell(cell.row, 36+i, val)
                
                email_body = f"""<html><body>Final Cek Kasbon {d['id']}. <a href='{BASE_URL}?id={d['id']}'>Link Final Cek</a></body></html>"""
                send_email(tgt_mgr_email, f"Final Cek {d['id']}", email_body)
                success_action("Verified")

        # 6. MANAGER FINAL CEK
        elif d['verif_real_stat'] in ["Ya, Sesuai", "Tidak Sesuai"] and d['final_stat'] == "":
            st.markdown(f'<span class="store-header">Portal Final Cek</span>', unsafe_allow_html=True)
            check_login_portal("Manager", tgt_mgr_nik, d['ass_mgr'], d['store'], 'mgr_final_logged_in')
            ui_rincian_pengajuan(d, "Waiting Manager Final Check")
            
            q1 = st.radio("1. Foto Nota Sesuai?", ["Ya, Sesuai", "Tidak Sesuai"], key="q1")
            r1 = st.text_input("Reason Q1") if q1 == "Tidak Sesuai" else "-"
            q2 = st.radio("2. Foto Item Sesuai?", ["Ya, Sesuai", "Tidak Sesuai"], key="q2")
            r2 = st.text_input("Reason Q2") if q2 == "Tidak Sesuai" else "-"
            
            if st.button("Posting"):
                if "Tidak Sesuai" in [q1, q2] and ("-" in [r1, r2] or "" in [r1, r2]): st.error("Isi Reason"); st.stop()
                vals = [datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S"), st.session_state.user_nik, q1, r1, q2, r2]
                for i, val in enumerate(vals): sheet.update_cell(cell.row, 44+i, val) # AR start
                success_action("Completed")
        
        # 7. COMPLETED / REJECTED
        else:
            status = "REJECTED" if "REJECTED" in [d['mgr_stat'], d['csr_stat']] else "COMPLETED"
            st.markdown(f'<span class="store-header">Status: {status}</span>', unsafe_allow_html=True)
            ui_rincian_pengajuan(d, status)

    except Exception as e: st.error(f"Error: {e}"); st.stop()

# --- FORM SUBMISSION ---
elif st.session_state.submitted:
    st.success("## ✅ PENGAJUAN TERKIRIM")
    d = st.session_state.data_ringkasan
    ui_rincian_pengajuan(d, "Submitted to Manager")
    if st.button("Baru"): st.session_state.submitted = False; st.rerun()

else:
    st.caption(f"Logged: {pic_email}")
    st.subheader("📍 Form Pengajuan")
    k_store = st.text_input("Kode Store").upper()
    
    if k_store:
        try:
            client = get_client()
            users = client.open_by_key(SPREADSHEET_ID).worksheet("DATABASE_USER").get_all_records()
            store = next((u for u in users if str(u['Kode_Store']) == k_store), None)
            if not store: st.error("Store Invalid"); st.stop()
            
            st.markdown(f'<span class="store-header">{store["Nama_Store"]}</span>', unsafe_allow_html=True)
            
            # Form Inputs
            nama = st.text_input("Dibayarkan Kepada")
            nip = st.text_input("NIP (6 Digit)", max_chars=6)
            dept = st.selectbox("Departemen", ["-", "Operational", "Sales", "Inventory", "HR", "Other"])
            nom = st.text_input("Nominal")
            if nom.isdigit(): st.caption(terbilang(int(nom)).title() + " Rupiah")
            ket = st.text_input("Keperluan")
            janji = st.date_input("Janji Penyelesaian", min_value=datetime.date.today())
            
            opsi = st.radio("Lampiran:", ["Upload", "Kamera"])
            bukti = st.file_uploader("File") if opsi == "Upload" else st.camera_input("Foto")
            
            mgrs = [f"{u['NIK']} - {u['Nama Lengkap']} ({u['Email']})" for u in users if str(u['Kode_Store']) == k_store and u['Role'] == 'Manager']
            csrs = [f"{u['NIK']} - {u['Nama Lengkap']} ({u['Email']})" for u in users if str(u['Kode_Store']) == k_store and u['Role'] == 'Senior Cashier']
            
            mgr_f = st.selectbox("Manager", ["-"] + mgrs)
            csr_f = st.selectbox("Cashier", ["-"] + csrs)
            
            if st.button("Kirim Pengajuan", type="primary"):
                if nama and len(nip)==6 and nom.isdigit() and bukti and mgr_f!="-" and csr_f!="-":
                    with st.spinner("Processing..."):
                        sheet = client.open_by_key(SPREADSHEET_ID).worksheet("DATA_KASBON_AZKO")
                        prefix = f"KB{k_store}-{datetime.datetime.now().strftime('%m%y')}-"
                        last = max([int(r[1].split("-")[-1]) for r in sheet.get_all_values() if len(r)>1 and r[1].startswith(prefix)] or [0])
                        no_p = f"{prefix}{str(last+1).zfill(3)}"
                        
                        link = upload_file(bukti, f"Lampiran_{no_p}.{bukti.type.split('/')[-1]}")
                        if not link: st.error("Upload Error"); st.stop()
                        
                        terbilang_txt = terbilang(int(nom)).title() + " Rupiah"
                        row_data = [
                            datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S"), no_p, k_store, pic_email,
                            nama, nip, dept, nom, terbilang_txt, ket, link, janji.strftime("%d/%m/%Y"), csr_f, mgr_f
                        ] + [""]*3 + ["Pending"] + [""]*3 + ["Pending"] + [""]*3 + ["Pending"] + [""]*10 + ["Pending"] + [""]*7 + [""]*6
                        
                        sheet.append_row(row_data)
                        
                        email_mgr = mgr_f.split("(")[1].split(")")[0]
                        email_body = f"""<html><body>Mohon Approval Kasbon {no_p}. <br><a href='{BASE_URL}?id={no_p}'>Link Approval</a></body></html>"""
                        send_email(email_mgr, f"Pengajuan Kasbon {no_p}", email_body)
                        
                        st.session_state.data_ringkasan = {'id': no_p, 'tgl': row_data[0], 'nama': nama, 'nip': nip, 'dept': dept, 'nom': int(nom), 'terbilang': terbilang_txt, 'ket': ket, 'link': link, 'link_real': '-', 'janji': row_data[11]}
                        st.session_state.submitted = True; st.rerun()
                else: st.error("Lengkapi Data!")
        except Exception as e: st.error(f"System Error: {e}")