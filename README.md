# 🏢 Digital Kasbon Petty Cash — Enterprise Modular Suite

Sistem manajemen siklus **Kasbon Digital Petty Cash (Step 1 s/d Step 7)** berstandar korporat dengan arsitektur **Multi-Project Google Apps Script (Modular)** yang terhubung langsung ke **Supabase Cloud (PostgreSQL Database & Object Storage)** tanpa ketergantungan pada notifikasi email SMTP.

---

## 🏗️ 1. Arsitektur & Struktur Direktori

```text
Digital-Kasbon/
├── supabase/                        # Database DDL, Storage Setup & Migrasi Data
│   ├── 01_schema_tables.sql         # Tabel kasbon_stores, kasbon_users, kasbon_requests, kasbon_audit_logs & View v_kasbon_dashboard
│   ├── 02_storage_setup.sql         # Konfigurasi Bucket 'kasbon-attachments' (Maks 10MB, Public Read)
│   └── migrate_from_sheets.py       # Script provisioning otomatis & migrasi data dari Google Sheets ke Supabase
│
├── gas-common/                      # Single Source of Truth (Shared Core Modules)
│   ├── Config.gs                    # Konfigurasi terpusat (Supabase URL/Key & Cross-Portal URLs via ScriptProperties)
│   ├── SupabaseClient.gs            # Lightweight PostgREST HTTP Client (select, insert, update, logAudit)
│   ├── StorageService.gs            # Direct upload bukti lampiran ke Supabase Storage ('kasbon-attachments')
│   ├── AuthService.gs               # Validasi otorisasi Role (Manager, Senior Cashier) & Double-Auth Pemohon
│   ├── KasbonHelper.gs              # Generator ID sekuensial (KB{Store}-{MMYY}-{Seq}), Terbilang Rupiah, Timezone WIB
│   ├── SharedStyles.html            # Enterprise Design System + validasi form modern (:user-invalid)
│   └── SharedUtils.html             # Client-side Terbilang, kompresi gambar Canvas HTML5, & sinkronisasi aria-invalid
│
├── gas-dashboard/                   # PROJECT 1: Control Center & Real-Time Monitoring Dashboard
├── gas-form-pengajuan/              # PROJECT 2: Form Pengajuan Kasbon Baru (Step 1)
├── gas-portal-manager/              # PROJECT 3: Portal Otorisasi Manager (Step 2 Approval & Step 7 Final Cek)
├── gas-portal-cashier/              # PROJECT 4: Portal Verifikasi Senior Cashier (Step 3 Pencairan & Step 6 Realisasi)
├── gas-portal-requester/            # PROJECT 5: Portal Pemohon (Step 4 Terima Uang & Step 5 Laporan Realisasi)
│
├── scripts/
│   ├── sync_gas_common.py           # Sinkronisasi otomatis modul gas-common ke seluruh 5 folder project GAS
│   └── verify_e2e_supabase.py       # Pengujian otomatis End-to-End 7 tahapan kasbon & upload Storage di Supabase
│
└── app.py                           # Aplikasi Streamlit (Legacy v1 - telah diperbaiki & dioptimasi)
```

---

## 🔄 2. Alur Siklus State Machine Kasbon (Step 1 – Step 7)

Seluruh perubahan status dikawal secara otomatis oleh **PostgreSQL Trigger (`trg_sync_kasbon_state`)** dan ditampilkan secara *real-time* melalui view **`v_kasbon_dashboard`**:

| Step | Tahapan Proses | Pelaksana | Status Sistem (`overall_status`) | Project Web App |
| :---: | :--- | :--- | :--- | :--- |
| **1** | Input Pengajuan Kasbon Baru | Requester | `WAITING_MGR_APPROVAL` | `gas-form-pengajuan` |
| **2** | Approval Pengajuan Kasbon | Manager | `WAITING_CASHIER_VERIF` / `REJECTED_BY_MGR` | `gas-portal-manager` |
| **3** | Verifikasi & Pencairan Dana | Senior Cashier | `WAITING_MONEY_RECEIPT` / `REJECTED_BY_CASHIER` | `gas-portal-cashier` |
| **4** | Konfirmasi Uang Fisik Diterima | Requester | `WAITING_REALIZATION_INPUT` | `gas-portal-requester` |
| **5** | Submit Laporan & Nota Realisasi | Requester | `WAITING_CASHIER_REA_VERIF` | `gas-portal-requester` |
| **6** | Verifikasi Fisik Uang & Nota Realisasi | Senior Cashier | `WAITING_MGR_FINAL_CHECK` | `gas-portal-cashier` |
| **7** | Final Cek Kesesuaian Nota & Item | Manager | `COMPLETED` | `gas-portal-manager` |

---

## 🚀 3. Panduan Deployment Multi-Project Google Apps Script

Setiap folder (`gas-dashboard`, `gas-form-pengajuan`, `gas-portal-manager`, `gas-portal-cashier`, `gas-portal-requester`) telah disinkronisasi agar **mandiri (*self-contained*)** dan siap di-deploy sebagai Web App terpisah:

### Opsi A: Menggunakan Google `clasp` CLI
1. Masuk ke masing-masing direktori project (contoh: `gas-dashboard`).
2. Jalankan `clasp create --type webapp --title "Kasbon - Dashboard"` (atau hubungkan ke `scriptId` yang sudah ada di `.clasp.json`).
3. Jalankan `clasp push` lalu `clasp deploy`.

### Opsi B: Menghubungkan URL Antar-Project
1. Setelah kelima Web App di-deploy dan menghasilkan URL `https://script.google.com/macros/s/.../exec`, buka **Web App Dashboard (`gas-dashboard`)**.
2. Klik tombol **`⚙️ Konfigurasi URL Form`** di pojok kanan atas navbar.
3. Tempelkan URL masing-masing Web App lalu klik **Simpan Konfigurasi** (atau simpan melalui `ScriptProperties` pada masing-masing project dengan key: `URL_DASHBOARD`, `URL_FORM_PENGAJUAN`, `URL_PORTAL_MANAGER`, `URL_PORTAL_CASHIER`, `URL_PORTAL_REQUESTER`).

---

## 🛠️ 4. Utilitas Sinkronisasi & Pengujian

* **Sinkronisasi Modul Bersama (`gas-common/`) ke 5 Project**:
  ```bash
  ./venv/Scripts/python scripts/sync_gas_common.py
  ```
* **Menjalankan Migrasi Skema & Data dari Google Sheets ke Supabase**:
  ```bash
  ./venv/Scripts/python supabase/migrate_from_sheets.py
  ```
* **Menjalankan Verifikasi End-to-End (Step 1 s/d Step 7 + Storage)**:
  ```bash
  ./venv/Scripts/python scripts/verify_e2e_supabase.py
  ```