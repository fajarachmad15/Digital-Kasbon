-- =============================================================================
-- DIGITAL KASBON PETTY CASH - ENTERPRISE SUPABASE SCHEMA (DDL)
-- Project Ref : uvjtzkpyunfuzcbhubpg
-- Description : Core tables, state-machine trigger, audit trail, and views
-- =============================================================================

-- 1. MASTER STORE TABLE
CREATE TABLE IF NOT EXISTS public.kasbon_stores (
    kode_store VARCHAR(16) PRIMARY KEY,
    nama_store VARCHAR(150) NOT NULL,
    is_active BOOLEAN NOT NULL DEFAULT TRUE,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 2. MASTER USERS / PIC TABLE
CREATE TABLE IF NOT EXISTS public.kasbon_users (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    nik VARCHAR(10) NOT NULL,
    nama_lengkap VARCHAR(150) NOT NULL,
    email VARCHAR(150) NOT NULL,
    role VARCHAR(50) NOT NULL CHECK (role IN ('Manager', 'Senior Cashier', 'Requester', 'Admin')),
    kode_store VARCHAR(16) NOT NULL REFERENCES public.kasbon_stores(kode_store) ON UPDATE CASCADE,
    password_pin VARCHAR(100) NOT NULL DEFAULT '123123',
    is_active BOOLEAN NOT NULL DEFAULT TRUE,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    CONSTRAINT uq_kasbon_users_nik_store_role UNIQUE (nik, kode_store, role)
);

CREATE INDEX IF NOT EXISTS idx_kasbon_users_store_role ON public.kasbon_users(kode_store, role);
CREATE INDEX IF NOT EXISTS idx_kasbon_users_email ON public.kasbon_users(LOWER(email));
CREATE INDEX IF NOT EXISTS idx_kasbon_users_nik ON public.kasbon_users(nik);

-- 3. MAIN KASBON TRANSACTIONS TABLE
CREATE TABLE IF NOT EXISTS public.kasbon_requests (
    no_pengajuan VARCHAR(32) PRIMARY KEY,
    kode_store VARCHAR(16) NOT NULL REFERENCES public.kasbon_stores(kode_store) ON UPDATE CASCADE,
    
    -- Step 1: Request Information
    timestamp_req TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    email_request VARCHAR(150) NOT NULL,
    dibayarkan_kepada VARCHAR(150) NOT NULL,
    nip_req VARCHAR(10) NOT NULL,
    departemen VARCHAR(80) NOT NULL,
    nominal BIGINT NOT NULL CHECK (nominal >= 0),
    terbilang_req TEXT NOT NULL,
    keperluan TEXT NOT NULL,
    bukti_lampiran_req TEXT DEFAULT '-',
    janji_penyelesaian VARCHAR(30) NOT NULL,
    
    -- Assigned Officers
    senior_cashier VARCHAR(255) NOT NULL,
    senior_cashier_nik VARCHAR(10),
    senior_cashier_name VARCHAR(150),
    senior_cashier_email VARCHAR(150),
    manager_incharge VARCHAR(255) NOT NULL,
    manager_incharge_nik VARCHAR(10),
    manager_incharge_name VARCHAR(150),
    manager_incharge_email VARCHAR(150),
    
    -- Step 2: Manager Approval
    timestamp_app TIMESTAMPTZ,
    nip_app VARCHAR(10),
    status_approval VARCHAR(30) NOT NULL DEFAULT 'Pending', -- 'Pending', 'Approved', 'Reject'
    reason_reject_app TEXT DEFAULT '-',
    
    -- Step 3: Senior Cashier Verification (Pencairan)
    timestamp_ver TIMESTAMPTZ,
    nip_ver VARCHAR(10),
    status_verifikasi VARCHAR(40) NOT NULL DEFAULT 'Pending', -- 'Pending', 'Verifikasi Approved', 'Verifikasi Reject'
    reason_reject_ver TEXT DEFAULT '-',
    
    -- Step 4: Requester Money Receipt Confirmation
    timestamp_rec TIMESTAMPTZ,
    nip_rec VARCHAR(10),
    status_uang VARCHAR(40) NOT NULL DEFAULT 'Pending', -- 'Pending', 'Sudah diterima'
    
    -- Step 5: Requester Realization Report
    timestamp_rea TIMESTAMPTZ,
    nip_rea VARCHAR(10),
    nominal_pembelanjaan_rea BIGINT DEFAULT 0,
    terbilang_rea1 TEXT DEFAULT '',
    uang_dikembalikan_rea BIGINT DEFAULT 0,
    terbilang_rea2 TEXT DEFAULT '',
    uang_diterima_rea BIGINT DEFAULT 0,
    terbilang_rea3 TEXT DEFAULT '',
    bukti_lampiran_rea TEXT DEFAULT '-',
    status_realisasi VARCHAR(40) NOT NULL DEFAULT 'Pending', -- 'Pending', 'Terrealisasi'
    
    -- Step 6: Senior Cashier Realization Verification
    timestamp_ver_rea TIMESTAMPTZ,
    nip_ver_rea VARCHAR(10),
    uang_dikembalikan_ver_rea BIGINT DEFAULT 0,
    terbilang_ver_rea1 TEXT DEFAULT '',
    uang_diterima_ver_rea BIGINT DEFAULT 0,
    terbilang_ver_rea2 TEXT DEFAULT '',
    status_verifikasi_rea VARCHAR(40) NOT NULL DEFAULT 'Pending', -- 'Pending', 'Ya, Sesuai', 'Tidak Sesuai'
    reason_ver_rea TEXT DEFAULT '-',
    
    -- Step 7: Manager Final Check
    timestamp_fin_cek TIMESTAMPTZ,
    nip_fin_cek VARCHAR(10),
    status_qty_value_nota VARCHAR(40) DEFAULT '', -- 'Ya, Sesuai', 'Tidak Sesuai'
    reason_fin_cek_1 TEXT DEFAULT '-',
    status_item VARCHAR(40) DEFAULT '', -- 'Ya, Sesuai', 'Tidak Sesuai'
    reason_fin_cek_2 TEXT DEFAULT '-',
    revisi_bukti_realisasi TEXT DEFAULT '-',
    
    -- Computed State Machine Columns (Maintained automatically by trigger)
    current_step SMALLINT NOT NULL DEFAULT 1,
    overall_status VARCHAR(50) NOT NULL DEFAULT 'WAITING_MGR_APPROVAL',
    
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_kasbon_requests_store ON public.kasbon_requests(kode_store);
CREATE INDEX IF NOT EXISTS idx_kasbon_requests_status ON public.kasbon_requests(overall_status);
CREATE INDEX IF NOT EXISTS idx_kasbon_requests_step ON public.kasbon_requests(current_step);
CREATE INDEX IF NOT EXISTS idx_kasbon_requests_ts ON public.kasbon_requests(timestamp_req DESC);
CREATE INDEX IF NOT EXISTS idx_kasbon_requests_mgr_email ON public.kasbon_requests(LOWER(manager_incharge_email));
CREATE INDEX IF NOT EXISTS idx_kasbon_requests_csr_email ON public.kasbon_requests(LOWER(senior_cashier_email));
CREATE INDEX IF NOT EXISTS idx_kasbon_requests_req_email ON public.kasbon_requests(LOWER(email_request));

-- 4. AUDIT TRAIL TABLE (CORPORATE GOVERNANCE)
CREATE TABLE IF NOT EXISTS public.kasbon_audit_logs (
    id BIGSERIAL PRIMARY KEY,
    no_pengajuan VARCHAR(32) NOT NULL REFERENCES public.kasbon_requests(no_pengajuan) ON DELETE CASCADE,
    step_number SMALLINT NOT NULL,
    action_type VARCHAR(60) NOT NULL,
    actor_nik VARCHAR(10),
    actor_name VARCHAR(150),
    actor_email VARCHAR(150),
    actor_role VARCHAR(50),
    previous_status VARCHAR(50),
    new_status VARCHAR(50),
    notes TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_kasbon_audit_no_pengajuan ON public.kasbon_audit_logs(no_pengajuan, created_at DESC);

-- 5. STATE MACHINE TRIGGER FUNCTION
CREATE OR REPLACE FUNCTION public.fn_sync_kasbon_state()
RETURNS TRIGGER
LANGUAGE plpgsql
AS $$
BEGIN
    NEW.updated_at := NOW();
    
    -- Normalize NIP/NIK to 6 digits if numeric
    IF NEW.nip_req IS NOT NULL AND NEW.nip_req ~ '^[0-9]+$' AND LENGTH(NEW.nip_req) < 6 THEN
        NEW.nip_req := LPAD(NEW.nip_req, 6, '0');
    END IF;

    -- Parse Senior Cashier string if sub-fields are empty: "129137 - Name (email@domain.com)"
    IF (NEW.senior_cashier_nik IS NULL OR NEW.senior_cashier_nik = '') AND NEW.senior_cashier LIKE '% - %(%)' THEN
        NEW.senior_cashier_nik := LPAD(TRIM(SPLIT_PART(NEW.senior_cashier, ' - ', 1)), 6, '0');
        NEW.senior_cashier_name := TRIM(SPLIT_PART(SPLIT_PART(NEW.senior_cashier, ' - ', 2), ' (', 1));
        NEW.senior_cashier_email := LOWER(TRIM(SPLIT_PART(SPLIT_PART(NEW.senior_cashier, '(', 2), ')', 1)));
    END IF;

    -- Parse Manager Incharge string if sub-fields are empty
    IF (NEW.manager_incharge_nik IS NULL OR NEW.manager_incharge_nik = '') AND NEW.manager_incharge LIKE '% - %(%)' THEN
        NEW.manager_incharge_nik := LPAD(TRIM(SPLIT_PART(NEW.manager_incharge, ' - ', 1)), 6, '0');
        NEW.manager_incharge_name := TRIM(SPLIT_PART(SPLIT_PART(NEW.manager_incharge, ' - ', 2), ' (', 1));
        NEW.manager_incharge_email := LOWER(TRIM(SPLIT_PART(SPLIT_PART(NEW.manager_incharge, '(', 2), ')', 1)));
    END IF;

    -- State Machine Evaluation (Step 1 -> Step 7)
    IF NEW.status_approval = 'Reject' THEN
        NEW.current_step := 2;
        NEW.overall_status := 'REJECTED_BY_MGR';
    ELSIF COALESCE(NEW.status_approval, '') IN ('', 'Pending') THEN
        NEW.current_step := 2;
        NEW.overall_status := 'WAITING_MGR_APPROVAL';
    ELSIF NEW.status_verifikasi = 'Verifikasi Reject' THEN
        NEW.current_step := 3;
        NEW.overall_status := 'REJECTED_BY_CASHIER';
    ELSIF COALESCE(NEW.status_verifikasi, '') IN ('', 'Pending') THEN
        NEW.current_step := 3;
        NEW.overall_status := 'WAITING_CASHIER_VERIF';
    ELSIF COALESCE(NEW.status_uang, '') IN ('', 'Pending') THEN
        NEW.current_step := 4;
        NEW.overall_status := 'WAITING_MONEY_RECEIPT';
    ELSIF COALESCE(NEW.status_realisasi, '') IN ('', 'Pending') THEN
        NEW.current_step := 5;
        NEW.overall_status := 'WAITING_REALIZATION_INPUT';
    ELSIF COALESCE(NEW.status_verifikasi_rea, '') IN ('', 'Pending') THEN
        NEW.current_step := 6;
        NEW.overall_status := 'WAITING_CASHIER_REA_VERIF';
    ELSIF NEW.timestamp_fin_cek IS NULL THEN
        NEW.current_step := 7;
        NEW.overall_status := 'WAITING_MGR_FINAL_CHECK';
    ELSE
        NEW.current_step := 7;
        NEW.overall_status := 'COMPLETED';
    END IF;

    RETURN NEW;
END;
$$;

DROP TRIGGER IF EXISTS trg_sync_kasbon_state ON public.kasbon_requests;
CREATE TRIGGER trg_sync_kasbon_state
BEFORE INSERT OR UPDATE ON public.kasbon_requests
FOR EACH ROW
EXECUTE FUNCTION public.fn_sync_kasbon_state();

-- 6. DASHBOARD ANALYTICS & MONITORING VIEW
CREATE OR REPLACE VIEW public.v_kasbon_dashboard AS
SELECT
    k.no_pengajuan,
    k.kode_store,
    COALESCE(s.nama_store, k.kode_store) AS nama_store,
    k.timestamp_req,
    k.email_request,
    k.dibayarkan_kepada,
    k.nip_req,
    k.departemen,
    k.nominal,
    k.terbilang_req,
    k.keperluan,
    k.bukti_lampiran_req,
    k.janji_penyelesaian,
    k.senior_cashier,
    k.senior_cashier_nik,
    k.senior_cashier_name,
    k.senior_cashier_email,
    k.manager_incharge,
    k.manager_incharge_nik,
    k.manager_incharge_name,
    k.manager_incharge_email,
    k.status_approval,
    k.timestamp_app,
    k.reason_reject_app,
    k.status_verifikasi,
    k.timestamp_ver,
    k.reason_reject_ver,
    k.status_uang,
    k.timestamp_rec,
    k.status_realisasi,
    k.timestamp_rea,
    k.nominal_pembelanjaan_rea,
    k.uang_dikembalikan_rea,
    k.uang_diterima_rea,
    k.bukti_lampiran_rea,
    k.status_verifikasi_rea,
    k.timestamp_ver_rea,
    k.uang_dikembalikan_ver_rea,
    k.uang_diterima_ver_rea,
    k.reason_ver_rea,
    k.timestamp_fin_cek,
    k.status_qty_value_nota,
    k.reason_fin_cek_1,
    k.status_item,
    k.reason_fin_cek_2,
    k.revisi_bukti_realisasi,
    k.current_step,
    k.overall_status,
    CASE k.overall_status
        WHEN 'WAITING_MGR_APPROVAL' THEN 'Step 2: Menunggu Approval Manager (' || COALESCE(k.manager_incharge_name, 'Manager') || ')'
        WHEN 'REJECTED_BY_MGR' THEN 'Ditolak oleh Manager (' || COALESCE(k.reason_reject_app, '-') || ')'
        WHEN 'WAITING_CASHIER_VERIF' THEN 'Step 3: Menunggu Verifikasi Cashier (' || COALESCE(k.senior_cashier_name, 'Cashier') || ')'
        WHEN 'REJECTED_BY_CASHIER' THEN 'Ditolak oleh Cashier (' || COALESCE(k.reason_reject_ver, '-') || ')'
        WHEN 'WAITING_MONEY_RECEIPT' THEN 'Step 4: Menunggu Konfirmasi Terima Uang (' || k.dibayarkan_kepada || ')'
        WHEN 'WAITING_REALIZATION_INPUT' THEN 'Step 5: Uang Diterima - Menunggu Laporan Realisasi (' || k.dibayarkan_kepada || ')'
        WHEN 'WAITING_CASHIER_REA_VERIF' THEN 'Step 6: Menunggu Verifikasi Realisasi Cashier (' || COALESCE(k.senior_cashier_name, 'Cashier') || ')'
        WHEN 'WAITING_MGR_FINAL_CHECK' THEN 'Step 7: Menunggu Final Cek Manager (' || COALESCE(k.manager_incharge_name, 'Manager') || ')'
        WHEN 'COMPLETED' THEN 'Selesai (Completed)'
        ELSE k.overall_status
    END AS status_label,
    CASE k.overall_status
        WHEN 'WAITING_MGR_APPROVAL' THEN 'Manager'
        WHEN 'WAITING_CASHIER_VERIF' THEN 'Senior Cashier'
        WHEN 'WAITING_MONEY_RECEIPT' THEN 'Requester'
        WHEN 'WAITING_REALIZATION_INPUT' THEN 'Requester'
        WHEN 'WAITING_CASHIER_REA_VERIF' THEN 'Senior Cashier'
        WHEN 'WAITING_MGR_FINAL_CHECK' THEN 'Manager'
        ELSE 'None'
    END AS next_action_role,
    CASE k.overall_status
        WHEN 'WAITING_MGR_APPROVAL' THEN 'portal-manager'
        WHEN 'WAITING_CASHIER_VERIF' THEN 'portal-cashier'
        WHEN 'WAITING_MONEY_RECEIPT' THEN 'portal-requester'
        WHEN 'WAITING_REALIZATION_INPUT' THEN 'portal-requester'
        WHEN 'WAITING_CASHIER_REA_VERIF' THEN 'portal-cashier'
        WHEN 'WAITING_MGR_FINAL_CHECK' THEN 'portal-manager'
        ELSE 'completed'
    END AS target_portal,
    (k.overall_status IN ('REJECTED_BY_MGR', 'REJECTED_BY_CASHIER')) AS is_rejected,
    (k.overall_status = 'COMPLETED') AS is_completed,
    GREATEST(0, EXTRACT(DAY FROM (COALESCE(k.timestamp_fin_cek, NOW()) - k.timestamp_req))::INT) AS aging_days,
    k.updated_at
FROM public.kasbon_requests k
LEFT JOIN public.kasbon_stores s ON s.kode_store = k.kode_store;
