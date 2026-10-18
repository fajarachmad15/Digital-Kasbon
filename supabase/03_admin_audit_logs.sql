-- =============================================================================
-- DIGITAL KASBON PETTY CASH - ADMIN AUDIT LOGS & INITIAL SUPER ADMIN SEED
-- Project Ref : uvjtzkpyunfuzcbhubpg
-- Description : Audit trail for Master Store & User/Role management + Admin seed
-- =============================================================================

-- 1. ADMIN AUDIT LOGS TABLE
CREATE TABLE IF NOT EXISTS public.kasbon_admin_logs (
    id BIGSERIAL PRIMARY KEY,
    action_type VARCHAR(60) NOT NULL,
    target_entity VARCHAR(50) NOT NULL,
    target_id VARCHAR(100) NOT NULL,
    actor_nik VARCHAR(10) NOT NULL,
    actor_name VARCHAR(150) NOT NULL,
    actor_email VARCHAR(150) NOT NULL,
    details TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_kasbon_admin_logs_ts ON public.kasbon_admin_logs(created_at DESC);

-- 2. SEED DEFAULT SUPER ADMIN ACCOUNT (Achmad Fajar - NIK 121014)
INSERT INTO public.kasbon_users (nik, nama_lengkap, email, role, kode_store, password_pin, is_active)
VALUES ('121014', 'Achmad Fajar', 'fajarachmad15@gmail.com', 'Admin', 'A300', '123123', TRUE)
ON CONFLICT (nik, kode_store, role) DO UPDATE
SET
    nama_lengkap = EXCLUDED.nama_lengkap,
    email = EXCLUDED.email,
    is_active = TRUE,
    updated_at = NOW();
