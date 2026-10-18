-- =============================================================================
-- DIGITAL KASBON PETTY CASH - SUPABASE STORAGE BUCKET & RLS POLICIES
-- Bucket ID : kasbon-attachments
-- =============================================================================

INSERT INTO storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
VALUES (
    'kasbon-attachments',
    'kasbon-attachments',
    TRUE,
    10485760, -- 10 MB max file size
    ARRAY['image/jpeg', 'image/png', 'image/webp', 'application/pdf']
)
ON CONFLICT (id) DO UPDATE SET
    public = EXCLUDED.public,
    file_size_limit = EXCLUDED.file_size_limit,
    allowed_mime_types = EXCLUDED.allowed_mime_types;

-- Allow public read access to attachments so dashboard/portals can preview receipts directly
DO $$
BEGIN
    IF NOT EXISTS (
        SELECT 1 FROM pg_policies 
        WHERE schemaname = 'storage' AND tablename = 'objects' AND policyname = 'kasbon_attachments_public_read'
    ) THEN
        CREATE POLICY kasbon_attachments_public_read
        ON storage.objects FOR SELECT
        USING (bucket_id = 'kasbon-attachments');
    END IF;
END
$$;
