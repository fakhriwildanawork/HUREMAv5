-- SKRIP SETUP MODUL LIBUR MANDIRI
-- Jalankan skrip ini di SQL Editor Supabase Anda

-- 1. Buat Tabel account_leave_requests
CREATE TABLE IF NOT EXISTS public.account_leave_requests (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    account_id UUID NOT NULL REFERENCES public.accounts(id) ON DELETE CASCADE,
    start_date DATE NOT NULL,
    end_date DATE NOT NULL,
    description TEXT,
    status TEXT DEFAULT 'pending', -- pending, approved, rejected
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- 2. Aktifkan RLS
ALTER TABLE public.account_leave_requests ENABLE ROW LEVEL SECURITY;

-- 3. Kebijakan RLS (Public Upsert untuk kemudahan modul)
CREATE POLICY "Allow public access for leave requests" ON public.account_leave_requests
    FOR ALL USING (true) WITH CHECK (true);

-- 4. Tambahkan Setting Default di app_settings
INSERT INTO public.app_settings (key, value, description)
VALUES 
    ('leave_approval_policy', '"manual"', 'Kebijakan pengajuan libur mandiri: manual atau auto')
ON CONFLICT (key) DO NOTHING;

-- Komentar tabel
COMMENT ON TABLE public.account_leave_requests IS 'Tabel untuk menyimpan data pengajuan libur mandiri karyawan';
