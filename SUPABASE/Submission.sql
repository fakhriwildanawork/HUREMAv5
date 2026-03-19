-- Table for Account Submissions (Workflow)
CREATE TABLE IF NOT EXISTS account_submissions (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    account_id UUID REFERENCES accounts(id) ON DELETE CASCADE,
    type TEXT NOT NULL, -- 'Lembur', 'Cuti', 'Izin', 'Koreksi Absen'
    status TEXT DEFAULT 'Pending' CHECK (status IN ('Pending', 'Disetujui', 'Ditolak', 'Dibatalkan')),
    submission_data JSONB NOT NULL, -- Flexible data based on type
    description TEXT,
    verifier_id UUID REFERENCES accounts(id) ON DELETE SET NULL,
    verified_at TIMESTAMP WITH TIME ZONE,
    verification_notes TEXT,
    file_id TEXT, -- Attachment from Google Drive
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Enable RLS
ALTER TABLE account_submissions ENABLE ROW LEVEL SECURITY;

-- Policies
-- 1. Karyawan bisa melihat pengajuannya sendiri
CREATE POLICY "Allow users to view own submissions" 
ON account_submissions FOR SELECT 
TO public 
USING (account_id = (SELECT id FROM accounts WHERE access_code = current_setting('request.jwt.claims', true)::json->>'access_code' LIMIT 1) OR true); 
-- Catatan: 'OR true' digunakan sementara jika sistem auth belum menggunakan JWT Claims Supabase secara penuh, 
-- namun logic idealnya adalah menggunakan auth.uid()

-- 2. Akses publik untuk operasional (karena HUREMA saat ini menggunakan anon key)
DROP POLICY IF EXISTS "Allow public read submissions" ON account_submissions;
CREATE POLICY "Allow public read submissions" ON account_submissions FOR SELECT TO public USING (true);

DROP POLICY IF EXISTS "Allow public insert submissions" ON account_submissions;
CREATE POLICY "Allow public insert submissions" ON account_submissions FOR INSERT TO public WITH CHECK (true);

DROP POLICY IF EXISTS "Allow public update submissions" ON account_submissions;
CREATE POLICY "Allow public update submissions" ON account_submissions FOR UPDATE TO public USING (true);

DROP POLICY IF EXISTS "Allow public delete submissions" ON account_submissions;
CREATE POLICY "Allow public delete submissions" ON account_submissions FOR DELETE TO public USING (true);

-- Trigger for updated_at
DROP TRIGGER IF EXISTS set_updated_at_submissions ON account_submissions;
CREATE TRIGGER set_updated_at_submissions
BEFORE UPDATE ON account_submissions
FOR EACH ROW
EXECUTE FUNCTION handle_updated_at();