-- 1. Membuat Tabel Izin
CREATE TABLE account_permission_requests (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  account_id UUID REFERENCES accounts(id) ON DELETE CASCADE,
  permission_type TEXT NOT NULL, -- Sakit, Urusan Keluarga, Izin Menikah, dsb
  start_date DATE NOT NULL,
  end_date DATE NOT NULL,
  description TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'pending', -- Status: pending, approved, rejected, negotiating, cancelled
  file_id TEXT, -- ID file dari Google Drive (dokumen pendukung)
  negotiation_data JSONB DEFAULT '[]'::jsonb, -- Menyimpan riwayat negosiasi bolak-balik
  current_negotiator_role TEXT DEFAULT 'admin', -- Menentukan giliran siapa yang merespons (admin/user)
  created_at TIMESTAMPTZ DEFAULT now(),
  updated_at TIMESTAMPTZ DEFAULT now()
);

-- 2. Menonaktifkan RLS (Karena sistem menggunakan login kustom/localStorage)
ALTER TABLE account_permission_requests DISABLE ROW LEVEL SECURITY;

-- 3. Memberikan izin akses (GRANT) eksplisit agar API bisa mengakses tabel
GRANT ALL ON TABLE account_permission_requests TO authenticated;
GRANT ALL ON TABLE account_permission_requests TO anon;
GRANT ALL ON TABLE account_permission_requests TO service_role;

-- 4. Opsional: Menambahkan index untuk performa pencarian
CREATE INDEX idx_permission_requests_account_id ON account_permission_requests(account_id);
CREATE INDEX idx_permission_requests_status ON account_permission_requests(status);

-- 5. PAKSA Supabase memuat ulang skema API (Sangat Penting)
NOTIFY pgrst, 'reload schema';
