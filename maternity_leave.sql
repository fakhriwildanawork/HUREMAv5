-- 1. Update Tabel Accounts (Tambah Kolom Jatah Cuti Melahirkan dalam Hari)
ALTER TABLE accounts ADD COLUMN IF NOT EXISTS maternity_leave_quota INTEGER DEFAULT 0;

-- 2. Membuat Tabel Cuti Melahirkan
CREATE TABLE account_maternity_leaves (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  account_id UUID REFERENCES accounts(id) ON DELETE CASCADE,
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

-- 3. Menonaktifkan RLS (Karena sistem menggunakan login kustom/localStorage)
ALTER TABLE account_maternity_leaves DISABLE ROW LEVEL SECURITY;

-- 4. Memberikan izin akses (GRANT) eksplisit agar API bisa mengakses tabel
GRANT ALL ON TABLE account_maternity_leaves TO authenticated;
GRANT ALL ON TABLE account_maternity_leaves TO anon;
GRANT ALL ON TABLE account_maternity_leaves TO service_role;

-- 5. Opsional: Menambahkan index untuk performa pencarian
CREATE INDEX idx_maternity_leaves_account_id ON account_maternity_leaves(account_id);
CREATE INDEX idx_maternity_leaves_status ON account_maternity_leaves(status);

-- 6. PAKSA Supabase memuat ulang skema API (Sangat Penting)
NOTIFY pgrst, 'reload schema';
