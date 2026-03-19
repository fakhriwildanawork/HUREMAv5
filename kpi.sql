-- kpi.sql
-- Tabel untuk menyimpan Key Performance Indicator (KPI) pegawai

CREATE TABLE account_kpis (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  account_id UUID REFERENCES accounts(id) ON DELETE CASCADE,
  title TEXT NOT NULL,
  description TEXT NOT NULL,
  weight INTEGER DEFAULT 1, -- Bobot KPI (misal 1-5)
  start_date DATE NOT NULL,
  deadline DATE NOT NULL,
  status TEXT NOT NULL DEFAULT 'Active', -- Active, Pause, Unverified, Verified, Unreported
  
  -- Data Pelaporan (JSONB)
  -- Struktur: { description: string, file_ids: string[], links: string[], self_assessment: number, reported_at: string }
  report_data JSONB DEFAULT NULL,
  
  -- Data Verifikasi (JSONB)
  -- Struktur: { score: number, notes: string, verified_at: string, verifier_id: string }
  verification_data JSONB DEFAULT NULL,
  
  created_at TIMESTAMPTZ DEFAULT now(),
  updated_at TIMESTAMPTZ DEFAULT now()
);

-- Menonaktifkan RLS agar API bisa mengakses tabel (sesuai arsitektur aplikasi)
ALTER TABLE account_kpis DISABLE ROW LEVEL SECURITY;

-- Memberikan izin akses
GRANT ALL ON TABLE account_kpis TO authenticated;
GRANT ALL ON TABLE account_kpis TO anon;
GRANT ALL ON TABLE account_kpis TO service_role;

-- Index untuk performa
CREATE INDEX idx_kpis_account_id ON account_kpis(account_id);
CREATE INDEX idx_kpis_status ON account_kpis(status);
CREATE INDEX idx_kpis_deadline ON account_kpis(deadline);

-- Paksa reload skema
NOTIFY pgrst, 'reload schema';
