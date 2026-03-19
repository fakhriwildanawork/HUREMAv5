-- update_only.sql
-- Menambahkan kolom untuk akumulasi cuti tahunan pada tabel accounts

ALTER TABLE accounts 
ADD COLUMN IF NOT EXISTS is_leave_accumulated BOOLEAN DEFAULT FALSE,
ADD COLUMN IF NOT EXISTS max_carry_over_days INTEGER DEFAULT 0,
ADD COLUMN IF NOT EXISTS carry_over_quota INTEGER DEFAULT 0;

COMMENT ON COLUMN accounts.is_leave_accumulated IS 'Apakah jatah cuti tahun lalu boleh diakumulasikan ke tahun ini';
COMMENT ON COLUMN accounts.max_carry_over_days IS 'Maksimal jatah cuti tahun lalu yang boleh dibawa ke tahun ini';
COMMENT ON COLUMN accounts.carry_over_quota IS 'Jumlah jatah cuti tahun lalu yang berhasil dibawa ke tahun ini';

-- Menambahkan kolom link pendukung pada tabel account_kpis
ALTER TABLE account_kpis 
ADD COLUMN IF NOT EXISTS supporting_links JSONB DEFAULT '[]'::jsonb;

COMMENT ON COLUMN account_kpis.supporting_links IS 'Daftar link referensi/pendukung dari Admin saat pembuatan KPI';
