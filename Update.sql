-- ==========================================
-- PERBAIKAN DATABASE HUREMA (IDEMPOTENT)
-- ==========================================

-- 1. MEMBERSIHKAN DATA YATIM (Fix Error 23503)
-- Memastikan tidak ada ID yang merujuk ke data yang sudah tidak ada sebelum memasang constraint
UPDATE accounts SET schedule_id = NULL WHERE schedule_id IS NOT NULL AND schedule_id NOT IN (SELECT id FROM schedules);
UPDATE accounts SET location_id = NULL WHERE location_id IS NOT NULL AND location_id NOT IN (SELECT id FROM locations);

-- 2. PERBAIKAN RELASI ACCOUNTS (Fix Error PGRST200)
ALTER TABLE accounts 
DROP CONSTRAINT IF EXISTS accounts_schedule_id_fkey,
ADD CONSTRAINT accounts_schedule_id_fkey 
FOREIGN KEY (schedule_id) REFERENCES schedules(id) ON DELETE SET NULL;

ALTER TABLE accounts 
DROP CONSTRAINT IF EXISTS accounts_location_id_fkey,
ADD CONSTRAINT accounts_location_id_fkey 
FOREIGN KEY (location_id) REFERENCES locations(id) ON DELETE SET NULL;

-- 3. PERBAIKAN TRIGGER ACCOUNTS (Fix Error 42710)
DROP TRIGGER IF EXISTS trg_update_account_search_all ON accounts;
CREATE TRIGGER trg_update_account_search_all
BEFORE INSERT OR UPDATE ON accounts
FOR EACH ROW
EXECUTE FUNCTION update_account_search_all();

-- 4. PERBAIKAN POLICIES ACCOUNTS (Fix Error 42710)
DO $$ 
BEGIN
    DROP POLICY IF EXISTS "Allow public read accounts" ON accounts;
    DROP POLICY IF EXISTS "Allow public insert accounts" ON accounts;
    DROP POLICY IF EXISTS "Allow public update accounts" ON accounts;
    DROP POLICY IF EXISTS "Allow public delete accounts" ON accounts;
END $$;

CREATE POLICY "Allow public read accounts" ON accounts FOR SELECT TO public USING (true);
CREATE POLICY "Allow public insert accounts" ON accounts FOR INSERT TO public WITH CHECK (true);
CREATE POLICY "Allow public update accounts" ON accounts FOR UPDATE TO public USING (true);
CREATE POLICY "Allow public delete accounts" ON accounts FOR DELETE TO public USING (true);

-- 5. PENAMBAHAN TABEL SERTIFIKASI (Jika belum ada)
CREATE TABLE IF NOT EXISTS account_certifications (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    account_id UUID REFERENCES accounts(id) ON DELETE CASCADE,
    entry_date DATE DEFAULT NOW(),
    cert_type TEXT,
    cert_name TEXT,
    cert_date DATE,
    file_id TEXT,
    notes TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

ALTER TABLE account_certifications ENABLE ROW LEVEL SECURITY;

DO $$ 
BEGIN
    DROP POLICY IF EXISTS "Allow public read certs" ON account_certifications;
    DROP POLICY IF EXISTS "Allow public insert certs" ON account_certifications;
    DROP POLICY IF EXISTS "Allow public update certs" ON account_certifications;
    DROP POLICY IF EXISTS "Allow public delete certs" ON account_certifications;
END $$;

CREATE POLICY "Allow public read certs" ON account_certifications FOR SELECT TO public USING (true);
CREATE POLICY "Allow public insert certs" ON account_certifications FOR INSERT TO public WITH CHECK (true);
CREATE POLICY "Allow public update certs" ON account_certifications FOR UPDATE TO public USING (true);
CREATE POLICY "Allow public delete certs" ON account_certifications FOR DELETE TO public USING (true);

-- Trigger untuk updated_at pada account_certifications
DROP TRIGGER IF EXISTS set_updated_at_certs ON account_certifications;
CREATE TRIGGER set_updated_at_certs
BEFORE UPDATE ON account_certifications
FOR EACH ROW
EXECUTE FUNCTION handle_updated_at();

-- 6. PENAMBAHAN KOLOM ALASAN PADA TABEL OVERTIMES
ALTER TABLE overtimes ADD COLUMN IF NOT EXISTS reason TEXT;