-- Enable UUID extension if not enabled
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- Table for Warning Logs (SP)
CREATE TABLE IF NOT EXISTS account_warning_logs (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    account_id UUID REFERENCES accounts(id) ON DELETE CASCADE,
    warning_type TEXT NOT NULL CHECK (warning_type IN ('Teguran', 'SP1', 'SP2', 'SP3')),
    reason TEXT NOT NULL,
    issue_date DATE NOT NULL,
    file_id TEXT, -- Google Drive File ID
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Table for Termination Logs (Exit)
CREATE TABLE IF NOT EXISTS account_termination_logs (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    account_id UUID REFERENCES accounts(id) ON DELETE CASCADE,
    termination_type TEXT NOT NULL CHECK (termination_type IN ('Pemecatan', 'Resign')),
    termination_date DATE NOT NULL,
    reason TEXT NOT NULL,
    severance_amount NUMERIC DEFAULT 0,
    penalty_amount NUMERIC DEFAULT 0,
    file_id TEXT, -- Google Drive File ID
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Enable RLS
ALTER TABLE account_warning_logs ENABLE ROW LEVEL SECURITY;
ALTER TABLE account_termination_logs ENABLE ROW LEVEL SECURITY;

-- Public Policies
DROP POLICY IF EXISTS "Allow public read warning logs" ON account_warning_logs;
CREATE POLICY "Allow public read warning logs" ON account_warning_logs FOR SELECT TO public USING (true);
DROP POLICY IF EXISTS "Allow public insert warning logs" ON account_warning_logs;
CREATE POLICY "Allow public insert warning logs" ON account_warning_logs FOR INSERT TO public WITH CHECK (true);
DROP POLICY IF EXISTS "Allow public delete warning logs" ON account_warning_logs;
CREATE POLICY "Allow public delete warning logs" ON account_warning_logs FOR DELETE TO public USING (true);

DROP POLICY IF EXISTS "Allow public read termination logs" ON account_termination_logs;
CREATE POLICY "Allow public read termination logs" ON account_termination_logs FOR SELECT TO public USING (true);
DROP POLICY IF EXISTS "Allow public insert termination logs" ON account_termination_logs;
CREATE POLICY "Allow public insert termination logs" ON account_termination_logs FOR INSERT TO public WITH CHECK (true);
DROP POLICY IF EXISTS "Allow public delete termination logs" ON account_termination_logs;
CREATE POLICY "Allow public delete termination logs" ON account_termination_logs FOR DELETE TO public USING (true);