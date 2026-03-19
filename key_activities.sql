-- Table for Key Activities Master
CREATE TABLE IF NOT EXISTS key_activities (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    title TEXT NOT NULL,
    description TEXT,
    weight INTEGER DEFAULT 1,
    start_date DATE NOT NULL,
    end_date DATE NOT NULL,
    status TEXT DEFAULT 'Active', -- Active, Pause, Completed
    recurrence_type TEXT NOT NULL, -- Once, Daily, Weekly, Monthly, EndOfMonth
    recurrence_rule JSONB DEFAULT '{}', -- { days_of_week: [1,3], dates_of_month: [5,20] }
    supporting_links JSONB DEFAULT '[]',
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Table for Multi-Assignment
CREATE TABLE IF NOT EXISTS key_activity_assignments (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    activity_id UUID REFERENCES key_activities(id) ON DELETE CASCADE,
    account_id UUID REFERENCES accounts(id) ON DELETE CASCADE,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    UNIQUE(activity_id, account_id)
);

-- Table for Activity Reports (Transactions)
CREATE TABLE IF NOT EXISTS key_activity_reports (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    activity_id UUID REFERENCES key_activities(id) ON DELETE CASCADE,
    account_id UUID REFERENCES accounts(id) ON DELETE CASCADE,
    due_date DATE NOT NULL, -- The date this report is for
    reported_at TIMESTAMPTZ DEFAULT NOW(),
    description TEXT,
    file_ids JSONB DEFAULT '[]',
    links JSONB DEFAULT '[]',
    status TEXT DEFAULT 'Unverified', -- Unverified, Verified
    verification_data JSONB DEFAULT NULL, -- { score: 100, notes: 'Good', verified_at: '...', verifier_id: '...' }
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW(),
    UNIQUE(activity_id, account_id, due_date)
);

-- Index for performance
CREATE INDEX IF NOT EXISTS idx_key_activity_reports_due_date ON key_activity_reports(due_date);
CREATE INDEX IF NOT EXISTS idx_key_activity_reports_account_id ON key_activity_reports(account_id);
CREATE INDEX IF NOT EXISTS idx_key_activity_assignments_account_id ON key_activity_assignments(account_id);

-- Enable RLS
ALTER TABLE key_activities ENABLE ROW LEVEL SECURITY;
ALTER TABLE key_activity_assignments ENABLE ROW LEVEL SECURITY;
ALTER TABLE key_activity_reports ENABLE ROW LEVEL SECURITY;

-- Policies for key_activities
CREATE POLICY "Enable read access for all authenticated users" ON key_activities
    FOR SELECT USING (auth.role() = 'authenticated');

CREATE POLICY "Enable insert for admins only" ON key_activities
    FOR INSERT WITH CHECK (
        EXISTS (SELECT 1 FROM accounts WHERE id = auth.uid() AND (access_code LIKE 'SP%' OR access_code LIKE '%ADM%'))
    );

CREATE POLICY "Enable update for admins only" ON key_activities
    FOR UPDATE USING (
        EXISTS (SELECT 1 FROM accounts WHERE id = auth.uid() AND (access_code LIKE 'SP%' OR access_code LIKE '%ADM%'))
    );

CREATE POLICY "Enable delete for admins only" ON key_activities
    FOR DELETE USING (
        EXISTS (SELECT 1 FROM accounts WHERE id = auth.uid() AND (access_code LIKE 'SP%' OR access_code LIKE '%ADM%'))
    );

-- Policies for key_activity_assignments
CREATE POLICY "Enable read for authenticated users" ON key_activity_assignments
    FOR SELECT USING (auth.role() = 'authenticated');

CREATE POLICY "Enable all for admins" ON key_activity_assignments
    FOR ALL USING (
        EXISTS (SELECT 1 FROM accounts WHERE id = auth.uid() AND (access_code LIKE 'SP%' OR access_code LIKE '%ADM%'))
    );

-- Policies for key_activity_reports
CREATE POLICY "Enable read for owner and admins" ON key_activity_reports
    FOR SELECT USING (
        account_id = auth.uid() OR 
        EXISTS (SELECT 1 FROM accounts WHERE id = auth.uid() AND (access_code LIKE 'SP%' OR access_code LIKE '%ADM%'))
    );

CREATE POLICY "Enable insert for owner" ON key_activity_reports
    FOR INSERT WITH CHECK (account_id = auth.uid());

CREATE POLICY "Enable update for owner and admins" ON key_activity_reports
    FOR UPDATE USING (
        account_id = auth.uid() OR 
        EXISTS (SELECT 1 FROM accounts WHERE id = auth.uid() AND (access_code LIKE 'SP%' OR access_code LIKE '%ADM%'))
    );
