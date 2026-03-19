
-- Create Overtimes Table
CREATE TABLE IF NOT EXISTS overtimes (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    account_id UUID REFERENCES accounts(id) ON DELETE CASCADE,
    
    -- Check In Data
    check_in TIMESTAMP WITH TIME ZONE,
    in_latitude NUMERIC,
    in_longitude NUMERIC,
    in_photo_id TEXT,
    in_address TEXT,

    -- Check Out Data
    check_out TIMESTAMP WITH TIME ZONE,
    out_latitude NUMERIC,
    out_longitude NUMERIC,
    out_photo_id TEXT,
    out_address TEXT,

    -- Summary
    duration_minutes INTEGER DEFAULT 0,
    work_duration TEXT, -- Format HH:mm:ss
    
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Enable RLS
ALTER TABLE overtimes ENABLE ROW LEVEL SECURITY;

-- Policies
DROP POLICY IF EXISTS "Allow public read overtimes" ON overtimes;
CREATE POLICY "Allow public read overtimes" ON overtimes FOR SELECT TO public USING (true);

DROP POLICY IF EXISTS "Allow public insert overtimes" ON overtimes;
CREATE POLICY "Allow public insert overtimes" ON overtimes FOR INSERT TO public WITH CHECK (true);

DROP POLICY IF EXISTS "Allow public update overtimes" ON overtimes;
CREATE POLICY "Allow public update overtimes" ON overtimes FOR UPDATE TO public USING (true);
