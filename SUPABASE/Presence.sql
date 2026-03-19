
-- Enable UUID extension
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- Create Attendances Table
CREATE TABLE IF NOT EXISTS attendances (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    account_id UUID REFERENCES accounts(id) ON DELETE CASCADE,
    
    -- Check In Data
    check_in TIMESTAMP WITH TIME ZONE,
    in_latitude NUMERIC,
    in_longitude NUMERIC,
    in_photo_id TEXT, -- ID Photo G-Drive dari proses Liveness
    in_address TEXT,
    status_in TEXT DEFAULT 'Tepat Waktu', -- Tepat Waktu, Terlambat
    late_minutes INTEGER DEFAULT 0,
    late_reason TEXT,

    -- Check Out Data
    check_out TIMESTAMP WITH TIME ZONE,
    out_latitude NUMERIC,
    out_longitude NUMERIC,
    out_photo_id TEXT, -- ID Photo G-Drive dari proses Liveness
    out_address TEXT,
    status_out TEXT DEFAULT 'Tepat Waktu', -- Tepat Waktu, Pulang Cepat
    early_departure_minutes INTEGER DEFAULT 0,
    early_departure_reason TEXT,

    -- Summary
    work_duration TEXT, -- Kalkulasi durasi (HH:mm:ss)
    
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Enable RLS
ALTER TABLE attendances ENABLE ROW LEVEL SECURITY;

-- Policies
DROP POLICY IF EXISTS "Allow public read attendance" ON attendances;
CREATE POLICY "Allow public read attendance" ON attendances FOR SELECT TO public USING (true);

DROP POLICY IF EXISTS "Allow public insert attendance" ON attendances;
CREATE POLICY "Allow public insert attendance" ON attendances FOR INSERT TO public WITH CHECK (true);

DROP POLICY IF EXISTS "Allow public update attendance" ON attendances;
CREATE POLICY "Allow public update attendance" ON attendances FOR UPDATE TO public USING (true);

-- Function to get current satellite/server time for the app
CREATE OR REPLACE FUNCTION get_server_time()
RETURNS TIMESTAMP WITH TIME ZONE AS $$
BEGIN
    RETURN NOW();
END;
$$ LANGUAGE plpgsql;
