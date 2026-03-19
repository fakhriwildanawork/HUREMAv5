
-- Tables for Schedule Management Module

-- 1. Master Schedules Table
CREATE TABLE IF NOT EXISTS schedules (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    name TEXT NOT NULL, -- "Data Nama Jadwal"
    type INTEGER NOT NULL CHECK (type IN (1, 2, 3, 4)), 
    -- 1: Hari Kerja (Fixed), 2: Shift (Uniform), 3: Libur Khusus, 4: Hari Kerja Khusus
    tolerance_minutes INTEGER DEFAULT 0, -- Ini untuk Toleransi Pulang
    tolerance_checkin_minutes INTEGER DEFAULT 0, -- Tambahan untuk Toleransi Datang
    start_date DATE, -- For type 3 & 4
    end_date DATE,   -- For type 3 & 4
    excluded_account_ids UUID[] DEFAULT '{}', -- Selective exclusion for Type 3 & 4
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    -- Proteksi Lapis Terakhir di Database
    CONSTRAINT check_periodic_dates CHECK (
        (type IN (3, 4) AND start_date IS NOT NULL AND end_date IS NOT NULL) OR 
        (type IN (1, 2))
    )
);

-- 2. Schedule Rules (Time details)
CREATE TABLE IF NOT EXISTS schedule_rules (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    schedule_id UUID REFERENCES schedules(id) ON DELETE CASCADE,
    day_of_week INTEGER CHECK (day_of_week BETWEEN 0 AND 6), -- 0: Sunday, 6: Saturday
    check_in_time TIME,
    check_out_time TIME,
    is_holiday BOOLEAN DEFAULT FALSE,
    UNIQUE(schedule_id, day_of_week)
);

-- 3. Junction table for Locations
CREATE TABLE IF NOT EXISTS schedule_locations (
    schedule_id UUID REFERENCES schedules(id) ON DELETE CASCADE,
    location_id UUID REFERENCES locations(id) ON DELETE CASCADE,
    PRIMARY KEY (schedule_id, location_id)
);

-- Enable RLS
ALTER TABLE schedules ENABLE ROW LEVEL SECURITY;
ALTER TABLE schedule_rules ENABLE ROW LEVEL SECURITY;
ALTER TABLE schedule_locations ENABLE ROW LEVEL SECURITY;

-- Public Policies
CREATE POLICY "Allow public read schedules" ON schedules FOR SELECT TO public USING (true);
CREATE POLICY "Allow public insert schedules" ON schedules FOR INSERT TO public WITH CHECK (true);
CREATE POLICY "Allow public update schedules" ON schedules FOR UPDATE TO public USING (true);
CREATE POLICY "Allow public delete schedules" ON schedules FOR DELETE TO public USING (true);

CREATE POLICY "Allow public read rules" ON schedule_rules FOR SELECT TO public USING (true);
CREATE POLICY "Allow public insert rules" ON schedule_rules FOR INSERT TO public WITH CHECK (true);
CREATE POLICY "Allow public update rules" ON schedule_rules FOR UPDATE TO public USING (true);
CREATE POLICY "Allow public delete rules" ON schedule_rules FOR DELETE TO public USING (true);

CREATE POLICY "Allow public read schedule_locations" ON schedule_locations FOR SELECT TO public USING (true);
CREATE POLICY "Allow public insert schedule_locations" ON schedule_locations FOR INSERT TO public WITH CHECK (true);
CREATE POLICY "Allow public delete schedule_locations" ON schedule_locations FOR DELETE TO public USING (true);

-- Update Trigger for updated_at
DROP TRIGGER IF EXISTS set_updated_at_schedules ON schedules;
CREATE TRIGGER set_updated_at_schedules
BEFORE UPDATE ON schedules
FOR EACH ROW
EXECUTE FUNCTION handle_updated_at();
