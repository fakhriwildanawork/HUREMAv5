
-- Enable UUID extension if not enabled
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- Create Locations Table
CREATE TABLE IF NOT EXISTS locations (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    name TEXT NOT NULL,
    location_type TEXT,
    address TEXT NOT NULL,
    city TEXT NOT NULL,
    province TEXT NOT NULL,
    zip_code TEXT,
    phone TEXT,
    latitude NUMERIC,
    longitude NUMERIC,
    radius INTEGER DEFAULT 100,
    description TEXT,
    search_all TEXT,
    image_google_id TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Create Location Administrations Table
CREATE TABLE IF NOT EXISTS location_administrations (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    location_id UUID REFERENCES locations(id) ON DELETE CASCADE,
    admin_date DATE NOT NULL,
    status TEXT NOT NULL, -- 'Milik Sendiri', 'Sewa/Kontrak', 'Kerjasama'
    due_date DATE,
    description TEXT,
    file_ids TEXT[], -- Array of Google Drive IDs
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Function to handle search_all indexing
CREATE OR REPLACE FUNCTION update_location_search_all()
RETURNS TRIGGER AS $$
BEGIN
    NEW.search_all := 
        COALESCE(NEW.name, '') || ' ' || 
        COALESCE(NEW.location_type, '') || ' ' || 
        COALESCE(NEW.address, '') || ' ' || 
        COALESCE(NEW.city, '') || ' ' || 
        COALESCE(NEW.province, '') || ' ' || 
        COALESCE(NEW.phone, '');
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

-- Trigger for search_all
DROP TRIGGER IF EXISTS trg_update_location_search_all ON locations;
CREATE TRIGGER trg_update_location_search_all
BEFORE INSERT OR UPDATE ON locations
FOR EACH ROW
EXECUTE FUNCTION update_location_search_all();

-- Enable Row Level Security
ALTER TABLE locations ENABLE ROW LEVEL SECURITY;
ALTER TABLE location_administrations ENABLE ROW LEVEL SECURITY;

-- Create Public Policies for locations
DROP POLICY IF EXISTS "Allow public read access" ON locations;
CREATE POLICY "Allow public read access" ON locations FOR SELECT TO public USING (true);
DROP POLICY IF EXISTS "Allow public insert" ON locations;
CREATE POLICY "Allow public insert" ON locations FOR INSERT TO public WITH CHECK (true);
DROP POLICY IF EXISTS "Allow public update" ON locations;
CREATE POLICY "Allow public update" ON locations FOR UPDATE TO public USING (true);
DROP POLICY IF EXISTS "Allow public delete" ON locations;
CREATE POLICY "Allow public delete" ON locations FOR DELETE TO public USING (true);

-- Create Public Policies for administrations
DROP POLICY IF EXISTS "Allow public read admin" ON location_administrations;
CREATE POLICY "Allow public read admin" ON location_administrations FOR SELECT TO public USING (true);
DROP POLICY IF EXISTS "Allow public insert admin" ON location_administrations;
CREATE POLICY "Allow public insert admin" ON location_administrations FOR INSERT TO public WITH CHECK (true);
DROP POLICY IF EXISTS "Allow public delete admin" ON location_administrations;
CREATE POLICY "Allow public delete admin" ON location_administrations FOR DELETE TO public USING (true);

-- Function to handle updated_at
CREATE OR REPLACE FUNCTION handle_updated_at()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = NOW();
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

-- Trigger for updated_at
DROP TRIGGER IF EXISTS set_updated_at ON locations;
CREATE TRIGGER set_updated_at
BEFORE UPDATE ON locations
FOR EACH ROW
EXECUTE FUNCTION handle_updated_at();
