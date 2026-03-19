-- Create meetings table for Meeting Minutes System
CREATE TABLE IF NOT EXISTS meetings (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    title VARCHAR(255) NOT NULL,
    description TEXT,
    scheduled_at TIMESTAMP WITH TIME ZONE NOT NULL,
    started_at TIMESTAMP WITH TIME ZONE,
    ended_at TIMESTAMP WITH TIME ZONE,
    location_type VARCHAR(50) NOT NULL, -- 'Online' or 'Offline'
    location_detail TEXT NOT NULL,
    latitude DOUBLE PRECISION,
    longitude DOUBLE PRECISION,
    participant_ids JSONB DEFAULT '[]'::jsonb, -- Array of account IDs
    notulen_ids JSONB DEFAULT '[]'::jsonb, -- Array of account IDs
    minutes_content TEXT,
    attachments JSONB DEFAULT '[]'::jsonb, -- Array of file IDs
    links JSONB DEFAULT '[]'::jsonb, -- Array of URLs
    status VARCHAR(50) DEFAULT 'Scheduled', -- 'Scheduled', 'In Progress', 'Completed', 'Cancelled'
    created_by UUID NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Disable RLS for now as per project pattern
ALTER TABLE meetings DISABLE ROW LEVEL SECURITY;

-- Add index for performance
CREATE INDEX IF NOT EXISTS idx_meetings_created_by ON meetings(created_by);
CREATE INDEX IF NOT EXISTS idx_meetings_status ON meetings(status);
CREATE INDEX IF NOT EXISTS idx_meetings_scheduled_at ON meetings(scheduled_at);
