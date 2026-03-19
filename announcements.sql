-- Announcements table
CREATE TABLE IF NOT EXISTS announcements (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  title TEXT NOT NULL,
  content TEXT NOT NULL,
  category TEXT NOT NULL CHECK (category IN ('Urgent', 'Info', 'Event', 'Policy')),
  target_type TEXT NOT NULL CHECK (target_type IN ('All', 'Department', 'Individual')),
  target_ids TEXT[] DEFAULT '{}',
  publish_start TIMESTAMP WITH TIME ZONE NOT NULL,
  publish_end TIMESTAMP WITH TIME ZONE NOT NULL,
  attachments TEXT[] DEFAULT '{}',
  created_by UUID NOT NULL REFERENCES accounts(id),
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Announcement reads tracking
CREATE TABLE IF NOT EXISTS announcement_reads (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  announcement_id UUID NOT NULL REFERENCES announcements(id) ON DELETE CASCADE,
  user_id UUID NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
  read_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
  UNIQUE(announcement_id, user_id)
);

-- Disable RLS as per project pattern
ALTER TABLE announcements DISABLE ROW LEVEL SECURITY;
ALTER TABLE announcement_reads DISABLE ROW LEVEL SECURITY;
