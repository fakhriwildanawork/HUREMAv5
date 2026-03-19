-- feedback.sql
-- Table for employee feedback

CREATE TABLE IF NOT EXISTS account_feedbacks (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  account_id UUID REFERENCES accounts(id) ON DELETE CASCADE,
  category TEXT NOT NULL, -- Gaji, Fasilitas, Hubungan Kerja, Lainnya
  priority TEXT NOT NULL DEFAULT 'Medium', -- Low, Medium, High
  is_anonymous BOOLEAN DEFAULT false,
  description TEXT NOT NULL,
  attachments JSONB DEFAULT '[]', -- Array of file IDs
  links JSONB DEFAULT '[]', -- Array of URLs
  status TEXT NOT NULL DEFAULT 'Unread', -- Unread, Read
  created_at TIMESTAMPTZ DEFAULT now(),
  updated_at TIMESTAMPTZ DEFAULT now()
);

-- Disable RLS for now to match the app's architecture
ALTER TABLE account_feedbacks DISABLE ROW LEVEL SECURITY;

-- Grant access
GRANT ALL ON TABLE account_feedbacks TO authenticated;
GRANT ALL ON TABLE account_feedbacks TO anon;
GRANT ALL ON TABLE account_feedbacks TO service_role;

-- Index for performance
CREATE INDEX IF NOT EXISTS idx_feedbacks_account_id ON account_feedbacks(account_id);
CREATE INDEX IF NOT EXISTS idx_feedbacks_status ON account_feedbacks(status);
CREATE INDEX IF NOT EXISTS idx_feedbacks_created_at ON account_feedbacks(created_at);
