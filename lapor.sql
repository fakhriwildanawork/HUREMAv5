-- Create account_reports table for Whistleblowing System
CREATE TABLE IF NOT EXISTS account_reports (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    account_id UUID NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
    category VARCHAR(255) NOT NULL,
    description TEXT NOT NULL,
    reported_account_ids JSONB DEFAULT '[]'::jsonb, -- Array of account IDs
    attachments JSONB DEFAULT '[]'::jsonb, -- Array of file IDs
    links JSONB DEFAULT '[]'::jsonb, -- Array of URLs
    status VARCHAR(50) DEFAULT 'Unread',
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Disable RLS for now as per project pattern
ALTER TABLE account_reports DISABLE ROW LEVEL SECURITY;

-- Add index for performance
CREATE INDEX IF NOT EXISTS idx_account_reports_account_id ON account_reports(account_id);
CREATE INDEX IF NOT EXISTS idx_account_reports_status ON account_reports(status);
