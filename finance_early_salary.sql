-- Table for Early Salary Requests (Kasbon Gaji)
CREATE TABLE IF NOT EXISTS finance_early_salary_requests (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    account_id UUID NOT NULL CONSTRAINT finance_early_salary_requests_account_id_fkey REFERENCES accounts(id) ON DELETE CASCADE,
    month INTEGER NOT NULL CHECK (month BETWEEN 1 AND 12),
    year INTEGER NOT NULL,
    amount DECIMAL(15, 2) NOT NULL DEFAULT 0,
    reason TEXT NOT NULL,
    status VARCHAR(20) NOT NULL DEFAULT 'Pending' CHECK (status IN ('Pending', 'Approved', 'Paid', 'Rejected')),
    payment_proof_id TEXT, -- Google Drive File ID
    rejection_reason TEXT,
    verifier_id UUID CONSTRAINT finance_early_salary_requests_verifier_id_fkey REFERENCES accounts(id),
    verified_at TIMESTAMP WITH TIME ZONE,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Disable RLS
ALTER TABLE finance_early_salary_requests DISABLE ROW LEVEL SECURITY;

-- Index for faster lookups
CREATE INDEX IF NOT EXISTS idx_early_salary_account ON finance_early_salary_requests(account_id);
CREATE INDEX IF NOT EXISTS idx_early_salary_period ON finance_early_salary_requests(month, year);
