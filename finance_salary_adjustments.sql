-- Table for Salary Adjustments (Additions/Deductions)
CREATE TABLE IF NOT EXISTS finance_salary_adjustments (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    account_id UUID NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
    type VARCHAR(20) NOT NULL CHECK (type IN ('Addition', 'Deduction')),
    amount DECIMAL(15, 2) NOT NULL DEFAULT 0,
    month INTEGER NOT NULL CHECK (month BETWEEN 1 AND 12),
    year INTEGER NOT NULL,
    description TEXT NOT NULL,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Table for Payroll Status (Locking Mechanism)
CREATE TABLE IF NOT EXISTS finance_payroll_status (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    month INTEGER NOT NULL CHECK (month BETWEEN 1 AND 12),
    year INTEGER NOT NULL,
    status VARCHAR(20) NOT NULL DEFAULT 'Draft' CHECK (status IN ('Draft', 'Approved', 'Paid')),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(month, year)
);

-- Index for performance
CREATE INDEX IF NOT EXISTS idx_salary_adjustments_period ON finance_salary_adjustments(month, year);
CREATE INDEX IF NOT EXISTS idx_salary_adjustments_account ON finance_salary_adjustments(account_id);
