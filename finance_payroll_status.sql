
-- Table for Payroll Status tracking
CREATE TABLE IF NOT EXISTS finance_payroll_status (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    month INTEGER NOT NULL CHECK (month BETWEEN 1 AND 12),
    year INTEGER NOT NULL,
    status VARCHAR(20) NOT NULL DEFAULT 'Draft' CHECK (status IN ('Draft', 'Approved', 'Paid')),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(month, year)
);

-- Disable RLS
ALTER TABLE finance_payroll_status DISABLE ROW LEVEL SECURITY;

-- Grant access
GRANT ALL ON TABLE public.finance_payroll_status TO service_role;
GRANT ALL ON TABLE public.finance_payroll_status TO anon;
GRANT ALL ON TABLE public.finance_payroll_status TO authenticated;
