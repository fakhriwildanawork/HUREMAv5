-- Table for Payroll Header
CREATE TABLE IF NOT EXISTS finance_payrolls (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    month INTEGER NOT NULL CHECK (month BETWEEN 1 AND 12),
    year INTEGER NOT NULL,
    start_date DATE NOT NULL,
    end_date DATE NOT NULL,
    status VARCHAR(20) NOT NULL DEFAULT 'Draft' CHECK (status IN ('Draft', 'Pending', 'Approved', 'Paid')),
    verifier_id UUID REFERENCES accounts(id),
    verified_at TIMESTAMP WITH TIME ZONE,
    verification_notes TEXT,
    created_by UUID REFERENCES accounts(id),
    updated_by UUID REFERENCES accounts(id),
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(month, year)
);

-- Table for Payroll Items (Individual Employee Payslips)
CREATE TABLE IF NOT EXISTS finance_payroll_items (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    payroll_id UUID NOT NULL REFERENCES finance_payrolls(id) ON DELETE CASCADE,
    account_id UUID NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
    
    -- Salary Type
    salary_type VARCHAR(20) NOT NULL, -- 'Harian' | 'Bulanan'
    
    -- Income Components
    basic_salary DECIMAL(15, 2) NOT NULL DEFAULT 0,
    basic_salary_notes TEXT,
    
    position_allowance DECIMAL(15, 2) NOT NULL DEFAULT 0,
    position_allowance_notes TEXT,
    
    placement_allowance DECIMAL(15, 2) NOT NULL DEFAULT 0,
    placement_allowance_notes TEXT,
    
    other_allowance DECIMAL(15, 2) NOT NULL DEFAULT 0,
    other_allowance_notes TEXT,
    
    overtime_pay DECIMAL(15, 2) NOT NULL DEFAULT 0,
    overtime_pay_notes TEXT,
    
    other_additions DECIMAL(15, 2) NOT NULL DEFAULT 0,
    other_additions_notes TEXT,
    
    -- Deduction Components
    late_deduction DECIMAL(15, 2) NOT NULL DEFAULT 0,
    late_deduction_notes TEXT,
    
    early_leave_deduction DECIMAL(15, 2) NOT NULL DEFAULT 0,
    early_leave_deduction_notes TEXT,
    
    absent_deduction DECIMAL(15, 2) NOT NULL DEFAULT 0,
    absent_deduction_notes TEXT,
    
    other_deductions DECIMAL(15, 2) NOT NULL DEFAULT 0,
    other_deductions_notes TEXT,
    
    bpjs_kesehatan DECIMAL(15, 2) NOT NULL DEFAULT 0,
    bpjs_ketenagakerjaan DECIMAL(15, 2) NOT NULL DEFAULT 0,
    pph21 DECIMAL(15, 2) NOT NULL DEFAULT 0,
    
    -- Totals
    total_income DECIMAL(15, 2) NOT NULL DEFAULT 0,
    total_deduction DECIMAL(15, 2) NOT NULL DEFAULT 0,
    take_home_pay DECIMAL(15, 2) NOT NULL DEFAULT 0,
    
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(payroll_id, account_id)
);

-- Table for Payroll Settings (Payslip Header)
CREATE TABLE IF NOT EXISTS finance_payroll_settings (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    company_name VARCHAR(255) NOT NULL,
    company_address TEXT,
    company_phone VARCHAR(50),
    company_email VARCHAR(100),
    company_website VARCHAR(100),
    company_logo_url TEXT,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Initial settings record
INSERT INTO finance_payroll_settings (company_name) 
SELECT 'HUREMA HRIS' WHERE NOT EXISTS (SELECT 1 FROM finance_payroll_settings);
