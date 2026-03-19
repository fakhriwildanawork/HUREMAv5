-- Finance Salary Schemes table
CREATE TABLE IF NOT EXISTS finance_salary_schemes (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  name TEXT NOT NULL,
  description TEXT,
  type TEXT NOT NULL CHECK (type IN ('Harian', 'Bulanan')),
  basic_salary NUMERIC DEFAULT 0,
  position_allowance NUMERIC DEFAULT 0,
  placement_allowance NUMERIC DEFAULT 0,
  other_allowance NUMERIC DEFAULT 0,
  overtime_rate_per_hour NUMERIC DEFAULT 0,
  late_deduction_per_minute NUMERIC DEFAULT 0,
  early_leave_deduction_per_minute NUMERIC DEFAULT 0,
  no_clock_out_deduction_per_day NUMERIC DEFAULT 0,
  absent_deduction_per_day NUMERIC DEFAULT 0,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Salary Assignments table (Mapping scheme to employees)
CREATE TABLE IF NOT EXISTS finance_salary_assignments (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  scheme_id UUID NOT NULL REFERENCES finance_salary_schemes(id) ON DELETE CASCADE,
  account_id UUID NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
  UNIQUE(account_id)
);

-- Disable RLS
ALTER TABLE finance_salary_schemes DISABLE ROW LEVEL SECURITY;
ALTER TABLE finance_salary_assignments DISABLE ROW LEVEL SECURITY;
