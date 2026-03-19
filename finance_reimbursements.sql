-- Finance Reimbursements table
CREATE TABLE IF NOT EXISTS finance_reimbursements (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  account_id UUID NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
  transaction_date DATE NOT NULL,
  category TEXT NOT NULL, -- Operasional, Akomodasi, Inventaris, dll
  description TEXT NOT NULL,
  amount_requested NUMERIC NOT NULL,
  amount_approved NUMERIC,
  proof_file_id TEXT, -- File ID from Google Drive (Bukti Transaksi)
  payment_method TEXT NOT NULL CHECK (payment_method IN ('Cash', 'Transfer')),
  target_type TEXT, -- Bank, E-Wallet
  target_name TEXT, -- Nama Bank/E-Wallet
  account_number TEXT,
  account_holder TEXT,
  status TEXT NOT NULL DEFAULT 'Pending' CHECK (status IN ('Pending', 'Approved', 'Partially Approved', 'Rejected')),
  is_read BOOLEAN DEFAULT FALSE,
  admin_notes TEXT,
  payment_proof_id TEXT, -- File ID from Google Drive (Bukti Transfer Admin)
  verifier_id UUID REFERENCES accounts(id),
  verified_at TIMESTAMP WITH TIME ZONE,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Disable RLS
ALTER TABLE finance_reimbursements DISABLE ROW LEVEL SECURITY;
