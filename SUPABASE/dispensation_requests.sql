-- Create dispensation_requests table
CREATE TABLE IF NOT EXISTS dispensation_requests (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    account_id UUID NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
    presence_id UUID REFERENCES attendance(id) ON DELETE SET NULL,
    date DATE NOT NULL,
    issues JSONB NOT NULL DEFAULT '[]',
    reason TEXT NOT NULL,
    file_id TEXT,
    is_read BOOLEAN NOT NULL DEFAULT FALSE,
    status TEXT NOT NULL DEFAULT 'PENDING' CHECK (status IN ('PENDING', 'APPROVED', 'REJECTED', 'PARTIAL')),
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Add RLS
ALTER TABLE dispensation_requests ENABLE ROW LEVEL SECURITY;

-- Policy for users to view their own requests
CREATE POLICY "Users can view their own dispensation requests" 
ON dispensation_requests FOR SELECT 
USING (auth.uid() = account_id);

-- Policy for users to insert their own requests
CREATE POLICY "Users can insert their own dispensation requests" 
ON dispensation_requests FOR INSERT 
WITH CHECK (auth.uid() = account_id);

-- Policy for users to update their own pending requests
CREATE POLICY "Users can update their own pending dispensation requests" 
ON dispensation_requests FOR UPDATE 
USING (auth.uid() = account_id AND status = 'PENDING')
WITH CHECK (auth.uid() = account_id AND status = 'PENDING');

-- Policy for users to delete their own pending requests
CREATE POLICY "Users can delete their own pending dispensation requests" 
ON dispensation_requests FOR DELETE 
USING (auth.uid() = account_id AND status = 'PENDING');

-- Policy for admins to view all requests
CREATE POLICY "Admins can view all dispensation requests" 
ON dispensation_requests FOR SELECT 
USING (EXISTS (SELECT 1 FROM accounts WHERE id = auth.uid() AND role = 'admin'));

-- Policy for admins to update all requests
CREATE POLICY "Admins can update all dispensation requests" 
ON dispensation_requests FOR UPDATE 
USING (EXISTS (SELECT 1 FROM accounts WHERE id = auth.uid() AND role = 'admin'));
