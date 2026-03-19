
-- Table for Compensation Logs (Severance/Penalty)
CREATE TABLE IF NOT EXISTS public.account_compensation_logs (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    account_id UUID REFERENCES public.accounts(id) ON DELETE CASCADE,
    termination_type TEXT CHECK (termination_type IN ('Resign', 'Pemecatan')),
    termination_date DATE NOT NULL,
    amount NUMERIC NOT NULL DEFAULT 0,
    type TEXT CHECK (type IN ('Severance', 'Penalty')),
    reason TEXT,
    status TEXT DEFAULT 'Pending' CHECK (status IN ('Pending', 'Completed')),
    is_read BOOLEAN DEFAULT FALSE,
    transaction_date TIMESTAMP WITH TIME ZONE,
    processed_amount NUMERIC,
    notes TEXT,
    proof_file_id TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- Enable RLS
ALTER TABLE public.account_compensation_logs ENABLE ROW LEVEL SECURITY;

-- Policies
-- Note: Using public access for now as the app uses a custom session management
-- but we keep the structure for future Supabase Auth integration.
CREATE POLICY "Allow public read compensation logs" ON public.account_compensation_logs
    FOR SELECT TO public USING (true);

CREATE POLICY "Allow public insert compensation logs" ON public.account_compensation_logs
    FOR INSERT TO public WITH CHECK (true);

CREATE POLICY "Allow public update compensation logs" ON public.account_compensation_logs
    FOR UPDATE TO public USING (true);

CREATE POLICY "Allow public delete compensation logs" ON public.account_compensation_logs
    FOR DELETE TO public USING (true);

-- Grant access
GRANT ALL ON TABLE public.account_compensation_logs TO service_role;
GRANT ALL ON TABLE public.account_compensation_logs TO anon;
GRANT ALL ON TABLE public.account_compensation_logs TO authenticated;
