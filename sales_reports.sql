-- Create sales_reports table
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

DROP TABLE IF EXISTS public.sales_reports CASCADE;

CREATE TABLE public.sales_reports (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    account_id UUID NOT NULL REFERENCES public.accounts(id) ON DELETE CASCADE,
    customer_name TEXT NOT NULL,
    activity_type TEXT NOT NULL,
    description TEXT,
    latitude DOUBLE PRECISION NOT NULL,
    longitude DOUBLE PRECISION NOT NULL,
    address TEXT,
    photo_urls JSONB DEFAULT '[]',
    file_ids JSONB DEFAULT '[]',
    reported_at TIMESTAMPTZ DEFAULT NOW(),
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Disable RLS for now to match the app's custom auth architecture (similar to kpi table)
ALTER TABLE public.sales_reports DISABLE ROW LEVEL SECURITY;

-- Grant access to public (anon)
GRANT ALL ON TABLE public.sales_reports TO anon;
GRANT ALL ON TABLE public.sales_reports TO authenticated;
GRANT ALL ON TABLE public.sales_reports TO service_role;

-- Create index for performance
CREATE INDEX IF NOT EXISTS idx_sales_reports_account_id ON public.sales_reports(account_id);
CREATE INDEX IF NOT EXISTS idx_sales_reports_reported_at ON public.sales_reports(reported_at);
