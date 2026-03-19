-- Digital Asset Management Tables

-- 1. Master Documents Table
CREATE TABLE IF NOT EXISTS documents (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    name TEXT NOT NULL,
    doc_type TEXT NOT NULL, -- SOP, Kebijakan, Form, dsb
    file_id TEXT NOT NULL, -- Google Drive ID
    description TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- 2. Document Access Control (Junction Table)
CREATE TABLE IF NOT EXISTS document_access (
    document_id UUID REFERENCES documents(id) ON DELETE CASCADE,
    account_id UUID REFERENCES accounts(id) ON DELETE CASCADE,
    PRIMARY KEY (document_id, account_id)
);

-- Enable RLS
ALTER TABLE documents ENABLE ROW LEVEL SECURITY;
ALTER TABLE document_access ENABLE ROW LEVEL SECURITY;

-- Public Policies
CREATE POLICY "Allow public read documents" ON documents FOR SELECT TO public USING (true);
CREATE POLICY "Allow public insert documents" ON documents FOR INSERT TO public WITH CHECK (true);
CREATE POLICY "Allow public update documents" ON documents FOR UPDATE TO public USING (true);
CREATE POLICY "Allow public delete documents" ON documents FOR DELETE TO public USING (true);

CREATE POLICY "Allow public read document_access" ON document_access FOR SELECT TO public USING (true);
CREATE POLICY "Allow public insert document_access" ON document_access FOR INSERT TO public WITH CHECK (true);
CREATE POLICY "Allow public delete document_access" ON document_access FOR DELETE TO public USING (true);

-- Trigger for updated_at
CREATE TRIGGER set_updated_at_documents
BEFORE UPDATE ON documents
FOR EACH ROW
EXECUTE FUNCTION handle_updated_at();