import { supabase } from '../lib/supabase';
import { KPI, KPIInput } from '../types';

export const kpiService = {
  async getAll(): Promise<KPI[]> {
    const { data, error } = await supabase
      .from('account_kpis')
      .select('*, account:accounts(full_name, internal_nik)')
      .order('created_at', { ascending: false });
    
    if (error) throw error;
    return data || [];
  },

  async getByAccountId(accountId: string): Promise<KPI[]> {
    const { data, error } = await supabase
      .from('account_kpis')
      .select('*')
      .eq('account_id', accountId)
      .order('deadline', { ascending: true });
    
    if (error) throw error;
    return data || [];
  },

  async create(input: KPIInput): Promise<KPI> {
    const { data, error } = await supabase
      .from('account_kpis')
      .insert([{ ...input, status: 'Active' }])
      .select()
      .single();
    
    if (error) throw error;
    return data;
  },

  async update(id: string, input: Partial<KPIInput>): Promise<KPI> {
    const { data, error } = await supabase
      .from('account_kpis')
      .update(input)
      .eq('id', id)
      .select()
      .single();
    
    if (error) throw error;
    return data;
  },

  async delete(id: string): Promise<void> {
    const { error } = await supabase
      .from('account_kpis')
      .delete()
      .eq('id', id);
    
    if (error) throw error;
  },

  async report(id: string, reportData: { description: string, file_ids: string[], links: string[], self_assessment: number }): Promise<void> {
    const reportedAt = new Date().toISOString();
    const { data: current } = await supabase.from('account_kpis').select('status, deadline').eq('id', id).single();
    
    let status = 'Unverified';
    // Jika status sebelumnya Unreported, tetap jadi Unverified tapi kita bisa tandai di data laporan
    
    const { error } = await supabase
      .from('account_kpis')
      .update({
        status,
        report_data: {
          ...reportData,
          reported_at: reportedAt
        },
        updated_at: reportedAt
      })
      .eq('id', id);
    
    if (error) throw error;
  },

  async verify(id: string, verifierId: string, verificationData: { score: number, notes: string }): Promise<void> {
    const verifiedAt = new Date().toISOString();
    const { error } = await supabase
      .from('account_kpis')
      .update({
        status: 'Verified',
        verification_data: {
          ...verificationData,
          verified_at: verifiedAt,
          verifier_id: verifierId
        },
        updated_at: verifiedAt
      })
      .eq('id', id);
    
    if (error) throw error;
  },

  async updateStatus(id: string, status: KPI['status']): Promise<void> {
    const { error } = await supabase
      .from('account_kpis')
      .update({ status, updated_at: new Date().toISOString() })
      .eq('id', id);
    
    if (error) throw error;
  }
};
