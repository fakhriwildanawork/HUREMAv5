import { supabase } from '../lib/supabase';
import { Whistleblowing, WhistleblowingInput } from '../types';

export const whistleblowingService = {
  async getAll(): Promise<Whistleblowing[]> {
    const { data, error } = await supabase
      .from('account_reports')
      .select('*, account:accounts(full_name, internal_nik)')
      .order('created_at', { ascending: false });
    
    if (error) throw error;

    // Fetch reported accounts details for each report
    const reportsWithDetails = await Promise.all((data || []).map(async (report) => {
      if (report.reported_account_ids && report.reported_account_ids.length > 0) {
        const { data: accounts } = await supabase
          .from('accounts')
          .select('id, full_name, internal_nik')
          .in('id', report.reported_account_ids);
        return { ...report, reported_accounts: accounts || [] };
      }
      return { ...report, reported_accounts: [] };
    }));

    return reportsWithDetails;
  },

  async getByAccountId(accountId: string): Promise<Whistleblowing[]> {
    const { data, error } = await supabase
      .from('account_reports')
      .select('*, account:accounts(full_name, internal_nik)')
      .eq('account_id', accountId)
      .order('created_at', { ascending: false });
    
    if (error) throw error;

    // Fetch reported accounts details for each report
    const reportsWithDetails = await Promise.all((data || []).map(async (report) => {
      if (report.reported_account_ids && report.reported_account_ids.length > 0) {
        const { data: accounts } = await supabase
          .from('accounts')
          .select('id, full_name, internal_nik')
          .in('id', report.reported_account_ids);
        return { ...report, reported_accounts: accounts || [] };
      }
      return { ...report, reported_accounts: [] };
    }));

    return reportsWithDetails;
  },

  async submitReport(input: WhistleblowingInput): Promise<void> {
    const { error } = await supabase
      .from('account_reports')
      .insert([input]);
    
    if (error) throw error;
  },

  async markAsRead(id: string): Promise<void> {
    const { error } = await supabase
      .from('account_reports')
      .update({ status: 'Read' })
      .eq('id', id);
    
    if (error) throw error;
  },

  async delete(id: string): Promise<void> {
    const { error } = await supabase
      .from('account_reports')
      .delete()
      .eq('id', id);
    
    if (error) throw error;
  }
};
