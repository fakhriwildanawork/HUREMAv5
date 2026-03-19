import { supabase } from '../lib/supabase';
import { SalesReport, SalesReportInput } from '../types';

export const salesReportService = {
  async getAll(): Promise<SalesReport[]> {
    const { data, error } = await supabase
      .from('sales_reports')
      .select('*, account:accounts(full_name, internal_nik)')
      .order('reported_at', { ascending: false });
    
    if (error) throw error;
    return data || [];
  },

  async getByAccountId(accountId: string): Promise<SalesReport[]> {
    const { data, error } = await supabase
      .from('sales_reports')
      .select('*')
      .eq('account_id', accountId)
      .order('reported_at', { ascending: false });
    
    if (error) throw error;
    return data || [];
  },

  async submitReport(input: SalesReportInput): Promise<void> {
    const { error } = await supabase
      .from('sales_reports')
      .insert([input]);
    
    if (error) throw error;
  },

  async delete(id: string): Promise<void> {
    const { error } = await supabase
      .from('sales_reports')
      .delete()
      .eq('id', id);
    
    if (error) throw error;
  }
};
