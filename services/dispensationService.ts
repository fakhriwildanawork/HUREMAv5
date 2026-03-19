import { supabase } from '../lib/supabase';
import { DispensationRequest, DispensationRequestInput } from '../types';

export const dispensationService = {
  async getAll() {
    const { data, error } = await supabase
      .from('dispensation_requests')
      .select('*, account:accounts(full_name, internal_nik)')
      .order('created_at', { ascending: false });
    if (error) throw error;
    return data as DispensationRequest[];
  },

  async getByAccountId(accountId: string) {
    const { data, error } = await supabase
      .from('dispensation_requests')
      .select('*')
      .eq('account_id', accountId)
      .order('created_at', { ascending: false });
    if (error) throw error;
    return data as DispensationRequest[];
  },

  async getById(id: string) {
    const { data, error } = await supabase
      .from('dispensation_requests')
      .select('*, account:accounts(full_name, internal_nik)')
      .eq('id', id)
      .single();
    if (error) throw error;
    return data as DispensationRequest;
  },

  async create(input: DispensationRequestInput) {
    const { data, error } = await supabase
      .from('dispensation_requests')
      .insert([input])
      .select();
    if (error) throw error;
    return data[0] as DispensationRequest;
  },

  async update(id: string, input: Partial<DispensationRequest>) {
    const { data, error } = await supabase
      .from('dispensation_requests')
      .update({ ...input, updated_at: new Date().toISOString() })
      .eq('id', id)
      .select();
    if (error) throw error;
    return data[0] as DispensationRequest;
  },

  async delete(id: string) {
    const { error } = await supabase
      .from('dispensation_requests')
      .delete()
      .eq('id', id);
    if (error) throw error;
  },

  async updateStatus(id: string, status: DispensationRequest['status'], issues: DispensationRequest['issues']) {
    const { data, error } = await supabase
      .from('dispensation_requests')
      .update({ status, issues, updated_at: new Date().toISOString() })
      .eq('id', id)
      .select();
    if (error) throw error;
    return data[0] as DispensationRequest;
  },

  async markAsRead(id: string) {
    const { error } = await supabase
      .from('dispensation_requests')
      .update({ is_read: true })
      .eq('id', id);
    if (error) throw error;
  },

  async getUnreadCount() {
    const { count, error } = await supabase
      .from('dispensation_requests')
      .select('*', { count: 'exact', head: true })
      .eq('is_read', false);
    if (error) throw error;
    return count || 0;
  },

  /**
   * Mendapatkan pengajuan dispensasi dalam rentang tanggal
   */
  async getByRange(accountId: string, startDate: string, endDate: string): Promise<DispensationRequest[]> {
    const { data, error } = await supabase
      .from('dispensation_requests')
      .select('*')
      .eq('account_id', accountId)
      .neq('status', 'REJECTED')
      .gte('date', startDate)
      .lte('date', endDate);
    
    if (error) throw error;
    return data || [];
  }
};
