
import { supabase } from '../lib/supabase';
import { Overtime, OvertimeInput } from '../types';

const sanitizePayload = (payload: any) => {
  const sanitized = { ...payload };
  Object.keys(sanitized).forEach(key => {
    if (sanitized[key] === '' || sanitized[key] === undefined) {
      sanitized[key] = null;
    }
  });
  return sanitized;
};

export const overtimeService = {
  async getTodayOvertime(accountId: string) {
    const today = new Date().toISOString().split('T')[0];
    const { data, error } = await supabase
      .from('overtimes')
      .select('*')
      .eq('account_id', accountId)
      .gte('created_at', `${today}T00:00:00Z`)
      .order('created_at', { ascending: false })
      .limit(1)
      .maybeSingle();
    
    if (error) throw error;
    return data as Overtime | null;
  },

  async checkIn(input: Partial<OvertimeInput>) {
    const sanitized = sanitizePayload(input);
    const { data, error } = await supabase
      .from('overtimes')
      .insert([sanitized])
      .select()
      .single();
    
    if (error) throw error;
    return data as Overtime;
  },

  async checkOut(id: string, input: Partial<OvertimeInput>) {
    if (!id) throw new Error("ID lembur tidak valid untuk proses check-out.");
    const sanitized = sanitizePayload(input);
    const { data, error } = await supabase
      .from('overtimes')
      .update(sanitized)
      .eq('id', id)
      .select()
      .single();
    
    if (error) throw error;
    return data as Overtime;
  },

  async getRecentHistory(accountId: string, limit = 31) {
    const { data, error } = await supabase
      .from('overtimes')
      .select('*')
      .eq('account_id', accountId)
      .order('created_at', { ascending: false })
      .limit(limit);
    
    if (error) throw error;
    return data as Overtime[];
  },
  
  async getOvertimeByRange(startDate: string, endDate: string) {
    const { data, error } = await supabase
      .from('overtimes')
      .select('*')
      .gte('created_at', `${startDate}T00:00:00Z`)
      .lte('created_at', `${endDate}T23:59:59Z`)
      .order('created_at', { ascending: true });
    
    if (error) throw error;
    return data as Overtime[];
  }
};
