import { supabase } from '../lib/supabase';
import { Feedback, FeedbackInput } from '../types';

export const feedbackService = {
  async getAll(): Promise<Feedback[]> {
    const { data, error } = await supabase
      .from('account_feedbacks')
      .select('*, account:accounts(full_name, internal_nik)')
      .order('created_at', { ascending: false });
    
    if (error) throw error;
    return data || [];
  },

  async getByAccountId(accountId: string): Promise<Feedback[]> {
    const { data, error } = await supabase
      .from('account_feedbacks')
      .select('*')
      .eq('account_id', accountId)
      .order('created_at', { ascending: false });
    
    if (error) throw error;
    return data || [];
  },

  async submitFeedback(input: FeedbackInput): Promise<void> {
    const { error } = await supabase
      .from('account_feedbacks')
      .insert([input]);
    
    if (error) throw error;
  },

  async markAsRead(id: string): Promise<void> {
    const { error } = await supabase
      .from('account_feedbacks')
      .update({ status: 'Read' })
      .eq('id', id);
    
    if (error) throw error;
  },

  async delete(id: string): Promise<void> {
    const { error } = await supabase
      .from('account_feedbacks')
      .delete()
      .eq('id', id);
    
    if (error) throw error;
  }
};
