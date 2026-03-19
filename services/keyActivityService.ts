import { supabase } from '../lib/supabase';
import { KeyActivity, KeyActivityInput, KeyActivityReport, KeyActivityReportInput } from '../types';

export const keyActivityService = {
  async getAll(): Promise<KeyActivity[]> {
    const { data, error } = await supabase
      .from('key_activities')
      .select('*, assignments:key_activity_assignments(account_id, account:accounts(full_name, internal_nik))')
      .order('created_at', { ascending: false });
    
    if (error) throw error;
    return data || [];
  },

  async getByAccountId(accountId: string): Promise<KeyActivity[]> {
    const { data, error } = await supabase
      .from('key_activity_assignments')
      .select('activity:key_activities(*)')
      .eq('account_id', accountId);
    
    if (error) throw error;
    return (data as any[])?.map(d => d.activity) || [];
  },

  async create(input: KeyActivityInput): Promise<void> {
    const { assigned_account_ids, ...activityData } = input;
    
    const { data: activity, error: activityError } = await supabase
      .from('key_activities')
      .insert([activityData])
      .select()
      .single();
    
    if (activityError) throw activityError;

    if (assigned_account_ids.length > 0) {
      const assignments = assigned_account_ids.map(accountId => ({
        activity_id: activity.id,
        account_id: accountId
      }));
      
      const { error: assignError } = await supabase
        .from('key_activity_assignments')
        .insert(assignments);
      
      if (assignError) throw assignError;
    }
  },

  async update(id: string, input: KeyActivityInput): Promise<void> {
    const { assigned_account_ids, ...activityData } = input;
    
    const { error: activityError } = await supabase
      .from('key_activities')
      .update(activityData)
      .eq('id', id);
    
    if (activityError) throw activityError;

    // Refresh assignments
    await supabase.from('key_activity_assignments').delete().eq('activity_id', id);
    
    if (assigned_account_ids.length > 0) {
      const assignments = assigned_account_ids.map(accountId => ({
        activity_id: id,
        account_id: accountId
      }));
      
      const { error: assignError } = await supabase
        .from('key_activity_assignments')
        .insert(assignments);
      
      if (assignError) throw assignError;
    }
  },

  async delete(id: string): Promise<void> {
    const { error } = await supabase
      .from('key_activities')
      .delete()
      .eq('id', id);
    
    if (error) throw error;
  },

  async getReportsByActivity(activityId: string): Promise<KeyActivityReport[]> {
    const { data, error } = await supabase
      .from('key_activity_reports')
      .select('*, account:accounts(full_name, internal_nik)')
      .eq('activity_id', activityId)
      .order('due_date', { ascending: false });
    
    if (error) throw error;
    return data || [];
  },

  async getReportsByAccount(accountId: string): Promise<KeyActivityReport[]> {
    const { data, error } = await supabase
      .from('key_activity_reports')
      .select('*, activity:key_activities(*)')
      .eq('account_id', accountId)
      .order('due_date', { ascending: false });
    
    if (error) throw error;
    return data || [];
  },

  async getAllReports(): Promise<KeyActivityReport[]> {
    const { data, error } = await supabase
      .from('key_activity_reports')
      .select('*, activity:key_activities(*), account:accounts(full_name, internal_nik)')
      .order('reported_at', { ascending: false });
    
    if (error) throw error;
    return data || [];
  },

  async submitReport(input: KeyActivityReportInput): Promise<void> {
    const { error } = await supabase
      .from('key_activity_reports')
      .insert([{ ...input, status: 'Unverified', reported_at: new Date().toISOString() }]);
    
    if (error) throw error;
  },

  async verifyReport(id: string, verifierId: string, score: number, notes: string): Promise<void> {
    const { error } = await supabase
      .from('key_activity_reports')
      .update({
        status: 'Verified',
        verification_data: {
          score,
          notes,
          verified_at: new Date().toISOString(),
          verifier_id: verifierId
        }
      })
      .eq('id', id);
    
    if (error) throw error;
  },

  // Helper to generate expected due dates for an activity up to a certain date
  generateDueDates(activity: KeyActivity, untilDate: string): string[] {
    const dates: string[] = [];
    const start = new Date(activity.start_date + 'T00:00:00');
    const end = new Date(activity.end_date + 'T00:00:00');
    const until = new Date(untilDate + 'T00:00:00');
    const limit = until < end ? until : end;

    let current = new Date(start);
    // Reset time to avoid issues with comparisons
    current.setHours(0, 0, 0, 0);
    limit.setHours(0, 0, 0, 0);

    while (current <= limit) {
      const dateStr = current.toLocaleDateString('en-CA');
      
      let shouldInclude = false;
      switch (activity.recurrence_type) {
        case 'Once':
          shouldInclude = dateStr === activity.start_date;
          if (shouldInclude) {
             dates.push(dateStr);
             return dates;
          }
          break;
        case 'Daily':
          shouldInclude = true;
          break;
        case 'Weekly':
          const day = current.getDay();
          shouldInclude = activity.recurrence_rule?.days_of_week?.includes(day) || false;
          break;
        case 'Monthly':
          const date = current.getDate();
          shouldInclude = activity.recurrence_rule?.dates_of_month?.includes(date) || false;
          break;
        case 'EndOfMonth':
          const nextDay = new Date(current);
          nextDay.setDate(current.getDate() + 1);
          shouldInclude = nextDay.getMonth() !== current.getMonth();
          break;
      }

      if (shouldInclude) {
        dates.push(dateStr);
      }
      
      current.setDate(current.getDate() + 1);
    }
    return dates;
  }
};
