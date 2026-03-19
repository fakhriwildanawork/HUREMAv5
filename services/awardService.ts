
import { supabase } from '../lib/supabase';
import { EmployeeOfThePeriod, EmployeeOfThePeriodInput } from '../types';

export const awardService = {
  async getEmployeeOfThePeriodAll() {
    const { data, error } = await supabase
      .from('employee_of_the_period')
      .select('*')
      .order('year', { ascending: false })
      .order('month', { ascending: false });
    
    if (error) throw error;

    // Fetch account details for each record
    const results = await Promise.all(data.map(async (item: any) => {
      const { data: accounts, error: accError } = await supabase
        .from('accounts')
        .select('id, full_name, internal_nik, photo_google_id, position')
        .in('id', item.account_ids);
      
      if (accError) throw accError;
      return { ...item, accounts } as EmployeeOfThePeriod;
    }));

    return results;
  },

  async getLatestEmployeeOfThePeriod() {
    const { data, error } = await supabase
      .from('employee_of_the_period')
      .select('*')
      .order('year', { ascending: false })
      .order('month', { ascending: false })
      .limit(1)
      .maybeSingle();
    
    if (error) throw error;
    if (!data) return null;

    const { data: accounts, error: accError } = await supabase
      .from('accounts')
      .select('id, full_name, internal_nik, photo_google_id, position')
      .in('id', data.account_ids);
    
    if (accError) throw accError;
    return { ...data, accounts } as EmployeeOfThePeriod;
  },

  async createEmployeeOfThePeriod(input: EmployeeOfThePeriodInput) {
    const { data, error } = await supabase
      .from('employee_of_the_period')
      .insert([input])
      .select();
    
    if (error) throw error;
    return data[0] as EmployeeOfThePeriod;
  },

  async deleteEmployeeOfThePeriod(id: string) {
    const { error } = await supabase
      .from('employee_of_the_period')
      .delete()
      .eq('id', id);
    
    if (error) throw error;
    return true;
  }
};
