import { supabase } from '../lib/supabase';
import { SalaryScheme, SalarySchemeInput, SalaryAssignment, SalaryAssignmentExtended, Reimbursement, ReimbursementInput, ReimbursementStatus, SalaryAdjustment, SalaryAdjustmentInput, PayrollStatus, Payroll, PayrollItem, PayrollSettings, EarlySalaryRequest, EarlySalaryRequestInput, EarlySalaryStatus, Compensation, CompensationInput, CompensationStatus } from '../types';

export const financeService = {
  // Early Salary Requests
  async getEarlySalaryRequests(filters?: { account_id?: string, month?: number, year?: number }) {
    let query = supabase
      .from('finance_early_salary_requests')
      .select(`
        *,
        account:accounts!finance_early_salary_requests_account_id_fkey(full_name, internal_nik)
      `)
      .order('created_at', { ascending: false });

    if (filters?.account_id) query = query.eq('account_id', filters.account_id);
    if (filters?.month) query = query.eq('month', filters.month);
    if (filters?.year) query = query.eq('year', filters.year);

    const { data, error } = await query;
    if (error) throw error;
    return data as EarlySalaryRequest[];
  },

  async createEarlySalaryRequest(request: EarlySalaryRequestInput & { account_id: string }) {
    const { data, error } = await supabase
      .from('finance_early_salary_requests')
      .insert([request])
      .select();
    
    if (error) throw error;
    return data[0] as EarlySalaryRequest;
  },

  async updateEarlySalaryStatus(id: string, update: { 
    status: EarlySalaryStatus, 
    payment_proof_id?: string | null,
    rejection_reason?: string | null,
    verifier_id: string 
  }) {
    const { data, error } = await supabase
      .from('finance_early_salary_requests')
      .update({
        ...update,
        verified_at: new Date().toISOString(),
        updated_at: new Date().toISOString()
      })
      .eq('id', id)
      .select();
    
    if (error) throw error;
    return data[0] as EarlySalaryRequest;
  },

  // Salary Schemes
  async getSchemes() {
    const { data, error } = await supabase
      .from('finance_salary_schemes')
      .select('*')
      .order('created_at', { ascending: false });
    
    if (error) throw error;
    return data as SalaryScheme[];
  },

  async getSchemeById(id: string) {
    const { data, error } = await supabase
      .from('finance_salary_schemes')
      .select('*')
      .eq('id', id)
      .single();
    
    if (error) throw error;
    return data as SalaryScheme;
  },

  async createScheme(scheme: SalarySchemeInput) {
    const { data, error } = await supabase
      .from('finance_salary_schemes')
      .insert([scheme])
      .select();
    
    if (error) throw error;
    return data[0] as SalaryScheme;
  },

  async updateScheme(id: string, scheme: Partial<SalarySchemeInput>) {
    const { data, error } = await supabase
      .from('finance_salary_schemes')
      .update(scheme)
      .eq('id', id)
      .select();
    
    if (error) throw error;
    return data[0] as SalaryScheme;
  },

  async deleteScheme(id: string) {
    const { error } = await supabase
      .from('finance_salary_schemes')
      .delete()
      .eq('id', id);
    
    if (error) throw error;
    return true;
  },

  // Salary Assignments
  async getAssignments() {
    const { data, error } = await supabase
      .from('finance_salary_assignments')
      .select(`
        *,
        account:accounts(full_name, internal_nik, position, grade, location_id),
        scheme:finance_salary_schemes(*)
      `)
      .order('created_at', { ascending: false });
    
    if (error) throw error;
    return data as SalaryAssignmentExtended[];
  },

  async assignScheme(schemeId: string, accountIds: string[]) {
    // We use upsert because account_id is unique
    const assignments = accountIds.map(accountId => ({
      scheme_id: schemeId,
      account_id: accountId,
      updated_at: new Date().toISOString()
    }));

    const { data, error } = await supabase
      .from('finance_salary_assignments')
      .upsert(assignments, { onConflict: 'account_id' })
      .select();
    
    if (error) throw error;
    return data;
  },

  async removeAssignment(id: string) {
    const { error } = await supabase
      .from('finance_salary_assignments')
      .delete()
      .eq('id', id);
    
    if (error) throw error;
    return true;
  },

  async getAssignmentByAccountId(accountId: string) {
    const { data, error } = await supabase
      .from('finance_salary_assignments')
      .select(`
        *,
        scheme:finance_salary_schemes(*)
      `)
      .eq('account_id', accountId)
      .maybeSingle();
    
    if (error) throw error;
    return data as SalaryAssignmentExtended | null;
  },

  // Reimbursements
  async getReimbursements(filters?: { account_id?: string, month?: number, year?: number }) {
    let query = supabase
      .from('finance_reimbursements')
      .select(`
        *,
        account:accounts!finance_reimbursements_account_id_fkey(full_name, internal_nik)
      `)
      .order('created_at', { ascending: false });

    if (filters?.account_id) {
      query = query.eq('account_id', filters.account_id);
    }

    if (filters?.month && filters?.year) {
      const startDate = new Date(filters.year, filters.month - 1, 1).toISOString();
      const endDate = new Date(filters.year, filters.month, 0, 23, 59, 59).toISOString();
      query = query.gte('transaction_date', startDate).lte('transaction_date', endDate);
    }

    const { data, error } = await query;
    if (error) throw error;
    return data as Reimbursement[];
  },

  async createReimbursement(reimbursement: ReimbursementInput & { account_id: string }) {
    const { data, error } = await supabase
      .from('finance_reimbursements')
      .insert([reimbursement])
      .select();
    
    if (error) throw error;
    return data[0] as Reimbursement;
  },

  async updateReimbursementStatus(id: string, update: { 
    status: ReimbursementStatus, 
    amount_approved?: number, 
    admin_notes?: string, 
    payment_proof_id?: string,
    verifier_id: string 
  }) {
    const { data, error } = await supabase
      .from('finance_reimbursements')
      .update({
        ...update,
        verified_at: new Date().toISOString(),
        updated_at: new Date().toISOString()
      })
      .eq('id', id)
      .select();
    
    if (error) throw error;
    return data[0] as Reimbursement;
  },

  async markAsRead(id: string) {
    const { error } = await supabase
      .from('finance_reimbursements')
      .update({ is_read: true })
      .eq('id', id);
    
    if (error) throw error;
    return true;
  },

  async getUnreadCount() {
    const { count, error } = await supabase
      .from('finance_reimbursements')
      .select('*', { count: 'exact', head: true })
      .eq('is_read', false);
    
    if (error) throw error;
    return count || 0;
  },

  // Salary Adjustments
  async getSalaryAdjustments(filters?: { month?: number, year?: number, account_id?: string }) {
    let query = supabase
      .from('finance_salary_adjustments')
      .select(`
        *,
        account:accounts(full_name, internal_nik)
      `)
      .order('created_at', { ascending: false });

    if (filters?.month) query = query.eq('month', filters.month);
    if (filters?.year) query = query.eq('year', filters.year);
    if (filters?.account_id) query = query.eq('account_id', filters.account_id);

    const { data, error } = await query;
    if (error) throw error;
    return data as SalaryAdjustment[];
  },

  async createSalaryAdjustment(adjustments: SalaryAdjustmentInput[]) {
    const { data, error } = await supabase
      .from('finance_salary_adjustments')
      .insert(adjustments)
      .select();
    
    if (error) throw error;
    return data;
  },

  async updateSalaryAdjustment(id: string, adjustment: Partial<SalaryAdjustmentInput>) {
    const { data, error } = await supabase
      .from('finance_salary_adjustments')
      .update({ ...adjustment, updated_at: new Date().toISOString() })
      .eq('id', id)
      .select();
    
    if (error) throw error;
    return data[0] as SalaryAdjustment;
  },

  async deleteSalaryAdjustment(id: string) {
    const { error } = await supabase
      .from('finance_salary_adjustments')
      .delete()
      .eq('id', id);
    
    if (error) throw error;
    return true;
  },

  // Payroll Status Check
  async getPayrollStatus(month: number, year: number) {
    const { data, error } = await supabase
      .from('finance_payroll_status')
      .select('*')
      .eq('month', month)
      .eq('year', year)
      .maybeSingle();
    
    if (error) throw error;
    return data as PayrollStatus | null;
  },

  // Payroll Management
  async getPayrolls() {
    const { data, error } = await supabase
      .from('finance_payrolls')
      .select(`
        *,
        verifier:accounts!finance_payrolls_verifier_id_fkey(full_name),
        creator:accounts!finance_payrolls_created_by_fkey(full_name)
      `)
      .order('year', { ascending: false })
      .order('month', { ascending: false });
    
    if (error) throw error;
    return data as Payroll[];
  },

  async createPayroll(payroll: Omit<Payroll, 'id' | 'created_at' | 'updated_at'>) {
    // Remove updated_by if it exists because the column might not exist in DB yet
    const { updated_by, ...insertData } = payroll as any;

    const { data, error } = await supabase
      .from('finance_payrolls')
      .insert([insertData])
      .select();
    
    if (error) throw error;
    return data[0] as Payroll;
  },

  async updatePayroll(id: string, payroll: Partial<Omit<Payroll, 'id' | 'created_at' | 'updated_at'>>) {
    // Remove updated_by if it exists in the object because the column might not exist in DB yet
    const { updated_by, ...updateData } = payroll as any;
    
    const { data, error } = await supabase
      .from('finance_payrolls')
      .update({ ...updateData, updated_at: new Date().toISOString() })
      .eq('id', id)
      .select();
    
    if (error) throw error;
    return data[0] as Payroll;
  },

  async updatePayrollStatus(id: string, status: Payroll['status'], verifierId?: string, notes?: string) {
    const updateData: any = { status, updated_at: new Date().toISOString() };
    if (verifierId) updateData.verifier_id = verifierId;
    if (notes) updateData.verification_notes = notes;
    if (status === 'Approved') updateData.verified_at = new Date().toISOString();

    const { data, error } = await supabase
      .from('finance_payrolls')
      .update(updateData)
      .eq('id', id)
      .select();
    
    if (error) throw error;
    return data[0] as Payroll;
  },

  async deletePayroll(id: string) {
    const { error } = await supabase
      .from('finance_payrolls')
      .delete()
      .eq('id', id);
    
    if (error) throw error;
    return true;
  },

  async getPayrollItems(payrollId: string) {
    const { data, error } = await supabase
      .from('finance_payroll_items')
      .select(`
        *,
        account:accounts(full_name, internal_nik, position, grade, location:locations(name))
      `)
      .eq('payroll_id', payrollId);
    
    if (error) throw error;
    return data as PayrollItem[];
  },

  async getEmployeePayslips(accountId: string) {
    const { data, error } = await supabase
      .from('finance_payroll_items')
      .select(`
        *,
        payroll:finance_payrolls!inner(*),
        account:accounts(full_name, internal_nik, position, grade, location:locations(name))
      `)
      .eq('account_id', accountId)
      .eq('payroll.status', 'Paid')
      .order('created_at', { ascending: false });
    
    if (error) throw error;
    return data as PayrollItem[];
  },

  async upsertPayrollItems(items: Omit<PayrollItem, 'id' | 'created_at' | 'updated_at' | 'account'>[]) {
    const { data, error } = await supabase
      .from('finance_payroll_items')
      .upsert(items, { onConflict: 'payroll_id,account_id' })
      .select();
    
    if (error) throw error;
    return data;
  },

  async updatePayrollItem(id: string, update: Partial<PayrollItem>) {
    const { data, error } = await supabase
      .from('finance_payroll_items')
      .update({ ...update, updated_at: new Date().toISOString() })
      .eq('id', id)
      .select();
    
    if (error) throw error;
    return data[0] as PayrollItem;
  },

  async deletePayrollItems(ids: string[]) {
    const { error } = await supabase
      .from('finance_payroll_items')
      .delete()
      .in('id', ids);
    
    if (error) throw error;
    return true;
  },

  async getPayrollSettings() {
    const { data, error } = await supabase
      .from('finance_payroll_settings')
      .select('*')
      .single();
    
    if (error) throw error;
    return data as PayrollSettings;
  },

  async updatePayrollSettings(settings: Partial<PayrollSettings>) {
    const { data, error } = await supabase
      .from('finance_payroll_settings')
      .update({ ...settings, updated_at: new Date().toISOString() })
      .eq('id', settings.id)
      .select();
    
    if (error) throw error;
    return data[0] as PayrollSettings;
  },

  // Compensations
  async getCompensations(filters?: { status?: CompensationStatus, account_id?: string }) {
    let query = supabase
      .from('account_compensation_logs')
      .select(`
        *,
        account:accounts(full_name, internal_nik)
      `)
      .order('created_at', { ascending: false });

    if (filters?.status) query = query.eq('status', filters.status);
    if (filters?.account_id) query = query.eq('account_id', filters.account_id);

    const { data, error } = await query;
    if (error) throw error;
    return data as Compensation[];
  },

  async createCompensation(compensation: CompensationInput) {
    const { data, error } = await supabase
      .from('account_compensation_logs')
      .insert([compensation])
      .select();
    
    if (error) throw error;
    return data[0] as Compensation;
  },

  async updateCompensation(id: string, update: Partial<Compensation>) {
    const { data, error } = await supabase
      .from('account_compensation_logs')
      .update({ ...update, updated_at: new Date().toISOString() })
      .eq('id', id)
      .select();
    
    if (error) throw error;
    return data[0] as Compensation;
  },

  async markCompensationAsRead(id: string) {
    const { error } = await supabase
      .from('account_compensation_logs')
      .update({ is_read: true })
      .eq('id', id);
    
    if (error) throw error;
    return true;
  },

  async getUnreadCompensationCount() {
    const { count, error } = await supabase
      .from('account_compensation_logs')
      .select('*', { count: 'exact', head: true })
      .eq('is_read', false);
    
    if (error) throw error;
    return count || 0;
  }
};
