import { supabase } from '../lib/supabase';
import { LeaveRequest, LeaveRequestInput, LeaveRequestExtended, AnnualLeaveRequest, AnnualLeaveRequestInput } from '../types';
import { settingsService } from './settingsService';
import { authService } from './authService';

export const leaveService = {
  /**
   * Mendapatkan semua pengajuan cuti tahunan untuk satu akun
   */
  async getAnnualByAccountId(accountId: string): Promise<AnnualLeaveRequest[]> {
    const { data, error } = await supabase
      .from('account_annual_leaves')
      .select('*')
      .eq('account_id', accountId)
      .order('created_at', { ascending: false });
    
    if (error) {
      console.error('Error fetching annual leaves:', error);
      throw error;
    }
    return data || [];
  },

  /**
   * Mendapatkan semua pengajuan cuti tahunan (Admin)
   */
  async getAllAnnual(): Promise<AnnualLeaveRequest[]> {
    const { data, error } = await supabase
      .from('account_annual_leaves')
      .select('*, account:accounts(full_name, internal_nik)')
      .order('created_at', { ascending: false });
    
    if (error) {
      console.error('Error fetching all annual leaves:', error);
      throw error;
    }
    return data || [];
  },

  /**
   * Membuat pengajuan cuti tahunan baru
   */
  async createAnnual(input: AnnualLeaveRequestInput): Promise<AnnualLeaveRequest> {
    // Validasi sesi kustom sebelum kirim
    const currentUser = authService.getCurrentUser();
    if (!currentUser) {
      console.error('No active custom session found during createAnnual');
      throw new Error('Sesi Anda telah berakhir. Silakan login kembali.');
    }

    console.log('Creating annual leave with payload:', input);

    const { data, error } = await supabase
      .from('account_annual_leaves')
      .insert({
        ...input,
        status: 'pending',
        current_negotiator_role: 'admin',
        negotiation_data: [{
          role: 'user',
          start_date: input.start_date,
          end_date: input.end_date,
          reason: input.description,
          timestamp: new Date().toISOString()
        }]
      })
      .select()
      .single();
    
    if (error) {
      console.error('Supabase error in createAnnual:', error);
      throw error;
    }

    // Sinkronisasi ke tabel submissions agar muncul di daftar verifikasi pusat
    try {
      await supabase.from('account_submissions').insert([{
        account_id: input.account_id,
        type: 'Cuti Tahunan',
        status: 'Pending',
        description: input.description,
        submission_data: {
          start_date: input.start_date,
          end_date: input.end_date,
          annual_leave_id: data.id
        }
      }]);
    } catch (subError) {
      console.warn('Failed to sync to submissions table:', subError);
      // We don't throw here to avoid blocking the main leave creation
    }

    return data;
  },



  /**
   * Negosiasi / Respon Cuti (Admin atau User)
   */
  async negotiateAnnual(
    id: string, 
    role: 'admin' | 'user', 
    startDate: string, 
    endDate: string, 
    reason: string,
    status: 'negotiating' | 'approved' | 'rejected' | 'cancelled'
  ): Promise<void> {
    // Validasi sesi kustom
    const currentUser = authService.getCurrentUser();
    if (!currentUser) {
      throw new Error('Sesi Anda telah berakhir. Silakan login kembali.');
    }

    const { data: current, error: fetchError } = await supabase
      .from('account_annual_leaves')
      .select('negotiation_data, account_id')
      .eq('id', id)
      .single();

    if (fetchError) {
      console.error('Error fetching current negotiation data:', fetchError);
      throw fetchError;
    }

    const newHistory = [
      ...(current?.negotiation_data || []),
      {
        role,
        start_date: startDate,
        end_date: endDate,
        reason,
        timestamp: new Date().toISOString()
      }
    ];

    const { error } = await supabase
      .from('account_annual_leaves')
      .update({
        status,
        start_date: startDate,
        end_date: endDate,
        negotiation_data: newHistory,
        current_negotiator_role: role === 'admin' ? 'user' : 'admin',
        updated_at: new Date().toISOString()
      })
      .eq('id', id);
    
    if (error) {
      console.error('Error updating annual leave negotiation:', error);
      throw error;
    }

    // Update status di tabel submissions pusat
    try {
      const submissionStatus = status === 'approved' ? 'Disetujui' : status === 'rejected' ? 'Ditolak' : status === 'cancelled' ? 'Dibatalkan' : 'Pending';
      
      // Cari submission yang berkaitan
      const { data: sub } = await supabase
        .from('account_submissions')
        .select('id')
        .eq('account_id', current.account_id)
        .eq('type', 'Cuti Tahunan')
        .contains('submission_data', { annual_leave_id: id })
        .maybeSingle();

      if (sub) {
        await supabase.from('account_submissions')
          .update({ 
            status: submissionStatus,
            description: reason,
            submission_data: {
              start_date: startDate,
              end_date: endDate,
              annual_leave_id: id
            }
          })
          .eq('id', sub.id);
      }
    } catch (subError) {
      console.warn('Failed to update central submission status:', subError);
    }
  },

  /**
   * Menghapus pengajuan cuti
   */
  async deleteAnnual(id: string): Promise<void> {
    const { error } = await supabase
      .from('account_annual_leaves')
      .delete()
      .eq('id', id);
    
    if (error) throw error;
  },

  /**
   * Menghitung total hari cuti tahunan yang sudah disetujui untuk satu akun di tahun tertentu
   */
  async getUsedAnnualLeaveDays(accountId: string, year: number): Promise<number> {
    const startDate = `${year}-01-01`;
    const endDate = `${year}-12-31`;

    const { data, error } = await supabase
      .from('account_annual_leaves')
      .select('start_date, end_date')
      .eq('account_id', accountId)
      .eq('status', 'approved')
      .gte('start_date', startDate)
      .lte('start_date', endDate);

    if (error) {
      console.error('Error fetching used annual leave days:', error);
      throw error;
    }

    if (!data) return 0;

    return data.reduce((total, leave) => {
      const start = new Date(leave.start_date);
      const end = new Date(leave.end_date);
      const diffTime = Math.abs(end.getTime() - start.getTime());
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
      return total + diffDays;
    }, 0);
  },

  /**
   * Mendapatkan semua pengajuan libur untuk satu akun
   */
  async getByAccountId(accountId: string): Promise<LeaveRequest[]> {
    const { data, error } = await supabase
      .from('account_leave_requests')
      .select('*')
      .eq('account_id', accountId)
      .order('start_date', { ascending: false });
    
    if (error) throw error;
    return data || [];
  },

  /**
   * Mendapatkan semua pengajuan libur (untuk admin)
   */
  async getAll(): Promise<LeaveRequestExtended[]> {
    const { data, error } = await supabase
      .from('account_leave_requests')
      .select('*, account:accounts(full_name, internal_nik)')
      .order('created_at', { ascending: false });
    
    if (error) throw error;
    return data || [];
  },

  /**
   * Membuat pengajuan libur baru
   */
  async create(input: LeaveRequestInput): Promise<LeaveRequest> {
    // Ambil kebijakan dari settings
    const policy = await settingsService.getSetting('leave_approval_policy', 'manual');
    const status = policy === 'auto' ? 'approved' : 'pending';

    // 1. Simpan ke tabel khusus libur mandiri
    const { data, error } = await supabase
      .from('account_leave_requests')
      .insert({
        ...input,
        status
      })
      .select()
      .single();
    
    if (error) throw error;

    // 2. Sinkronisasi ke tabel submissions agar muncul di daftar verifikasi pusat
    const submissionStatus = status === 'approved' ? 'Disetujui' : 'Pending';
    await supabase.from('account_submissions').insert([{
      account_id: input.account_id,
      type: 'Libur Mandiri',
      status: submissionStatus,
      description: input.description,
      submission_data: {
        start_date: input.start_date,
        end_date: input.end_date,
        leave_request_id: data.id // Simpan ID referensi untuk sinkronisasi balik
      }
    }]);

    return data;
  },

  /**
   * Memperbarui status pengajuan libur
   */
  async updateStatus(id: string, status: 'approved' | 'rejected'): Promise<void> {
    const { error } = await supabase
      .from('account_leave_requests')
      .update({ 
        status,
        updated_at: new Date().toISOString()
      })
      .eq('id', id);
    
    if (error) throw error;
  },

  /**
   * Menghapus pengajuan libur
   */
  async delete(id: string): Promise<void> {
    const { error } = await supabase
      .from('account_leave_requests')
      .delete()
      .eq('id', id);
    
    if (error) throw error;
  },

  /**
   * Mendapatkan pengajuan cuti tahunan dalam rentang tanggal
   */
  async getAnnualByRange(accountId: string, startDate: string, endDate: string): Promise<AnnualLeaveRequest[]> {
    const { data, error } = await supabase
      .from('account_annual_leaves')
      .select('*')
      .eq('account_id', accountId)
      .neq('status', 'rejected')
      .neq('status', 'cancelled')
      .or(`start_date.lte.${endDate},end_date.gte.${startDate}`);
    
    if (error) throw error;
    return data || [];
  },

  /**
   * Mendapatkan pengajuan libur mandiri dalam rentang tanggal
   */
  async getMandatoryByRange(accountId: string, startDate: string, endDate: string): Promise<LeaveRequest[]> {
    const { data, error } = await supabase
      .from('account_leave_requests')
      .select('*')
      .eq('account_id', accountId)
      .neq('status', 'rejected')
      .or(`start_date.lte.${endDate},end_date.gte.${startDate}`);
    
    if (error) throw error;
    return data || [];
  }
};
