import { supabase } from '../lib/supabase';
import { MaternityLeaveRequest, MaternityLeaveRequestInput } from '../types';
import { authService } from './authService';

export const maternityLeaveService = {
  /**
   * Mendapatkan semua pengajuan cuti melahirkan untuk satu akun
   */
  async getByAccountId(accountId: string): Promise<MaternityLeaveRequest[]> {
    const { data, error } = await supabase
      .from('account_maternity_leaves')
      .select('*')
      .eq('account_id', accountId)
      .order('created_at', { ascending: false });
    
    if (error) {
      console.error('Error fetching maternity leaves:', error);
      throw error;
    }
    return data || [];
  },

  /**
   * Mendapatkan semua pengajuan cuti melahirkan (Admin)
   */
  async getAll(): Promise<MaternityLeaveRequest[]> {
    const { data, error } = await supabase
      .from('account_maternity_leaves')
      .select('*, account:accounts(full_name, internal_nik)')
      .order('created_at', { ascending: false });
    
    if (error) {
      console.error('Error fetching all maternity leaves:', error);
      throw error;
    }
    return data || [];
  },

  /**
   * Membuat pengajuan cuti melahirkan baru
   */
  async create(input: MaternityLeaveRequestInput): Promise<MaternityLeaveRequest> {
    // Validasi sesi kustom sebelum kirim
    const currentUser = authService.getCurrentUser();
    if (!currentUser) {
      console.error('No active custom session found during create maternity leave');
      throw new Error('Sesi Anda telah berakhir. Silakan login kembali.');
    }

    console.log('Creating maternity leave request with payload:', input);

    const { data, error } = await supabase
      .from('account_maternity_leaves')
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
      console.error('Supabase error in create maternity leave:', error);
      throw error;
    }

    // Sinkronisasi ke tabel submissions agar muncul di daftar verifikasi pusat
    try {
      await supabase.from('account_submissions').insert([{
        account_id: input.account_id,
        type: 'Cuti Melahirkan',
        status: 'Pending',
        description: input.description,
        submission_data: {
          start_date: input.start_date,
          end_date: input.end_date,
          maternity_leave_id: data.id
        }
      }]);
    } catch (subError) {
      console.warn('Failed to sync to submissions table:', subError);
    }

    return data;
  },

  /**
   * Negosiasi / Respon Cuti Melahirkan (Admin atau User)
   */
  async negotiate(
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
      .from('account_maternity_leaves')
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
      .from('account_maternity_leaves')
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
      console.error('Error updating maternity leave negotiation:', error);
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
        .eq('type', 'Cuti Melahirkan')
        .contains('submission_data', { maternity_leave_id: id })
        .maybeSingle();

      if (sub) {
        await supabase.from('account_submissions')
          .update({ 
            status: submissionStatus,
            description: reason,
            submission_data: {
              start_date: startDate,
              end_date: endDate,
              maternity_leave_id: id
            }
          })
          .eq('id', sub.id);
      }
    } catch (subError) {
      console.warn('Failed to update central submission status:', subError);
    }
  },

  /**
   * Menghapus pengajuan cuti melahirkan
   */
  async delete(id: string): Promise<void> {
    const { error } = await supabase
      .from('account_maternity_leaves')
      .delete()
      .eq('id', id);
    
    if (error) throw error;
  },

  /**
   * Mendapatkan pengajuan cuti melahirkan dalam rentang tanggal
   */
  async getByRange(accountId: string, startDate: string, endDate: string): Promise<MaternityLeaveRequest[]> {
    const { data, error } = await supabase
      .from('account_maternity_leaves')
      .select('*')
      .eq('account_id', accountId)
      .neq('status', 'rejected')
      .neq('status', 'cancelled')
      .or(`start_date.lte.${endDate},end_date.gte.${startDate}`);
    
    if (error) throw error;
    return data || [];
  }
};
