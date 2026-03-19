import { supabase } from '../lib/supabase';
import { Submission, SubmissionInput, SubmissionStatus } from '../types';

const sanitizePayload = (payload: any) => {
  const sanitized = { ...payload };
  Object.keys(sanitized).forEach(key => {
    if (sanitized[key] === '' || sanitized[key] === undefined) {
      sanitized[key] = null;
    }
  });
  return sanitized;
};

export const submissionService = {
  async getAll() {
    const { data, error } = await supabase
      .from('account_submissions')
      .select(`
        *,
        account:accounts!account_id(full_name, internal_nik)
      `)
      .order('created_at', { ascending: false });
    
    if (error) throw error;
    return data as Submission[];
  },

  async getByType(type: string) {
    const { data, error } = await supabase
      .from('account_submissions')
      .select(`
        *,
        account:accounts!account_id(full_name, internal_nik)
      `)
      .eq('type', type)
      .order('created_at', { ascending: false });
    
    if (error) throw error;
    return data as Submission[];
  },

  async getSubmissionsByRange(startDate: string, endDate: string) {
    const { data, error } = await supabase
      .from('account_submissions')
      .select('*')
      .gte('created_at', `${startDate}T00:00:00Z`)
      .lte('created_at', `${endDate}T23:59:59Z`)
      .order('created_at', { ascending: true });
    
    if (error) throw error;
    return data as Submission[];
  },

  async getByAccountId(accountId: string) {
    const { data, error } = await supabase
      .from('account_submissions')
      .select('*')
      .eq('account_id', accountId)
      .order('created_at', { ascending: false });
    
    if (error) throw error;
    return data as Submission[];
  },

  async create(input: SubmissionInput) {
    const sanitized = sanitizePayload(input);
    const { data, error } = await supabase
      .from('account_submissions')
      .insert([sanitized])
      .select()
      .single();
    
    if (error) throw error;
    return data as Submission;
  },

  async verify(id: string, status: SubmissionStatus, verifierId: string, notes?: string) {
    const { data: submission, error: fetchError } = await supabase
      .from('account_submissions')
      .select('*')
      .eq('id', id)
      .single();

    if (fetchError) throw fetchError;

    const { data, error } = await supabase
      .from('account_submissions')
      .update({
        status,
        verifier_id: verifierId,
        verified_at: new Date().toISOString(),
        verification_notes: notes
      })
      .eq('id', id)
      .select()
      .single();
    
    if (error) throw error;

    // Logic khusus setelah disetujui (Post-Approval Automation)
    if (status === 'Disetujui') {
      if (submission.type === 'Cuti') {
        const { duration_days } = submission.submission_data;
        if (duration_days) {
          // Potong kuota cuti di profil akun
          const { data: account } = await supabase.from('accounts').select('leave_quota').eq('id', submission.account_id).single();
          if (account) {
            await supabase.from('accounts').update({ 
              leave_quota: Math.max(0, account.leave_quota - duration_days) 
            }).eq('id', submission.account_id);
          }
        }
      } else if (submission.type === 'Libur Mandiri') {
        const { leave_request_id } = submission.submission_data;
        if (leave_request_id) {
          // Sinkronisasi status ke tabel libur mandiri
          await supabase.from('account_leave_requests')
            .update({ status: 'approved', updated_at: new Date().toISOString() })
            .eq('id', leave_request_id);
        }
      } else if (submission.type === 'Cuti Tahunan') {
        const { annual_leave_id } = submission.submission_data;
        if (annual_leave_id) {
          // Sinkronisasi status ke tabel cuti tahunan
          const { data: current } = await supabase.from('account_annual_leaves').select('negotiation_data').eq('id', annual_leave_id).single();
          const newHistory = [...(current?.negotiation_data || []), {
            role: 'admin',
            start_date: submission.submission_data.start_date,
            end_date: submission.submission_data.end_date,
            reason: notes || 'Disetujui via Modul Pengajuan',
            timestamp: new Date().toISOString()
          }];
          await supabase.from('account_annual_leaves')
            .update({ 
              status: 'approved', 
              negotiation_data: newHistory,
              updated_at: new Date().toISOString() 
            })
            .eq('id', annual_leave_id);
        }
      } else if (submission.type === 'Izin') {
        const { permission_request_id } = submission.submission_data;
        if (permission_request_id) {
          // Sinkronisasi status ke tabel izin
          const { data: current } = await supabase.from('account_permission_requests').select('negotiation_data').eq('id', permission_request_id).single();
          const newHistory = [...(current?.negotiation_data || []), {
            role: 'admin',
            start_date: submission.submission_data.start_date,
            end_date: submission.submission_data.end_date,
            reason: notes || 'Disetujui via Modul Pengajuan',
            timestamp: new Date().toISOString()
          }];
          await supabase.from('account_permission_requests')
            .update({ 
              status: 'approved', 
              negotiation_data: newHistory,
              updated_at: new Date().toISOString() 
            })
            .eq('id', permission_request_id);
        }
      } else if (submission.type === 'Cuti Melahirkan') {
        const { maternity_leave_id } = submission.submission_data;
        if (maternity_leave_id) {
          // Sinkronisasi status ke tabel cuti melahirkan
          const { data: current } = await supabase.from('account_maternity_leaves').select('negotiation_data').eq('id', maternity_leave_id).single();
          const newHistory = [...(current?.negotiation_data || []), {
            role: 'admin',
            start_date: submission.submission_data.start_date,
            end_date: submission.submission_data.end_date,
            reason: notes || 'Disetujui via Modul Pengajuan',
            timestamp: new Date().toISOString()
          }];
          await supabase.from('account_maternity_leaves')
            .update({ 
              status: 'approved', 
              negotiation_data: newHistory,
              updated_at: new Date().toISOString() 
            })
            .eq('id', maternity_leave_id);
        }
      }
      // Tambahkan logic otomatisasi lain di sini (misal: insert log lembur otomatis)
    } else if (status === 'Ditolak') {
      if (submission.type === 'Libur Mandiri') {
        const { leave_request_id } = submission.submission_data;
        if (leave_request_id) {
          // Sinkronisasi status ke tabel libur mandiri
          await supabase.from('account_leave_requests')
            .update({ status: 'rejected', updated_at: new Date().toISOString() })
            .eq('id', leave_request_id);
        }
      } else if (submission.type === 'Cuti Tahunan') {
        const { annual_leave_id } = submission.submission_data;
        if (annual_leave_id) {
          const { data: current } = await supabase.from('account_annual_leaves').select('negotiation_data').eq('id', annual_leave_id).single();
          const newHistory = [...(current?.negotiation_data || []), {
            role: 'admin',
            start_date: submission.submission_data.start_date,
            end_date: submission.submission_data.end_date,
            reason: notes || 'Ditolak via Modul Pengajuan',
            timestamp: new Date().toISOString()
          }];
          await supabase.from('account_annual_leaves')
            .update({ 
              status: 'rejected', 
              negotiation_data: newHistory,
              updated_at: new Date().toISOString() 
            })
            .eq('id', annual_leave_id);
        }
      } else if (submission.type === 'Izin') {
        const { permission_request_id } = submission.submission_data;
        if (permission_request_id) {
          const { data: current } = await supabase.from('account_permission_requests').select('negotiation_data').eq('id', permission_request_id).single();
          const newHistory = [...(current?.negotiation_data || []), {
            role: 'admin',
            start_date: submission.submission_data.start_date,
            end_date: submission.submission_data.end_date,
            reason: notes || 'Ditolak via Modul Pengajuan',
            timestamp: new Date().toISOString()
          }];
          await supabase.from('account_permission_requests')
            .update({ 
              status: 'rejected', 
              negotiation_data: newHistory,
              updated_at: new Date().toISOString() 
            })
            .eq('id', permission_request_id);
        }
      } else if (submission.type === 'Cuti Melahirkan') {
        const { maternity_leave_id } = submission.submission_data;
        if (maternity_leave_id) {
          const { data: current } = await supabase.from('account_maternity_leaves').select('negotiation_data').eq('id', maternity_leave_id).single();
          const newHistory = [...(current?.negotiation_data || []), {
            role: 'admin',
            start_date: submission.submission_data.start_date,
            end_date: submission.submission_data.end_date,
            reason: notes || 'Ditolak via Modul Pengajuan',
            timestamp: new Date().toISOString()
          }];
          await supabase.from('account_maternity_leaves')
            .update({ 
              status: 'rejected', 
              negotiation_data: newHistory,
              updated_at: new Date().toISOString() 
            })
            .eq('id', maternity_leave_id);
        }
      }
    }

    return data as Submission;
  },

  async delete(id: string) {
    const { error } = await supabase
      .from('account_submissions')
      .delete()
      .eq('id', id);
    
    if (error) throw error;
    return true;
  },

  async getPendingCounts() {
    const { data, error } = await supabase
      .from('account_submissions')
      .select('type')
      .eq('status', 'Pending');
    
    if (error) throw error;
    
    const counts: Record<string, number> = {
      'Libur Mandiri': 0,
      'Lembur': 0,
      'Izin': 0,
      'Cuti Tahunan': 0,
      'Cuti Melahirkan': 0
    };

    data?.forEach(item => {
      if (counts[item.type] !== undefined) {
        counts[item.type]++;
      } else {
        counts[item.type] = (counts[item.type] || 0) + 1;
      }
    });

    return counts;
  }
};