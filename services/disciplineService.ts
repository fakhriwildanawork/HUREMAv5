import { supabase } from '../lib/supabase';
import { WarningLog, WarningLogExtended, WarningLogInput, TerminationLog, TerminationLogExtended, TerminationLogInput } from '../types';
import { accountService } from './accountService';
import { financeService } from './financeService';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

const sanitizePayload = (payload: any) => {
  const sanitized = { ...payload };
  Object.keys(sanitized).forEach(key => {
    if (sanitized[key] === '' || sanitized[key] === undefined) {
      sanitized[key] = null;
    }
  });
  return sanitized;
};

export const disciplineService = {
  // --- Warnings ---
  async getWarningsAll() {
    const { data, error } = await supabase
      .from('account_warning_logs')
      .select('*, account:accounts(full_name, internal_nik, role, access_code, photo_google_id)')
      .order('issue_date', { ascending: false });
    if (error) throw error;
    // Filter out logs where account access_code contains SPADMIN (case-insensitive)
    return (data as any[]).filter(log => !log.account?.access_code?.toUpperCase().includes('SPADMIN')) as WarningLogExtended[];
  },

  async getWarningsByAccountId(accountId: string) {
    const { data, error } = await supabase
      .from('account_warning_logs')
      .select('*')
      .eq('account_id', accountId)
      .order('issue_date', { ascending: false });
    if (error) throw error;
    return data as WarningLog[];
  },

  async createWarning(input: WarningLogInput) {
    const sanitized = sanitizePayload(input);
    const { data, error } = await supabase
      .from('account_warning_logs')
      .insert([sanitized])
      .select();
    if (error) throw error;
    return data[0] as WarningLog;
  },

  async updateWarning(id: string, input: Partial<WarningLogInput>) {
    const sanitized = sanitizePayload(input);
    const { data, error } = await supabase
      .from('account_warning_logs')
      .update(sanitized)
      .eq('id', id)
      .select();
    if (error) throw error;
    return data[0] as WarningLog;
  },

  async updateTermination(id: string, input: Partial<TerminationLogInput>) {
    const sanitized = sanitizePayload(input);
    const { data, error } = await supabase
      .from('account_termination_logs')
      .update(sanitized)
      .eq('id', id)
      .select();
    if (error) throw error;

    // Update end_date di profile akun jika termination_date berubah
    if (input.account_id && input.termination_date) {
      await accountService.update(input.account_id, {
        end_date: input.termination_date
      });
    }

    return data[0] as TerminationLog;
  },

  async deleteWarning(id: string) {
    // 1. Ambil ID file
    const { data } = await supabase.from('account_warning_logs').select('file_id').eq('id', id).single();
    
    // 2. Hapus file dari Drive
    if (data?.file_id) {
      const { googleDriveService } = await import('./googleDriveService');
      await googleDriveService.deleteFile(data.file_id);
    }

    // 3. Hapus dari DB
    const { error } = await supabase.from('account_warning_logs').delete().eq('id', id);
    if (error) throw error;
    return true;
  },

  async bulkDeleteWarnings(ids: string[]) {
    for (const id of ids) {
      await this.deleteWarning(id);
    }
    return true;
  },

  // --- Terminations ---
  async getTerminationsAll() {
    const { data, error } = await supabase
      .from('account_termination_logs')
      .select('*, account:accounts(full_name, internal_nik, role, access_code, photo_google_id)')
      .order('termination_date', { ascending: false });
    if (error) throw error;
    // Filter out logs where account access_code contains SPADMIN (case-insensitive)
    return (data as any[]).filter(log => !log.account?.access_code?.toUpperCase().includes('SPADMIN')) as TerminationLogExtended[];
  },

  async getTerminationByAccountId(accountId: string) {
    const { data, error } = await supabase
      .from('account_termination_logs')
      .select('*')
      .eq('account_id', accountId)
      .maybeSingle();
    if (error) throw error;
    return data as TerminationLog | null;
  },

  async createTermination(input: TerminationLogInput) {
    const sanitized = sanitizePayload(input);
    const { data, error } = await supabase
      .from('account_termination_logs')
      .insert([sanitized])
      .select();
    if (error) throw error;

    // Otomatis update end_date di profile akun
    await accountService.update(input.account_id, {
      end_date: input.termination_date
    });

    // Create Compensation Record if amount > 0
    if (input.severance_amount > 0 || input.penalty_amount > 0) {
      await financeService.createCompensation({
        account_id: input.account_id,
        termination_type: input.termination_type,
        termination_date: input.termination_date,
        amount: input.termination_type === 'Pemecatan' ? input.severance_amount : input.penalty_amount,
        type: input.termination_type === 'Pemecatan' ? 'Severance' : 'Penalty',
        reason: input.reason
      });
    }

    return data[0] as TerminationLog;
  },

  async deleteTermination(id: string, accountId: string) {
    // 1. Ambil ID file
    const { data } = await supabase.from('account_termination_logs').select('file_id').eq('id', id).single();
    
    // 2. Hapus file dari Drive
    if (data?.file_id) {
      const { googleDriveService } = await import('./googleDriveService');
      await googleDriveService.deleteFile(data.file_id);
    }

    // 3. Hapus dari DB
    const { error } = await supabase.from('account_termination_logs').delete().eq('id', id);
    if (error) throw error;

    // Aktifkan kembali akun dengan me-nullkan end_date
    await accountService.update(accountId, {
      end_date: null
    });
    return true;
  },

  async bulkDeleteTerminations(items: { id: string, account_id: string }[]) {
    for (const item of items) {
      await this.deleteTermination(item.id, item.account_id);
    }
    return true;
  },

  // --- Import / Export ---
  async downloadWarningTemplate() {
    // Optimasi: Hanya ambil data yang diperlukan saja untuk mencegah lag
    const { data: accounts, error } = await supabase
      .from('accounts')
      .select('id, internal_nik, full_name')
      .is('end_date', null)
      .not('access_code', 'ilike', '%SPADMIN%');

    if (error) throw error;

    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Warning_Import');
    
    ws.addRow(['Account ID (Hidden)', 'NIK Internal', 'Nama Karyawan', 'Jenis Peringatan (*)', 'Alasan (*)', 'Tanggal (YYYY-MM-DD) (*)', 'Link Surat G-Drive']);
    ws.getRow(1).font = { bold: true };

    accounts?.forEach(acc => {
      ws.addRow([acc.id, acc.internal_nik, acc.full_name, '', '', '', '']);
    });

    const types = ['Teguran', 'SP1', 'SP2', 'SP3'];
    const rowCount = ws.rowCount;
    // Optimasi: Loop hanya pada baris yang memiliki data
    for (let i = 2; i <= rowCount; i++) {
      ws.getCell(`D${i}`).dataValidation = { 
        type: 'list', 
        allowBlank: true, 
        formulae: [`"${types.join(',')}"`] 
      };
      const dateCell = ws.getCell(`F${i}`);
      dateCell.dataValidation = { 
        type: 'date', 
        operator: 'greaterThan', 
        allowBlank: true, 
        formulae: [new Date(1900, 0, 1)] 
      };
      dateCell.numFmt = 'yyyy-mm-dd';
    }

    ws.columns.forEach((col, idx) => { col.width = [20, 15, 25, 20, 30, 22, 30][idx]; });
    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), `HUREMA_Warning_Template_${new Date().toISOString().split('T')[0]}.xlsx`);
  },

  async processWarningImport(file: File) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
          const results = jsonData.map((row: any) => {
            const parseDate = (val: any) => {
              if (typeof val === 'number') return new Date((val - 25569) * 86400 * 1000).toISOString().split('T')[0];
              return val;
            };
            const issueDate = parseDate(row['Tanggal (YYYY-MM-DD) (*)']);
            return {
              account_id: row['Account ID (Hidden)'],
              full_name: row['Nama Karyawan'],
              warning_type: row['Jenis Peringatan (*)'],
              reason: row['Alasan (*)'],
              issue_date: issueDate,
              file_link: row['Link Surat G-Drive'],
              isValid: !!(row['Account ID (Hidden)'] && row['Jenis Peringatan (*)'] && issueDate)
            };
          });
          resolve(results);
        } catch (err) { reject(err); }
      };
      reader.readAsArrayBuffer(file);
    });
  },

  async commitWarningImport(data: any[]) {
    const validData = data.filter(d => d.isValid);
    if (validData.length === 0) return;

    // Optimasi: Gunakan Bulk Insert untuk menghilangkan lag saat commit
    const payload = validData.map(item => ({
      account_id: item.account_id,
      warning_type: item.warning_type,
      reason: item.reason,
      issue_date: item.issue_date,
      file_id: item.file_link ? item.file_link.match(/[-\w]{25,}/)?.[0] : null
    }));

    const { error } = await supabase.from('account_warning_logs').insert(payload);
    if (error) throw error;
  },

  // --- Termination Import Logic ---
  async downloadTerminationTemplate() {
    const { data: accounts, error } = await supabase
      .from('accounts')
      .select('id, internal_nik, full_name')
      .is('end_date', null)
      .not('access_code', 'ilike', '%SPADMIN%');

    if (error) throw error;

    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Termination_Import');
    
    ws.addRow(['Account ID (Hidden)', 'NIK Internal', 'Nama Karyawan', 'Tipe Exit (*)', 'Tanggal Exit (YYYY-MM-DD) (*)', 'Alasan (*)', 'Uang Pesangon (PHK)', 'Biaya Penalti (Resign)', 'Link Dokumen G-Drive']);
    ws.getRow(1).font = { bold: true };

    accounts?.forEach(acc => {
      ws.addRow([acc.id, acc.internal_nik, acc.full_name, '', '', '', 0, 0, '']);
    });

    const types = ['Resign', 'Pemecatan'];
    const rowCount = ws.rowCount;
    for (let i = 2; i <= rowCount; i++) {
      ws.getCell(`D${i}`).dataValidation = { 
        type: 'list', 
        allowBlank: true, 
        formulae: [`"${types.join(',')}"`] 
      };
      const dateCell = ws.getCell(`E${i}`);
      dateCell.dataValidation = { 
        type: 'date', 
        operator: 'greaterThan', 
        allowBlank: true, 
        formulae: [new Date(1900, 0, 1)] 
      };
      dateCell.numFmt = 'yyyy-mm-dd';
    }

    ws.columns.forEach((col, idx) => { col.width = [20, 15, 25, 15, 22, 30, 20, 20, 30][idx]; });
    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), `HUREMA_Termination_Template_${new Date().toISOString().split('T')[0]}.xlsx`);
  },

  async processTerminationImport(file: File) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
          const results = jsonData.map((row: any) => {
            const parseDate = (val: any) => {
              if (typeof val === 'number') return new Date((val - 25569) * 86400 * 1000).toISOString().split('T')[0];
              return val;
            };
            const termDate = parseDate(row['Tanggal Exit (YYYY-MM-DD) (*)']);
            return {
              account_id: row['Account ID (Hidden)'],
              full_name: row['Nama Karyawan'],
              termination_type: row['Tipe Exit (*)'],
              termination_date: termDate,
              reason: row['Alasan (*)'],
              severance_amount: row['Uang Pesangon (PHK)'] || 0,
              penalty_amount: row['Biaya Penalti (Resign)'] || 0,
              file_link: row['Link Dokumen G-Drive'],
              isValid: !!(row['Account ID (Hidden)'] && row['Tipe Exit (*)'] && termDate)
            };
          });
          resolve(results);
        } catch (err) { reject(err); }
      };
      reader.readAsArrayBuffer(file);
    });
  },

  async commitTerminationImport(data: any[]) {
    const validData = data.filter(d => d.isValid);
    if (validData.length === 0) return;

    // 1. Bulk Insert Termination Logs
    const logPayload = validData.map(item => ({
      account_id: item.account_id,
      termination_type: item.termination_type,
      termination_date: item.termination_date,
      reason: item.reason,
      severance_amount: Number(item.severance_amount) || 0,
      penalty_amount: Number(item.penalty_amount) || 0,
      file_id: item.file_link ? item.file_link.match(/[-\w]{25,}/)?.[0] : null
    }));

    const { error: logError } = await supabase.from('account_termination_logs').insert(logPayload);
    if (logError) throw logError;

    // 2. Update Accounts end_date (Sequential for safety, but in a loop)
    for (const item of validData) {
      await accountService.update(item.account_id, {
        end_date: item.termination_date
      });

      // Create Compensation Record if amount > 0
      if (item.severance_amount > 0 || item.penalty_amount > 0) {
        await financeService.createCompensation({
          account_id: item.account_id,
          termination_type: item.termination_type,
          termination_date: item.termination_date,
          amount: item.termination_type === 'Pemecatan' ? item.severance_amount : item.penalty_amount,
          type: item.termination_type === 'Pemecatan' ? 'Severance' : 'Penalty',
          reason: item.reason
        });
      }
    }
  }
};