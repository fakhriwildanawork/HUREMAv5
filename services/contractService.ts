
import { supabase } from '../lib/supabase';
import { AccountContract, AccountContractExtended, AccountContractInput } from '../types';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { accountService } from './accountService';

const sanitizePayload = (payload: any) => {
  const sanitized = { ...payload };
  Object.keys(sanitized).forEach(key => {
    if (sanitized[key] === '' || sanitized[key] === undefined) {
      sanitized[key] = null;
    }
  });
  return sanitized;
};

export const contractService = {
  async getAllGlobal() {
    const { data, error } = await supabase
      .from('account_contracts')
      .select(`
        *,
        account:accounts(full_name, internal_nik, role, access_code, photo_google_id)
      `)
      .order('created_at', { ascending: false });
    
    if (error) throw error;
    // Filter out logs where account access_code contains SPADMIN (case-insensitive)
    return (data as any[]).filter(log => !log.account?.access_code?.toUpperCase().includes('SPADMIN')) as AccountContractExtended[];
  },

  async getByAccountId(accountId: string) {
    const { data, error } = await supabase
      .from('account_contracts')
      .select('*')
      .eq('account_id', accountId)
      .order('start_date', { ascending: false });
    
    if (error) throw error;
    return data as AccountContract[];
  },

  async create(input: AccountContractInput) {
    const sanitized = sanitizePayload(input);
    const { data, error } = await supabase
      .from('account_contracts')
      .insert([sanitized])
      .select();
    
    if (error) throw error;

    // Sinkronisasi: Update profil utama karyawan
    await accountService.update(input.account_id, {
      start_date: input.start_date,
      end_date: input.end_date || null
    });

    return data[0] as AccountContract;
  },

  async update(id: string, input: Partial<AccountContractInput>) {
    const sanitized = sanitizePayload(input);
    const { data, error } = await supabase
      .from('account_contracts')
      .update(sanitized)
      .eq('id', id)
      .select();
    
    if (error) throw error;

    if (input.account_id) {
       await accountService.update(input.account_id, {
         start_date: input.start_date || undefined,
         end_date: input.end_date || null
       });
    }

    return data[0] as AccountContract;
  },

  async delete(id: string) {
    // 1. Ambil ID file
    const { data } = await supabase.from('account_contracts').select('file_id').eq('id', id).single();
    
    // 2. Hapus file dari Drive
    if (data?.file_id) {
      const { googleDriveService } = await import('./googleDriveService');
      await googleDriveService.deleteFile(data.file_id);
    }

    // 3. Hapus dari DB
    const { error } = await supabase
      .from('account_contracts')
      .delete()
      .eq('id', id);
    if (error) throw error;
    return true;
  },

  async bulkDelete(ids: string[]) {
    for (const id of ids) {
      await this.delete(id);
    }
    return true;
  },

  async downloadTemplate() {
    const accounts = await accountService.getAll();
    const workbook = new ExcelJS.Workbook();
    const wsImport = workbook.addWorksheet('Contract_Import');
    
    const headers = [
      'Account ID (Hidden)', 
      'NIK Internal', 
      'Nama Karyawan', 
      'Nomor Kontrak (*)', 
      'Jenis Kontrak (*)', 
      'Tgl Mulai (YYYY-MM-DD) (*)', 
      'Tgl Akhir (YYYY-MM-DD) (*)', 
      'Keterangan'
    ];
    wsImport.addRow(headers);
    
    // Style headers: Bold and Red for mandatory columns
    const headerRow = wsImport.getRow(1);
    headerRow.font = { bold: true };
    
    // Mandatory columns: D (No Kontrak), E (Jenis), F (Mulai), G (Akhir)
    [4, 5, 6, 7].forEach(colIdx => {
      const cell = headerRow.getCell(colIdx);
      cell.font = { color: { argb: 'FFFF0000' }, bold: true };
    });

    accounts.forEach(acc => {
      wsImport.addRow([
        acc.id,
        acc.internal_nik,
        acc.full_name,
        '', '', '', '', ''
      ]);
    });

    const contractTypes = ['PKWT', 'PKWTT', 'Magang', 'Harian', 'Addendum'];
    const maxRow = wsImport.rowCount + 500;
    for (let i = 2; i <= maxRow; i++) {
      wsImport.getCell(`E${i}`).dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: [`"${contractTypes.join(',')}"`]
      };
      
      ['F', 'G'].forEach(col => {
        const cell = wsImport.getCell(`${col}${i}`);
        cell.dataValidation = {
          type: 'date',
          operator: 'greaterThan',
          allowBlank: true,
          formulae: [new Date(1900, 0, 1)]
        };
        cell.numFmt = 'yyyy-mm-dd';
      });
    }

    wsImport.columns.forEach((col, idx) => {
      col.width = [20, 15, 25, 22, 18, 22, 22, 22][idx];
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const dataBlob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(dataBlob, `HUREMA_Contract_Template_${new Date().toISOString().split('T')[0]}.xlsx`);
  },

  async processImport(file: File, bulkFiles: Record<string, string> = {}) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const jsonData = XLSX.utils.sheet_to_json(firstSheet);

          const results = jsonData.map((row: any) => {
            const parseDate = (val: any) => {
              if (typeof val === 'number') {
                const date = new Date((val - 25569) * 86400 * 1000);
                return date.toISOString().split('T')[0];
              }
              return val;
            };

            const startDate = parseDate(row['Tgl Mulai (YYYY-MM-DD) (*)']);
            const endDate = parseDate(row['Tgl Akhir (YYYY-MM-DD) (*)']);
            const contractNumber = row['Nomor Kontrak (*)'] || '';

            // Smart matching logic
            let matchedFileId = null;
            if (contractNumber) {
              const normalizedNo = String(contractNumber).replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
              // Try to find in bulkFiles
              const match = Object.entries(bulkFiles).find(([fileName]) => {
                const normalizedFileName = fileName.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
                return normalizedFileName === normalizedNo;
              });
              if (match) matchedFileId = match[1];
            }

            return {
              account_id: row['Account ID (Hidden)'],
              full_name: row['Nama Karyawan'],
              internal_nik: row['NIK Internal'],
              contract_number: contractNumber,
              contract_type: row['Jenis Kontrak (*)'],
              start_date: startDate,
              end_date: endDate || null,
              notes: row['Keterangan'] || null,
              file_id: matchedFileId,
              isValid: !!(row['Account ID (Hidden)'] && contractNumber && startDate)
            };
          });
          resolve(results);
        } catch (err) { reject(err); }
      };
      reader.readAsArrayBuffer(file);
    });
  },

  async commitImport(data: any[]) {
    const validData = data.filter(d => d.isValid);
    const results = [];
    for (const item of validData) {
      const res = await this.create({
        account_id: item.account_id,
        contract_number: item.contract_number,
        contract_type: item.contract_type,
        start_date: item.start_date,
        end_date: item.end_date,
        notes: item.notes,
        file_id: item.file_id || null
      });
      results.push(res);
    }
    return results;
  }
};
