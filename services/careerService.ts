import { supabase } from '../lib/supabase';
import { CareerLogExtended } from '../types';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { accountService } from './accountService';

export const careerService = {
  async getAllGlobal() {
    const { data, error } = await supabase
      .from('account_career_logs')
      .select(`
        *,
        account:accounts(full_name, internal_nik, role, access_code, photo_google_id)
      `)
      .order('change_date', { ascending: false });
    
    if (error) throw error;
    // Filter out logs where account access_code contains SPADMIN (case-insensitive)
    return (data as any[]).filter(log => !log.account?.access_code?.toUpperCase().includes('SPADMIN')) as CareerLogExtended[];
  },

  async downloadTemplate() {
    // Optimasi: Hanya ambil data minimal untuk mencegah lag saat download
    const { data: accounts, error } = await supabase
      .from('accounts')
      .select('id, internal_nik, full_name')
      .is('end_date', null)
      .not('access_code', 'ilike', '%SPADMIN%');

    if (error) throw error;

    const workbook = new ExcelJS.Workbook();
    const wsImport = workbook.addWorksheet('Career_Import');
    
    wsImport.addRow(["Harap isi data riwayat karir terbaru karyawan. Baris dengan (*) wajib diisi."]);
    wsImport.addRow(['']); 
    
    const headers = [
      'Account ID (Hidden)', 
      'NIK Internal', 
      'Nama Karyawan', 
      'Jabatan Baru (*)', 
      'Grade Baru (*)', 
      'ID Lokasi (*)', 
      'Nama Lokasi (*)', 
      'ID Jadwal (*)', 
      'Tanggal Perubahan (YYYY-MM-DD) (*)', 
      'Catatan / Keterangan', 
      'Link SK G-Drive (Opsional)'
    ];
    wsImport.addRow(headers);

    const headerRow = wsImport.getRow(3);
    headerRow.font = { bold: true };

    accounts?.forEach(acc => {
      wsImport.addRow([acc.id, acc.internal_nik, acc.full_name, '', '', '', '', '', '', '', '']);
    });

    const rowCount = wsImport.rowCount;
    // Optimasi: Loop hanya pada baris yang memiliki data
    for (let i = 4; i <= rowCount; i++) {
      const cellI = wsImport.getCell(`I${i}`);
      cellI.dataValidation = {
        type: 'date',
        operator: 'greaterThan',
        allowBlank: true,
        formulae: [new Date(1900, 0, 1)]
      };
      cellI.numFmt = 'yyyy-mm-dd';
    }

    wsImport.columns.forEach((col, idx) => {
      col.width = [20, 15, 25, 20, 15, 15, 20, 15, 22, 25, 30][idx];
    });

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), `HUREMA_Career_Template_${new Date().toISOString().split('T')[0]}.xlsx`);
  },

  async processImport(file: File) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { range: 2 });

          const results = jsonData.map((row: any) => {
            let effectiveDate = row['Tanggal Perubahan (YYYY-MM-DD) (*)'];
            if (typeof effectiveDate === 'number') {
              effectiveDate = new Date((effectiveDate - 25569) * 86400 * 1000).toISOString().split('T')[0];
            }

            return {
              account_id: row['Account ID (Hidden)'],
              full_name: row['Nama Karyawan'],
              position: row['Jabatan Baru (*)'],
              grade: row['Grade Baru (*)'],
              location_id: row['ID Lokasi (*)'],
              location_name: row['Nama Lokasi (*)'],
              schedule_id: row['ID Jadwal (*)'],
              change_date: effectiveDate,
              notes: row['Catatan / Keterangan'] || null,
              file_sk_link: row['Link SK G-Drive (Opsional)'] || null,
              isValid: !!(row['Account ID (Hidden)'] && row['Jabatan Baru (*)'] && row['Grade Baru (*)'] && row['ID Lokasi (*)'] && row['Nama Lokasi (*)'] && row['ID Jadwal (*)'] && effectiveDate)
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
    for (const item of validData) {
      await accountService.createCareerLog({
        account_id: item.account_id,
        position: item.position,
        grade: item.grade,
        location_id: item.location_id,
        location_name: item.location_name,
        schedule_id: item.schedule_id,
        notes: item.notes,
        change_date: item.change_date,
        file_sk_id: item.file_sk_link ? item.file_sk_link.match(/[-\w]{25,}/)?.[0] : null
      });
    }
  },

  async delete(id: string) {
    return accountService.deleteCareerLog(id);
  },

  async bulkDelete(ids: string[]) {
    for (const id of ids) {
      await this.delete(id);
    }
    return true;
  }
};
