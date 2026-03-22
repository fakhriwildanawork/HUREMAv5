import { supabase } from '../lib/supabase';
import { CareerLogExtended } from '../types';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { accountService } from './accountService';
import { locationService } from './locationService';
import { scheduleService } from './scheduleService';

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
    const accounts = await accountService.getAll();
    const locations = await locationService.getAll();
    const allSchedules = await scheduleService.getAll();

    const workbook = new ExcelJS.Workbook();
    const wsImport = workbook.addWorksheet('Career_Import');
    const wsData = workbook.addWorksheet('Data_Reference');
    wsData.state = 'hidden';

    // Add locations to reference sheet
    wsData.getCell('A1').value = 'Locations';
    locations.forEach((loc, idx) => {
      wsData.getCell(`A${idx + 2}`).value = loc.name;
    });

    // Add schedules to reference sheet
    wsData.getCell('B1').value = 'Schedules';
    const uniqueSchedules = Array.from(new Set(allSchedules.map(s => s.name)));
    uniqueSchedules.push('Fleksibel');
    uniqueSchedules.push('Shift Dinamis');
    uniqueSchedules.forEach((sch, idx) => {
      wsData.getCell(`B${idx + 2}`).value = sch;
    });

    wsImport.addRow(["Harap isi data riwayat karir terbaru karyawan. Baris dengan (*) wajib diisi."]);
    wsImport.addRow(['']); 
    
    const headers = [
      'Account ID (Hidden)', 
      'NIK Internal', 
      'Nama Karyawan', 
      'Nomor SK (*)',
      'Jabatan Baru (*)', 
      'Departemen/Divisi Baru (*)', 
      'Nama Lokasi (*)', 
      'Nama Jadwal (*)', 
      'Tanggal Perubahan (YYYY-MM-DD) (*)', 
      'Catatan / Keterangan'
    ];
    wsImport.addRow(headers);

    const headerRow = wsImport.getRow(3);
    headerRow.font = { bold: true };

    // Mandatory columns: D (Nomor SK), E (Jabatan), F (Dept/Div), G (Nama Lokasi), H (Nama Jadwal), I (Tanggal)
    [4, 5, 6, 7, 8, 9].forEach(colIdx => {
      const cell = headerRow.getCell(colIdx);
      cell.font = { color: { argb: 'FFFF0000' }, bold: true };
    });

    accounts?.forEach(acc => {
      wsImport.addRow([acc.id, acc.internal_nik, acc.full_name, '', '', '', '', '', '', '']);
    });

    const rowCount = wsImport.rowCount;
    for (let i = 4; i <= rowCount; i++) {
      // Location Dropdown (Column G)
      const cellG = wsImport.getCell(`G${i}`);
      cellG.dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: [`Data_Reference!$A$2:$A$${locations.length + 1}`]
      };

      // Schedule Dropdown (Column H)
      const cellH = wsImport.getCell(`H${i}`);
      cellH.dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: [`Data_Reference!$B$2:$B$${uniqueSchedules.length + 1}`]
      };

      // Date Validation (Column I)
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
      col.width = [20, 15, 25, 30, 20, 20, 25, 25, 22, 25][idx];
    });

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), `HUREMA_Career_Template_${new Date().toISOString().split('T')[0]}.xlsx`);
  },

  async processImport(file: File, bulkFiles: Record<string, string> = {}) {
    const locations = await locationService.getAll();
    const allSchedules = await scheduleService.getAll();

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

            const skNumber = row['Nomor SK (*)'] || '';
            let matchedFileId = null;
            if (skNumber) {
              const normalizedNo = String(skNumber).replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
              const match = Object.entries(bulkFiles).find(([fileName]) => {
                const normalizedFileName = fileName.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
                return normalizedFileName === normalizedNo;
              });
              if (match) matchedFileId = match[1];
            }

            // Resolve Location ID
            const locationName = row['Nama Lokasi (*)'] || '';
            const location = locations.find(l => l.name.trim().toLowerCase() === locationName.trim().toLowerCase());
            const locationId = location?.id || null;

            // Resolve Schedule ID & Type
            const scheduleName = row['Nama Jadwal (*)'] || '';
            let scheduleId = null;
            let scheduleType = '';

            if (scheduleName.toLowerCase() === 'fleksibel') {
              scheduleType = 'Fleksibel';
            } else if (scheduleName.toLowerCase() === 'shift dinamis') {
              scheduleType = 'Shift Dinamis';
            } else {
              const sch = allSchedules.find(s => s.name.trim().toLowerCase() === scheduleName.trim().toLowerCase());
              if (sch) {
                scheduleId = sch.id;
                scheduleType = sch.name;
              }
            }

            return {
              account_id: row['Account ID (Hidden)'],
              full_name: row['Nama Karyawan'],
              sk_number: skNumber,
              position: row['Jabatan Baru (*)'],
              grade: row['Departemen/Divisi Baru (*)'],
              location_id: locationId,
              location_name: locationName,
              schedule_id: scheduleId,
              schedule_type: scheduleType,
              change_date: effectiveDate,
              notes: row['Catatan / Keterangan'] || null,
              file_sk_id: matchedFileId,
              isValid: !!(row['Account ID (Hidden)'] && skNumber && row['Jabatan Baru (*)'] && row['Departemen/Divisi Baru (*)'] && locationId && scheduleType && effectiveDate)
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
        schedule_type: item.schedule_type,
        notes: item.notes,
        change_date: item.change_date,
        file_sk_id: item.file_sk_id || null
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
