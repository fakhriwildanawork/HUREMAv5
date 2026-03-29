
import { supabase } from '../lib/supabase';
import { Account, AccountInput, CareerLog, AccountInput as AccountInputType, CareerLogInput, HealthLog, HealthLogInput } from '../types';
import { locationService } from './locationService';
import { scheduleService } from './scheduleService';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';

const VALID_OPTIONS = {
  gender: ['Laki-laki', 'Perempuan'],
  religion: ['Islam', 'Kristen', 'Katolik', 'Hindu', 'Budha', 'Konghucu', 'Kepercayaan Lain'],
  marital_status: ['Belum Menikah', 'Menikah', 'Cerai Hidup', 'Cerai Mati'],
  employee_type: ['Tetap', 'Kontrak', 'Harian', 'Magang'],
  last_education: ['Tidak Sekolah', 'SD', 'SMP/Setara', 'SMA/Setara', 'Diploma 1-4', 'Sarjana', 'Profesi', 'Master', 'Doktor'],
  yes_no: ['Ya', 'Tidak']
};

const normalizeOption = (val: string, options: string[]) => {
  const trimmed = String(val || '').trim();
  if (!trimmed) return '';
  const found = options.find(opt => opt.toLowerCase() === trimmed.toLowerCase());
  return found || trimmed;
};

/**
 * Fungsi pembantu untuk membersihkan data sebelum dikirim ke Supabase.
 * Mengubah string kosong ('') menjadi null agar tidak error saat masuk ke kolom UUID atau DATE.
 */
const sanitizePayload = (payload: any) => {
  const sanitized: any = {};
  const excludedKeys = ['errorMsg', 'isValid', 'grad_date'];
  
  Object.keys(payload).forEach(key => {
    if (excludedKeys.includes(key)) return;
    if (payload[key] === undefined) return;
    sanitized[key] = payload[key] === '' ? null : payload[key];
  });
  return sanitized;
};

export const accountService = {
  async getAll() {
    const { data, error } = await supabase
      .from('accounts')
      .select(`
        *,
        location:locations(name)
      `)
      .not('access_code', 'ilike', '%SPADMIN%')
      .order('created_at', { ascending: false });
    
    if (error) {
      console.error("SUPABASE_GET_ALL_ERROR:", error.message);
      throw error;
    }
    return data;
  },

  async getById(id: string) {
    const { data, error } = await supabase
      .from('accounts')
      .select(`
        *,
        location:locations(*),
        schedule:schedules!schedule_id(*, rules:schedule_rules(*))
      `)
      .eq('id', id)
      .maybeSingle(); // Menggunakan maybeSingle() agar tidak error jika record tidak ditemukan
    
    if (error) {
      console.error("SUPABASE_GET_BY_ID_ERROR:", error.message);
      throw error;
    }
    return data;
  },

  async getDistinctAttributes() {
    // Mengambil data unik Jabatan & Golongan dari kedua tabel untuk memastikan dropdown lengkap
    const [accRes, logRes] = await Promise.all([
      supabase.from('accounts').select('position, grade'),
      supabase.from('account_career_logs').select('position, grade')
    ]);

    const allPositions = [
      ...(accRes.data?.map(p => p.position) || []),
      ...(logRes.data?.map(p => p.position) || [])
    ];
    
    const allGrades = [
      ...(accRes.data?.map(g => g.grade) || []),
      ...(logRes.data?.map(g => g.grade) || [])
    ];

    const uniquePositions = Array.from(new Set(allPositions.filter(Boolean))).sort();
    const uniqueGrades = Array.from(new Set(allGrades.filter(Boolean))).sort();
    
    return { positions: uniquePositions, grades: uniqueGrades };
  },

  async getCareerLogs(accountId: string) {
    const { data, error } = await supabase
      .from('account_career_logs')
      .select('*')
      .eq('account_id', accountId)
      .order('change_date', { ascending: false });
    
    if (error) throw error;
    return data as CareerLog[];
  },

  async getHealthLogs(accountId: string) {
    const { data, error } = await supabase
      .from('account_health_logs')
      .select('*')
      .eq('account_id', accountId)
      .order('change_date', { ascending: false });
    
    if (error) throw error;
    return data as HealthLog[];
  },

  async create(account: AccountInput & { file_sk_id?: string, file_mcu_id?: string, contract_initial?: any }) {
    const { file_sk_id, file_mcu_id, contract_initial, ...rest } = account;
    
    const sanitizedAccount = sanitizePayload(rest);
    
    // 1. Insert ke tabel accounts
    const { data, error } = await supabase
      .from('accounts')
      .insert([sanitizedAccount])
      .select();
    
    if (error) {
      console.error("SUPABASE_CREATE_ERROR:", error.message);
      throw error;
    }
    
    const newAccount = data[0] as Account;

    // 2. Otomatis buat log karier awal
    const { data: locData } = await supabase
      .from('locations')
      .select('name')
      .eq('id', newAccount.location_id)
      .single();

    await supabase.from('account_career_logs').insert([{
      account_id: newAccount.id,
      position: newAccount.position,
      grade: newAccount.grade,
      location_id: newAccount.location_id,
      location_name: locData?.name || '-',
      schedule_id: newAccount.schedule_id,
      file_sk_id: file_sk_id || null,
      notes: 'Initial Career Record'
    }]);

    // 3. Otomatis buat log kesehatan awal - DIHAPUS sesuai permintaan user
    /*
    await supabase.from('account_health_logs').insert([{
      account_id: newAccount.id,
      file_mcu_id: file_mcu_id || null,
      notes: 'Initial Health Record'
    }]);
    */

    // 4. Otomatis buat kontrak awal jika disediakan
    if (contract_initial && contract_initial.contract_number) {
      await supabase.from('account_contracts').insert([{
        account_id: newAccount.id,
        contract_number: contract_initial.contract_number,
        contract_type: contract_initial.contract_type || (account.employee_type === 'Kontrak' ? 'PKWT' : account.employee_type),
        start_date: contract_initial.start_date || account.start_date,
        end_date: contract_initial.end_date || account.end_date,
        file_id: contract_initial.file_id || null,
        notes: 'Initial Contract Record'
      }]);
    }

    return newAccount;
  },

  async update(id: string, account: Partial<AccountInput>) {
    // Pastikan field tambahan untuk log awal tidak ikut dikirim ke tabel accounts
    const { file_sk_id, file_mcu_id, contract_initial, ...rest } = account as any;
    
    const sanitizedAccount = sanitizePayload(rest);
    
    const { data, error } = await supabase
      .from('accounts')
      .update(sanitizedAccount)
      .eq('id', id)
      .select(`
        *,
        location:locations(name)
      `);
    
    if (error) {
      console.error("SUPABASE_UPDATE_ERROR:", error.message);
      throw error;
    }
    return data[0] as Account;
  },

  async delete(id: string) {
    // 1. Ambil data akun untuk mendapatkan ID file Drive
    const { data: account } = await supabase
      .from('accounts')
      .select('photo_google_id, ktp_google_id, diploma_google_id')
      .eq('id', id)
      .single();

    // 2. Ambil semua log untuk mendapatkan ID file Drive mereka
    const [careerLogs, healthLogs, contracts, warnings, terminations] = await Promise.all([
      supabase.from('account_career_logs').select('file_sk_id').eq('account_id', id),
      supabase.from('account_health_logs').select('file_mcu_id').eq('account_id', id),
      supabase.from('account_contracts').select('file_id').eq('account_id', id),
      supabase.from('account_warning_logs').select('file_id').eq('account_id', id),
      supabase.from('account_termination_logs').select('file_id').eq('account_id', id)
    ]);

    // 3. Kumpulkan semua file ID unik
    const fileIds = new Set<string>();
    if (account?.photo_google_id) fileIds.add(account.photo_google_id);
    if (account?.ktp_google_id) fileIds.add(account.ktp_google_id);
    if (account?.diploma_google_id) fileIds.add(account.diploma_google_id);
    
    careerLogs.data?.forEach(log => log.file_sk_id && fileIds.add(log.file_sk_id));
    healthLogs.data?.forEach(log => log.file_mcu_id && fileIds.add(log.file_mcu_id));
    contracts.data?.forEach(log => log.file_id && fileIds.add(log.file_id));
    warnings.data?.forEach(log => log.file_id && fileIds.add(log.file_id));
    terminations.data?.forEach(log => log.file_id && fileIds.add(log.file_id));

    // 4. Hapus file dari Google Drive secara paralel (dengan sedikit delay/batching jika banyak)
    const { googleDriveService } = await import('./googleDriveService');
    for (const fileId of fileIds) {
      try {
        await googleDriveService.deleteFile(fileId);
      } catch (e) {
        console.error(`Gagal menghapus file ${fileId} saat hapus akun:`, e);
      }
    }

    // 5. Hapus dari database (Supabase akan handle cascade jika diatur, tapi kita pastikan)
    const { error } = await supabase
      .from('accounts')
      .delete()
      .eq('id', id);
    
    if (error) {
      console.error("SUPABASE_DELETE_ERROR:", error.message);
      throw error;
    }
    return true;
  },

  async downloadTemplate() {
    try {
      // Fetch data for reference sheets
      const [locations, schedules] = await Promise.all([
        locationService.getAll(),
        scheduleService.getAll()
      ]);

      const workbook = new ExcelJS.Workbook();
      const templateSheet = workbook.addWorksheet('Template');
      const refSheet = workbook.addWorksheet('Lists');
      refSheet.state = 'hidden';

      const headers = [
        'Nama Lengkap (*)', 'NIK KTP (*)', 'Gender (*)', 'Agama (*)', 'Tgl Lahir (YYYY-MM-DD) (*)', 
        'Alamat (*)', 'No Telepon (*)', 'Email (*)', 'Status Nikah (*)', 'Tanggungan', 
        'NIK Internal (*)', 'Jabatan (*)', 'Departemen/Divisi (*)', 'Lokasi Penempatan (*)', 
        'Nomor Kontrak (*)', 'Jenis Kontrak (*)', 'Mulai Kontrak (YYYY-MM-DD) (*)', 'Akhir Kontrak (YYYY-MM-DD)',
        'Pendidikan Terakhir (*)', 'Jurusan', 'Tgl Lulus (YYYY-MM-DD)',
        'Nama Kontak Darurat', 'Hubungan Kontak Darurat', 'No HP Kontak Darurat',
        'Pilih Jadwal Kerja (*)', 'Jatah Cuti Tahunan (*)', 'Jatah Cuti Melahirkan', 
        'Akumulasi Cuti (Ya/Tidak) (*)', 'Maksimal Carry-over (*)', 'Jatah Carry-over Saat Ini (*)',
        'Batasi Check-in Datang (Ya/Tidak) (*)', 'Batasi Check-out Pulang (Ya/Tidak) (*)', 
        'Batasi Check-in Lembur (Ya/Tidak) (*)', 'Batasi Check-out Lembur (Ya/Tidak) (*)',
        'Kode Akses (*)', 'Password (*)'
      ];

      templateSheet.addRow(headers);

      // Add Description Row (Row 2)
      const descriptionRow = [
        'Wajib diisi', 'Wajib diisi', 'Pilih dari daftar', 'Pilih dari daftar', 'Format: YYYY-MM-DD',
        'Wajib diisi', 'Wajib diisi', 'Wajib diisi', 'Pilih dari daftar', 'Opsional (angka)',
        'Wajib diisi', 'Wajib diisi', 'Wajib diisi', 'Pilih dari daftar',
        'Wajib diisi', 'Pilih dari daftar', 'Format: YYYY-MM-DD', 'Kosongkan jika PKWTT',
        'Pilih dari daftar', 'Opsional', 'Format: YYYY-MM-DD (Opsional)',
        'Opsional', 'Opsional', 'Opsional',
        'Pilih dari daftar', 'Wajib diisi (angka)', 'Khusus Perempuan (angka)',
        'Pilih dari daftar', 'Wajib diisi (angka)', 'Wajib diisi (angka)',
        'Pilih dari daftar', 'Pilih dari daftar',
        'Pilih dari daftar', 'Pilih dari daftar',
        'Wajib diisi', 'Wajib diisi'
      ];
      templateSheet.addRow(descriptionRow);

      // Add Example Row (Row 3)
      const exampleRow = [
        'Contoh Nama', '1234567890123456', 'Laki-laki', 'Islam', '1990-01-01',
        'Jl. Contoh No. 123', '08123456789', 'contoh@email.com', 'Belum Menikah', '0',
        'EMP001', 'Staff', 'Operasional', locations[0]?.name || 'Pusat',
        'KTR/001/2024', 'PKWT', '2024-01-01', '2025-01-01',
        'Sarjana', 'Teknik Informatika', '2012-01-01',
        'Nama Kontak', 'Orang Tua', '08123456789',
        'Fleksibel', '12', '0',
        'Tidak', '0', '0',
        'Ya', 'Ya',
        'Ya', 'Ya',
        'USER001', 'pass123'
      ];
      templateSheet.addRow(exampleRow);
      
      // Style headers
      const headerRow = templateSheet.getRow(1);
      headerRow.eachCell((cell) => {
        const isMandatory = cell.value?.toString().includes('(*)');
        cell.font = { 
          bold: true, 
          color: { argb: isMandatory ? 'FFFF0000' : 'FF000000' } 
        };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFE0E0E0' }
        };
      });

      // Style description row
      const descRow = templateSheet.getRow(2);
      descRow.font = { italic: true, size: 10 };
      descRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFF9C4' } // Light yellow
      };

      // Prepare lists for dropdowns
      const genderList = ['Laki-laki', 'Perempuan'];
      const religionList = ['Islam', 'Kristen', 'Katolik', 'Hindu', 'Budha', 'Konghucu', 'Kepercayaan Lain'];
      const maritalList = ['Belum Menikah', 'Menikah', 'Cerai Hidup', 'Cerai Mati'];
      const empTypeList = ['Tetap', 'Kontrak', 'Harian', 'Magang'];
      const contractTypeList = ['PKWT', 'PKWTT', 'Magang', 'Harian'];
      const educationList = ['Tidak Sekolah', 'SD', 'SMP/Setara', 'SMA/Setara', 'Diploma 1-4', 'Sarjana', 'Profesi', 'Master', 'Doktor'];
      const yesNoList = ['Ya', 'Tidak'];
      const locList = locations.map(l => l.name);
      const schList = ['Fleksibel', 'Shift Dinamis', ...schedules.map(s => s.name)];

      // Write lists to hidden sheet
      refSheet.getColumn(1).values = genderList;
      refSheet.getColumn(2).values = religionList;
      refSheet.getColumn(3).values = maritalList;
      refSheet.getColumn(4).values = empTypeList;
      refSheet.getColumn(5).values = yesNoList;
      refSheet.getColumn(6).values = locList;
      refSheet.getColumn(7).values = schList;
      refSheet.getColumn(8).values = educationList;
      refSheet.getColumn(9).values = contractTypeList;

      // Apply Data Validations (Dropdowns)
      for (let i = 4; i <= 203; i++) {
        templateSheet.getCell(`C${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$A$1:$A$${genderList.length}`] };
        templateSheet.getCell(`D${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$B$1:$B$${religionList.length}`] };
        templateSheet.getCell(`I${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$C$1:$C$${maritalList.length}`] };
        templateSheet.getCell(`N${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$F$1:$F$${locList.length}`] };
        // Jenis Kontrak (Col 16 - P)
        templateSheet.getCell(`P${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$I$1:$I$${contractTypeList.length}`] };
        // Pendidikan Terakhir (Col 19 - S)
        templateSheet.getCell(`S${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$H$1:$H$${educationList.length}`] };
        // Pilih Jadwal Kerja (Col 25 - Y)
        templateSheet.getCell(`Y${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$G$1:$G$${schList.length}`] };
        // Akumulasi Cuti (Col 28 - AB)
        templateSheet.getCell(`AB${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$E$1:$E$2`] };
        // Radius Limits (Col 31-34 - AE-AH)
        templateSheet.getCell(`AE${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$E$1:$E$2`] };
        templateSheet.getCell(`AF${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$E$1:$E$2`] };
        templateSheet.getCell(`AG${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$E$1:$E$2`] };
        templateSheet.getCell(`AH${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$E$1:$E$2`] };
      }

      templateSheet.columns.forEach(column => {
        column.width = 25;
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `HUREMA_Account_Template_${new Date().toISOString().split('T')[0]}.xlsx`);
    } catch (error) {
      console.error(error);
      throw error;
    }
  }
,

  async processImport(file: File, bulkFiles: Record<string, string> = {}) {
    // Fetch locations and schedules for validation
    const [locations, schedules] = await Promise.all([
      locationService.getAll(),
      scheduleService.getAll()
    ]);

    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);

          const results = jsonData
            .filter((row: any) => {
              const name = String(row['Nama Lengkap (*)'] || '').trim();
              return name !== '' && name !== 'Contoh Nama' && name !== 'Wajib diisi';
            })
            .map((row: any) => {
              const getVal = (key: string) => row[key] || '';
              
              const forceString = (val: any) => {
                if (val === undefined || val === null) return '';
                // Handle scientific notation for large numbers (like NIK)
                if (typeof val === 'number') {
                  return val.toLocaleString('fullwide', { useGrouping: false });
                }
                return String(val).trim();
              };

              const formatExcelDate = (val: any) => {
                if (!val) return null;
                if (typeof val === 'number') {
                  // Excel date serial to YYYY-MM-DD
                  const date = new Date((val - 25569) * 86400 * 1000);
                  if (isNaN(date.getTime())) return null;
                  return date.toISOString().split('T')[0];
                }
                const str = String(val).trim();
                if (!str) return null;
                // Basic check for YYYY-MM-DD
                if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
                // Try parsing other formats if possible, but YYYY-MM-DD is preferred
                const parsed = new Date(str);
                if (!isNaN(parsed.getTime())) return parsed.toISOString().split('T')[0];
                return str; // Return as is if can't parse, let DB handle or error
              };

              const internalNik = forceString(getVal('NIK Internal (*)'));
              const fileMatches: any = {
                photo_google_id: null,
                ktp_google_id: null,
                diploma_google_id: null,
                file_sk_id: null,
                contract_file_id: null
              };

              if (internalNik) {
                const normalizedNik = internalNik.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
                Object.entries(bulkFiles).forEach(([fileName, fileId]) => {
                  const normalizedFileName = fileName.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
                  if (normalizedFileName.includes(normalizedNik)) {
                    if (normalizedFileName.includes('photo')) fileMatches.photo_google_id = fileId;
                    else if (normalizedFileName.includes('ktp')) fileMatches.ktp_google_id = fileId;
                    else if (normalizedFileName.includes('sk')) fileMatches.file_sk_id = fileId;
                    else if (normalizedFileName.includes('contract')) fileMatches.contract_file_id = fileId;
                    else if (normalizedFileName.includes('diploma')) fileMatches.diploma_google_id = fileId;
                    else if (normalizedFileName === normalizedNik) fileMatches.photo_google_id = fileId;
                  }
                });
              }

              const requiredFields = [
                'Nama Lengkap (*)', 'NIK KTP (*)', 'Gender (*)', 'Agama (*)', 'Tgl Lahir (YYYY-MM-DD) (*)',
                'Alamat (*)', 'No Telepon (*)', 'Email (*)', 'Status Nikah (*)',
                'NIK Internal (*)', 'Jabatan (*)', 'Departemen/Divisi (*)', 'Lokasi Penempatan (*)',
                'Nomor Kontrak (*)', 'Jenis Kontrak (*)', 'Mulai Kontrak (YYYY-MM-DD) (*)',
                'Pendidikan Terakhir (*)',
                'Pilih Jadwal Kerja (*)', 'Jatah Cuti Tahunan (*)',
                'Akumulasi Cuti (Ya/Tidak) (*)', 'Maksimal Carry-over (*)', 'Jatah Carry-over Saat Ini (*)',
                'Batasi Check-in Datang (Ya/Tidak) (*)', 'Batasi Check-out Pulang (Ya/Tidak) (*)',
                'Batasi Check-in Lembur (Ya/Tidak) (*)', 'Batasi Check-out Lembur (Ya/Tidak) (*)',
                'Kode Akses (*)', 'Password (*)'
              ];

              // Pre-validation for Location and Schedule
              let errorMsg = '';
              
              // Check for missing required fields
              const missingFields = requiredFields.filter(field => {
                const val = row[field];
                return val === undefined || val === null || String(val).trim() === '';
              });

              if (missingFields.length > 0) {
                // Clean up field names for display (remove (*) and simplify)
                const cleanNames = missingFields.map(f => f.replace(' (*)', '').replace(' (Ya/Tidak)', '').replace(' (YYYY-MM-DD)', ''));
                errorMsg = `Kolom wajib belum lengkap: [${cleanNames.join(', ')}]`;
              }
              
              // Normalize Dropdown Values
              const gender = normalizeOption(getVal('Gender (*)'), VALID_OPTIONS.gender);
              const religion = normalizeOption(getVal('Agama (*)'), VALID_OPTIONS.religion);
              const maritalStatus = normalizeOption(getVal('Status Nikah (*)'), VALID_OPTIONS.marital_status);
              const contractType = normalizeOption(getVal('Jenis Kontrak (*)'), ['PKWT', 'PKWTT', 'Magang', 'Harian', 'Addendum']);
              const lastEducation = normalizeOption(getVal('Pendidikan Terakhir (*)'), VALID_OPTIONS.last_education);
              const isLeaveAccumulatedStr = normalizeOption(getVal('Akumulasi Cuti (Ya/Tidak) (*)'), VALID_OPTIONS.yes_no);
              
              // Mapping Jenis Kontrak ke Jenis Karyawan
              let employeeType = 'Kontrak';
              if (contractType === 'PKWTT') employeeType = 'Tetap';
              else if (contractType === 'Magang') employeeType = 'Magang';
              else if (contractType === 'Harian') employeeType = 'Harian';
              else if (contractType === 'PKWT') employeeType = 'Kontrak';

              const limitCheckin = normalizeOption(getVal('Batasi Check-in Datang (Ya/Tidak) (*)'), VALID_OPTIONS.yes_no);
              const limitCheckout = normalizeOption(getVal('Batasi Check-out Pulang (Ya/Tidak) (*)'), VALID_OPTIONS.yes_no);
              const limitOtIn = normalizeOption(getVal('Batasi Check-in Lembur (Ya/Tidak) (*)'), VALID_OPTIONS.yes_no);
              const limitOtOut = normalizeOption(getVal('Batasi Check-out Lembur (Ya/Tidak) (*)'), VALID_OPTIONS.yes_no);

              // Validate Dropdown Options
              if (gender && !VALID_OPTIONS.gender.includes(gender)) errorMsg = `Gender '${gender}' tidak valid.`;
              else if (religion && !VALID_OPTIONS.religion.includes(religion)) errorMsg = `Agama '${religion}' tidak valid.`;
              else if (maritalStatus && !VALID_OPTIONS.marital_status.includes(maritalStatus)) errorMsg = `Status Nikah '${maritalStatus}' tidak valid.`;
              else if (employeeType && !VALID_OPTIONS.employee_type.includes(employeeType)) errorMsg = `Jenis Karyawan '${employeeType}' tidak valid.`;
              else if (lastEducation && !VALID_OPTIONS.last_education.includes(lastEducation)) errorMsg = `Pendidikan '${lastEducation}' tidak valid.`;
              else if (isLeaveAccumulatedStr && !VALID_OPTIONS.yes_no.includes(isLeaveAccumulatedStr)) errorMsg = `Opsi Akumulasi Cuti tidak valid.`;

              if (errorMsg) {
                // Skip further logic if basic dropdown is invalid
              } else {
                const locationName = getVal('Lokasi Penempatan (*)');
                const scheduleName = getVal('Pilih Jadwal Kerja (*)');
                
                const location = locations.find(l => l.name === locationName);
                if (!location) {
                  errorMsg = `Lokasi '${locationName}' tidak ditemukan.`;
                } else {
                  if (scheduleName === 'Fleksibel') {
                    // Fleksibel is always valid if location exists
                  } else if (scheduleName === 'Shift Dinamis') {
                    // Check if location has at least one shift schedule (type 2)
                    const hasShift = schedules.some(s => s.location_ids.includes(location.id) && s.type === 2);
                    if (!hasShift) {
                      errorMsg = `Lokasi '${locationName}' tidak mendukung Shift Dinamis (tidak ada jadwal shift).`;
                    }
                  } else {
                    // Regular schedule: must exist in this location and be type 1 or 2
                    const schedule = schedules.find(s => 
                      s.name === scheduleName && 
                      s.location_ids.includes(location.id) && 
                      (s.type === 1 || s.type === 2)
                    );
                    if (!schedule) {
                      errorMsg = `Jadwal '${scheduleName}' tidak tersedia di Lokasi '${locationName}'.`;
                    }
                  }
                }

                // Dependency Validations
                if (!errorMsg) {
                  const maternityQuota = parseInt(getVal('Jatah Cuti Melahirkan')) || 0;
                  const isAccumulated = isLeaveAccumulatedStr === 'Ya';
                  const maxCarryOver = parseInt(getVal('Maksimal Carry-over (*)')) || 0;
                  const currentCarryOver = parseInt(getVal('Jatah Carry-over Saat Ini (*)')) || 0;

                  if (gender === 'Laki-laki' && maternityQuota > 0) {
                    errorMsg = 'Laki-laki tidak berhak mendapatkan jatah cuti melahirkan.';
                  } else if (!isAccumulated && (maxCarryOver > 0 || currentCarryOver > 0)) {
                    errorMsg = 'Akumulasi cuti non-aktif, jatah carry-over harus 0.';
                  } else if (currentCarryOver > maxCarryOver) {
                    errorMsg = 'Jatah carry-over saat ini tidak boleh melebihi batas maksimal.';
                  }
                }
              }

              const isValid = requiredFields.every(field => {
                const val = row[field];
                return val !== undefined && val !== null && String(val).trim() !== '';
              }) && !errorMsg;

              return {
                full_name: getVal('Nama Lengkap (*)'),
                nik_ktp: forceString(getVal('NIK KTP (*)')),
                gender: gender || getVal('Gender (*)'),
                religion: religion || getVal('Agama (*)'),
                dob: formatExcelDate(row['Tgl Lahir (YYYY-MM-DD) (*)']),
                address: getVal('Alamat (*)'),
                phone: forceString(getVal('No Telepon (*)')),
                email: getVal('Email (*)'),
                marital_status: maritalStatus || getVal('Status Nikah (*)'),
                dependents_count: parseInt(getVal('Tanggungan')) || 0,
                internal_nik: internalNik,
                position: getVal('Jabatan (*)'),
                grade: getVal('Departemen/Divisi (*)'),
                location_name: getVal('Lokasi Penempatan (*)'),
                employee_type: employeeType,
                start_date: formatExcelDate(row['Mulai Kontrak (YYYY-MM-DD) (*)']),
                end_date: formatExcelDate(row['Akhir Kontrak (YYYY-MM-DD)']),
                last_education: lastEducation || getVal('Pendidikan Terakhir (*)'),
                major: getVal('Jurusan'),
                grad_date: formatExcelDate(row['Tgl Lulus (YYYY-MM-DD)']),
                emergency_contact_name: getVal('Nama Kontak Darurat'),
                emergency_contact_rel: getVal('Hubungan Kontak Darurat'),
                emergency_contact_phone: forceString(getVal('No HP Kontak Darurat')),
                schedule_name: getVal('Pilih Jadwal Kerja (*)'),
                leave_quota: parseInt(getVal('Jatah Cuti Tahunan (*)')) || 12,
                maternity_leave_quota: parseInt(getVal('Jatah Cuti Melahirkan')) || 0,
                is_leave_accumulated: isLeaveAccumulatedStr === 'Ya',
                max_carry_over_days: parseInt(getVal('Maksimal Carry-over (*)')) || 0,
                carry_over_quota: parseInt(getVal('Jatah Carry-over Saat Ini (*)')) || 0,
                is_presence_limited_checkin: limitCheckin !== 'Tidak',
                is_presence_limited_checkout: limitCheckout !== 'Tidak',
                is_presence_limited_ot_in: limitOtIn !== 'Tidak',
                is_presence_limited_ot_out: limitOtOut !== 'Tidak',
                access_code: String(getVal('Kode Akses (*)')),
                password: String(getVal('Password (*)')),
                photo_google_id: fileMatches.photo_google_id,
                ktp_google_id: fileMatches.ktp_google_id,
                diploma_google_id: fileMatches.diploma_google_id,
                file_sk_id: fileMatches.file_sk_id,
                contract_initial: getVal('Nomor Kontrak') ? {
                  contract_number: getVal('Nomor Kontrak'),
                  contract_type: getVal('Jenis Kontrak') || (getVal('Jenis Karyawan (*)') === 'Tetap' ? 'PKWTT' : 'PKWT'),
                  start_date: formatExcelDate(row['Mulai Kontrak (YYYY-MM-DD)']) || formatExcelDate(row['Tgl Mulai (YYYY-MM-DD) (*)']),
                  end_date: formatExcelDate(row['Akhir Kontrak (YYYY-MM-DD)']) || formatExcelDate(row['Tgl Akhir (YYYY-MM-DD)']),
                  file_id: fileMatches.contract_file_id
                } : null,
                isValid,
                errorMsg
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
    return this.bulkCreate(validData);
  },

  async bulkCreate(accounts: (AccountInput & { location_name?: string, schedule_name?: string })[]) {
    const results = {
      success: 0,
      failed: 0,
      errors: [] as string[]
    };

    // 1. Ambil semua lokasi dan jadwal sekaligus untuk pemetaan (Caching)
    const [{ data: locations }, { data: schedules }] = await Promise.all([
      supabase.from('locations').select('id, name'),
      supabase.from('schedules').select('id, name, type')
    ]);

    const locationMap = new Map(locations?.map(l => [l.name.toLowerCase(), l.id]));
    const locationNameMap = new Map(locations?.map(l => [l.id, l.name]));
    const scheduleMap = new Map(schedules?.map(s => [s.name.toLowerCase(), s]));

    // 2. Proses dalam batch paralel untuk kecepatan (Concurrency: 5)
    const batchSize = 5;
    for (let i = 0; i < accounts.length; i += batchSize) {
      const batch = accounts.slice(i, i + batchSize);
      
      await Promise.all(batch.map(async (acc) => {
        try {
          const { location_name, schedule_name, ...rest } = acc;
          
          // Map location name to ID
          if (!rest.location_id && location_name) {
            const locId = locationMap.get(location_name.toLowerCase());
            if (!locId) {
              throw new Error(`Lokasi "${location_name}" tidak ditemukan.`);
            }
            rest.location_id = locId;
          } else if (!rest.location_id) {
            throw new Error(`Kolom Lokasi Penempatan wajib diisi.`);
          }

          // Map schedule name to ID and Type
          if (schedule_name) {
            const lowerName = schedule_name.toLowerCase();
            if (lowerName === 'fleksibel') {
              rest.schedule_type = 'FLEKSIBEL';
              rest.schedule_id = null;
            } else if (lowerName === 'shift dinamis') {
              rest.schedule_type = 'DINAMIS';
              rest.schedule_id = null;
            } else {
              const sch = scheduleMap.get(lowerName);
              if (!sch) {
                throw new Error(`Jadwal "${schedule_name}" tidak ditemukan.`);
              }
              rest.schedule_id = sch.id;
              rest.schedule_type = sch.type === 1 ? 'Office Hour' : 'Shift';
            }
          }

          // Gunakan versi optimasi yang tidak melakukan query lokasi ulang
          const locName = locationNameMap.get(rest.location_id) || '-';
          await this.createWithPredefinedLocation(rest, locName);
          results.success++;
        } catch (err: any) {
          results.failed++;
          results.errors.push(`${acc.full_name || 'Tanpa Nama'}: ${err.message}`);
        }
      }));
    }

    return results;
  },

  /**
   * Versi optimasi dari create() khusus untuk bulk import.
   * Menghindari query SELECT tambahan untuk nama lokasi karena sudah ada di Map.
   */
  async createWithPredefinedLocation(account: AccountInput & { file_sk_id?: string, file_mcu_id?: string, contract_initial?: any }, locationName: string) {
    const { file_sk_id, file_mcu_id, contract_initial, ...rest } = account;
    const sanitizedAccount = sanitizePayload(rest);
    
    // 1. Insert ke tabel accounts
    const { data, error } = await supabase
      .from('accounts')
      .insert([sanitizedAccount])
      .select();
    
    if (error) throw error;
    const newAccount = data[0] as Account;

    // 2. Buat log & kontrak secara paralel
    const promises = [
      supabase.from('account_career_logs').insert([{
        account_id: newAccount.id,
        position: newAccount.position,
        grade: newAccount.grade,
        location_id: newAccount.location_id,
        location_name: locationName,
        schedule_id: newAccount.schedule_id,
        file_sk_id: file_sk_id || null,
        notes: 'Initial Career Record'
      }])
      // Log kesehatan awal dihapus sesuai permintaan user
    ];

    if (contract_initial && contract_initial.contract_number) {
      promises.push(supabase.from('account_contracts').insert([{
        account_id: newAccount.id,
        contract_number: contract_initial.contract_number,
        contract_type: contract_initial.contract_type || (account.employee_type === 'Kontrak' ? 'PKWT' : account.employee_type),
        start_date: contract_initial.start_date || account.start_date,
        end_date: contract_initial.end_date || account.end_date,
        file_id: contract_initial.file_id || null,
        notes: 'Initial Contract Record'
      }]));
    }

    await Promise.all(promises);
    return newAccount;
  },

  async updateImageByNikOrName(identifier: string, type: 'photo' | 'ktp' | 'ijazah' | 'sk' | 'mcu' | 'kontrak', fileId: string) {
    // Normalize identifier for name matching
    const normalizedId = identifier.toLowerCase().replace(/[^a-z0-9]/g, '');
    
    // 1. Try matching by internal_nik first
    const { data: byNik, error: nikError } = await supabase
      .from('accounts')
      .select('id, full_name')
      .eq('internal_nik', identifier)
      .maybeSingle();

    if (nikError) throw nikError;

    let targetId = byNik?.id;
    let targetName = byNik?.full_name;

    if (!targetId) {
      // 2. Try matching by full_name (normalized)
      const { data: allAccounts, error: allErr } = await supabase.from('accounts').select('id, full_name');
      if (allErr) throw allErr;

      const matches = allAccounts?.filter(acc => 
        acc.full_name.toLowerCase().replace(/[^a-z0-9]/g, '') === normalizedId
      );

      if (!matches || matches.length === 0) {
        throw new Error(`Akun dengan Nama/NIK "${identifier}" tidak ditemukan.`);
      }

      if (matches.length > 1) {
        throw new Error(`Nama "${identifier}" ditemukan lebih dari satu akun (${matches.length}). Gunakan NIK agar spesifik.`);
      }

      targetId = matches[0].id;
      targetName = matches[0].full_name;
    }

    // 3. Update field sesuai tipe
    if (type === 'photo' || type === 'ktp' || type === 'ijazah') {
      const updateData: any = {};
      if (type === 'photo') updateData.photo_google_id = fileId;
      else if (type === 'ktp') updateData.ktp_google_id = fileId;
      else if (type === 'ijazah') updateData.diploma_google_id = fileId;

      const { error } = await supabase.from('accounts').update(updateData).eq('id', targetId);
      if (error) throw error;
    } else if (type === 'sk') {
      const { data: latestLog } = await supabase
        .from('account_career_logs')
        .select('id')
        .eq('account_id', targetId)
        .order('change_date', { ascending: false })
        .limit(1);
      if (latestLog && latestLog.length > 0) {
        await supabase.from('account_career_logs').update({ file_sk_id: fileId }).eq('id', latestLog[0].id);
      }
    } else if (type === 'mcu') {
      const { data: latestLog } = await supabase
        .from('account_health_logs')
        .select('id')
        .eq('account_id', targetId)
        .order('change_date', { ascending: false })
        .limit(1);
      if (latestLog && latestLog.length > 0) {
        await supabase.from('account_health_logs').update({ file_mcu_id: fileId }).eq('id', latestLog[0].id);
      }
    } else if (type === 'kontrak') {
      const { data: latestContract } = await supabase
        .from('account_contracts')
        .select('id')
        .eq('account_id', targetId)
        .order('start_date', { ascending: false })
        .limit(1);
      if (latestContract && latestContract.length > 0) {
        await supabase.from('account_contracts').update({ file_id: fileId }).eq('id', latestContract[0].id);
      }
    }

    return { success: true, name: targetName };
  },

  // Manual Log Management
  async syncAccountWithLatestCareerLog(accountId: string) {
    const { data: latestLog } = await supabase
      .from('account_career_logs')
      .select('*')
      .eq('account_id', accountId)
      .order('change_date', { ascending: false })
      .limit(1)
      .maybeSingle();

    if (latestLog) {
      await supabase.from('accounts').update({
        position: latestLog.position,
        grade: latestLog.grade,
        location_id: latestLog.location_id || null,
        schedule_id: latestLog.schedule_id || null,
        schedule_type: latestLog.schedule_type || null
      }).eq('id', accountId);
    }
  },

  async syncAccountWithLatestHealthLog(accountId: string) {
    const { data: latestLog } = await supabase
      .from('account_health_logs')
      .select('*')
      .eq('account_id', accountId)
      .order('change_date', { ascending: false })
      .limit(1)
      .maybeSingle();

    if (latestLog) {
      await supabase.from('accounts').update({
        mcu_status: latestLog.mcu_status,
        health_risk: latestLog.health_risk
      }).eq('id', accountId);
    }
  },

  async createCareerLog(logInput: CareerLogInput) {
    // Filtrasi: Pastikan hanya kolom yang ada di tabel account_career_logs yang dikirim
    const { account_id, position, grade, location_name, file_sk_id, notes, location_id, schedule_id, schedule_type, change_date } = logInput;
    const payload = sanitizePayload({ account_id, position, grade, location_name, file_sk_id, notes, change_date, location_id, schedule_id, schedule_type });
    
    const { data, error } = await supabase
      .from('account_career_logs')
      .insert([payload])
      .select()
      .single();

    if (error) {
      console.error("CAREER_LOG_CREATE_ERROR:", error.message);
      throw error;
    }

    // Sinkronisasi ke profil utama dengan log terbaru
    await this.syncAccountWithLatestCareerLog(account_id);

    return data as CareerLog;
  },

  async updateCareerLog(id: string, logInput: Partial<CareerLogInput>) {
    const { account_id, position, grade, location_name, file_sk_id, notes, location_id, schedule_id, schedule_type, change_date } = logInput;
    const payload = sanitizePayload({ account_id, position, grade, location_name, file_sk_id, notes, change_date, location_id, schedule_id, schedule_type });

    const { data, error } = await supabase
      .from('account_career_logs')
      .update(payload)
      .eq('id', id)
      .select()
      .single();

    if (error) {
      console.error("CAREER_LOG_UPDATE_ERROR:", error.message);
      throw error;
    }

    if (data && data.account_id) {
      await this.syncAccountWithLatestCareerLog(data.account_id);
    }

    return data as CareerLog;
  },

  async deleteCareerLog(id: string) {
    // 1. Ambil ID file & Account ID
    const { data } = await supabase.from('account_career_logs').select('file_sk_id, account_id').eq('id', id).single();
    
    // 2. Hapus file dari Drive
    if (data?.file_sk_id) {
      const { googleDriveService } = await import('./googleDriveService');
      await googleDriveService.deleteFile(data.file_sk_id);
    }

    // 3. Hapus dari DB
    const { error } = await supabase
      .from('account_career_logs')
      .delete()
      .eq('id', id);
    if (error) throw error;

    // 4. Sinkronisasi
    if (data?.account_id) {
      await this.syncAccountWithLatestCareerLog(data.account_id);
    }
    return true;
  },

  async createHealthLog(logInput: HealthLogInput) {
    // Filtrasi: Hapus field career (location_id, location_name) yang sering terbawa dari state form
    const { account_id, mcu_status, health_risk, file_mcu_id, notes, change_date } = logInput;
    const payload = sanitizePayload({ account_id, mcu_status, health_risk, file_mcu_id, notes, change_date });

    const { data, error } = await supabase
      .from('account_health_logs')
      .insert([payload])
      .select()
      .single();

    if (error) {
      console.error("HEALTH_LOG_CREATE_ERROR:", error.message);
      throw error;
    }

    // Sinkronisasi ke profil utama dengan log terbaru
    await this.syncAccountWithLatestHealthLog(account_id);

    return data as HealthLog;
  },

  async updateHealthLog(id: string, logInput: Partial<HealthLogInput>) {
    const { account_id, mcu_status, health_risk, file_mcu_id, notes, change_date } = logInput;
    const payload = sanitizePayload({ account_id, mcu_status, health_risk, file_mcu_id, notes, change_date });

    const { data, error } = await supabase
      .from('account_health_logs')
      .update(payload)
      .eq('id', id)
      .select()
      .single();

    if (error) {
      console.error("HEALTH_LOG_UPDATE_ERROR:", error.message);
      throw error;
    }

    if (data && data.account_id) {
      await this.syncAccountWithLatestHealthLog(data.account_id);
    }

    return data as HealthLog;
  },

  async deleteHealthLog(id: string) {
    // 1. Ambil ID file & Account ID
    const { data } = await supabase.from('account_health_logs').select('file_mcu_id, account_id').eq('id', id).single();
    
    // 2. Hapus file dari Drive
    if (data?.file_mcu_id) {
      const { googleDriveService } = await import('./googleDriveService');
      await googleDriveService.deleteFile(data.file_mcu_id);
    }

    // 3. Hapus dari DB
    const { error } = await supabase
      .from('account_health_logs')
      .delete()
      .eq('id', id);
    if (error) throw error;

    // 4. Sinkronisasi
    if (data?.account_id) {
      await this.syncAccountWithLatestHealthLog(data.account_id);
    }
    return true;
  }
};
