
import React, { useState, useEffect, useRef } from 'react';
import { 
  Plus, Search, Users, Grid, List as ListIcon, 
  ArrowLeft, UserCircle, UserCheck, UserX,
  History, FileBadge, Award, Activity, ShieldAlert,
  Download, Upload, Image as ImageIcon,
  MapPin, Mail, Phone, Edit2, LogOut, Shield, Briefcase, Trash2
} from 'lucide-react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import Swal from 'sweetalert2';
import { accountService } from '../../services/accountService';
import { locationService } from '../../services/locationService';
import { scheduleService } from '../../services/scheduleService';
import { financeService } from '../../services/financeService';
import { authService } from '../../services/authService';
import { Account, AccountInput, AuthUser, SalaryScheme } from '../../types';
import AccountForm from './AccountForm';
import AccountDetail from './AccountDetail';
import { CardSkeleton } from '../../components/Common/Skeleton';
import LoadingSpinner from '../../components/Common/LoadingSpinner';
import { googleDriveService } from '../../services/googleDriveService';

// Import sub-modules
import CareerLogMain from '../career/CareerLogMain';
import HealthLogMain from '../health/HealthLogMain';
import ContractMain from '../contract/ContractMain';
import CertificationMain from '../certification/CertificationMain';
import DisciplineMain from '../discipline/DisciplineMain';

interface AccountMainProps {
  user?: AuthUser | null;
  setUser?: (user: AuthUser | null) => void;
  isSelfProfile?: boolean;
}

const AccountMain: React.FC<AccountMainProps> = ({ user, setUser, isSelfProfile = false }) => {
  const [accounts, setAccounts] = useState<Account[]>([]);
  const [selfAccount, setSelfAccount] = useState<Account | null>(null);
  const [salaryScheme, setSalaryScheme] = useState<SalaryScheme | null>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [viewMode, setViewMode] = useState<'grid' | 'list'>('list');
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState<'aktif' | 'non-aktif'>('aktif');
  const [showForm, setShowForm] = useState(false);
  const [selectedAccountId, setSelectedAccountId] = useState<string | null>(null);
  const [editingAccount, setEditingAccount] = useState<Account | null>(null);
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  
  const fileInputRef = useRef<HTMLInputElement>(null);
  const bulkImageInputRef = useRef<HTMLInputElement>(null);

  // Tab State Internal Modul Akun
  const [activeSubTab, setActiveSubTab] = useState<'data' | 'career' | 'contract' | 'cert' | 'health' | 'discipline'>('data');

  useEffect(() => {
    if (isSelfProfile && user) {
      fetchSelfAccount();
    } else {
      fetchAccounts();
    }
  }, [isSelfProfile, user]);

  const fetchSelfAccount = async () => {
    if (!user) return;
    try {
      setIsLoading(true);
      const acc = await accountService.getById(user.id);
      setSelfAccount(acc);
      
      // Fetch salary scheme
      const assignment = await financeService.getAssignmentByAccountId(user.id);
      if (assignment && assignment.scheme) {
        setSalaryScheme(assignment.scheme);
      }
    } catch (error) {
      console.error('Error fetching self account:', error);
    } finally {
      setIsLoading(false);
    }
  };

  const fetchAccounts = async () => {
    try {
      setIsLoading(true);
      const data = await accountService.getAll();
      setAccounts(data as any);
    } catch (error) {
      Swal.fire('Gagal', 'Gagal memuat data akun', 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleCreate = async (input: AccountInput) => {
    setIsSaving(true);
    const tempId = `temp-${Math.random().toString(36).substring(7)}`;
    const optimisticAccount: Account = { 
      ...input, 
      id: tempId, 
      created_at: new Date().toISOString()
    };
    
    setAccounts(prev => [optimisticAccount, ...prev]);
    setShowForm(false);

    try {
      const created = await accountService.create(input);
      setAccounts(prev => prev.map(acc => acc.id === tempId ? created : acc));
      Swal.fire({
        title: 'Berhasil!',
        text: 'Akun baru telah ditambahkan.',
        icon: 'success',
        timer: 1500,
        showConfirmButton: false
      });
    } catch (error) {
      setAccounts(prev => prev.filter(acc => acc.id !== tempId));
      Swal.fire('Gagal', 'Terjadi kesalahan saat menyimpan data', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const downloadTemplate = async () => {
    try {
      setIsSaving(true);
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
        'Nama Lengkap', 'NIK KTP', 'Gender', 'Agama', 'Tgl Lahir (YYYY-MM-DD)', 
        'Alamat', 'No Telepon', 'Email', 'Status Nikah', 'Tanggungan', 
        'NIK Internal', 'Jabatan', 'Golongan', 'Lokasi Penempatan', 
        'Jenis Karyawan', 'Tgl Mulai (YYYY-MM-DD)', 'Tgl Akhir (YYYY-MM-DD)',
        'Pendidikan Terakhir', 'Jurusan', 'Tgl Lulus (YYYY-MM-DD)',
        'Nama Kontak Darurat', 'Hubungan Kontak Darurat', 'No HP Kontak Darurat',
        'Pilih Jadwal Kerja', 'Jatah Cuti Tahunan', 'Jatah Cuti Melahirkan', 
        'Akumulasi Cuti (Ya/Tidak)', 'Maksimal Carry-over', 'Jatah Carry-over Saat Ini',
        'Batasi Check-in Datang (Ya/Tidak)', 'Batasi Check-out Pulang (Ya/Tidak)', 
        'Batasi Check-in Lembur (Ya/Tidak)', 'Batasi Check-out Lembur (Ya/Tidak)',
        'Kode Akses', 'Password'
      ];

      templateSheet.addRow(headers);
      
      // Style headers
      templateSheet.getRow(1).font = { bold: true };
      templateSheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE0E0E0' }
      };

      // Prepare lists for dropdowns
      const genderList = ['Laki-laki', 'Perempuan'];
      const religionList = ['Islam', 'Kristen', 'Katolik', 'Hindu', 'Budha', 'Konghucu', 'Kepercayaan Lain'];
      const maritalList = ['Belum Menikah', 'Menikah', 'Cerai Hidup', 'Cerai Mati'];
      const empTypeList = ['Tetap', 'Kontrak', 'Harian', 'Magang'];
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

      // Apply Data Validations (Dropdowns)
      // We apply validation to 100 rows
      for (let i = 2; i <= 101; i++) {
        // Gender (Col 3)
        templateSheet.getCell(`C${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$A$1:$A$${genderList.length}`] };
        // Agama (Col 4)
        templateSheet.getCell(`D${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$B$1:$B$${religionList.length}`] };
        // Status Nikah (Col 9)
        templateSheet.getCell(`I${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$C$1:$C$${maritalList.length}`] };
        // Lokasi Penempatan (Col 14)
        templateSheet.getCell(`N${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$G$1:$G$${locList.length}`] };
        // Jenis Karyawan (Col 15)
        templateSheet.getCell(`O${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$D$1:$D$${empTypeList.length}`] };
        // Pilih Jadwal Kerja (Col 24)
        templateSheet.getCell(`X${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$H$1:$H$${schList.length}`] };
        // Akumulasi Cuti (Col 27)
        templateSheet.getCell(`AA${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$E$1:$E$2`] };
        // Radius Limits (Col 30-33)
        templateSheet.getCell(`AD${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$E$1:$E$2`] };
        templateSheet.getCell(`AE${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$E$1:$E$2`] };
        templateSheet.getCell(`AF${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$E$1:$E$2`] };
        templateSheet.getCell(`AG${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`Lists!$E$1:$E$2`] };
      }

      // Auto-size columns
      templateSheet.columns.forEach(column => {
        column.width = 20;
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = 'template_impor_akun.xlsx';
      anchor.click();
      window.URL.revokeObjectURL(url);
    } catch (error) {
      console.error(error);
      Swal.fire('Gagal', 'Gagal menyiapkan template', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const buffer = evt.target?.result as ArrayBuffer;
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        const worksheet = workbook.getWorksheet(1);
        
        if (!worksheet) return;

        const jsonData: any[] = [];
        const headers: string[] = [];
        
        worksheet.getRow(1).eachCell((cell, colNumber) => {
          headers[colNumber] = cell.text;
        });

        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) return;
          const rowData: any = {};
          row.eachCell((cell, colNumber) => {
            rowData[headers[colNumber]] = cell.value;
          });
          jsonData.push(rowData);
        });

        if (jsonData.length === 0) {
          Swal.fire('Gagal', 'File Excel kosong atau tidak valid', 'error');
          return;
        }

        const confirm = await Swal.fire({
          title: 'Konfirmasi Impor',
          text: `Apakah Anda yakin ingin mengimpor ${jsonData.length} data akun?`,
          icon: 'question',
          showCancelButton: true,
          confirmButtonColor: '#006E62',
          confirmButtonText: 'Ya, Impor Sekarang',
          cancelButtonText: 'Batal'
        });

        if (confirm.isConfirmed) {
          setIsSaving(true);
          try {
            const formattedData: (AccountInput & { location_name?: string, schedule_name?: string })[] = jsonData.map((row: any) => {
              // Helper to handle ExcelJS cell values which might be objects
              const getVal = (val: any) => {
                if (val && typeof val === 'object' && 'result' in val) return val.result;
                if (val && typeof val === 'object' && 'text' in val) return val.text;
                return val;
              };

              // Helper for robust date formatting from Excel
              const formatExcelDate = (val: any) => {
                const raw = getVal(val);
                if (!raw) return null;
                
                // If it's a Date object
                if (raw instanceof Date) {
                  return raw.toISOString().split('T')[0];
                }
                
                // If it's a serial number (Excel date)
                if (typeof raw === 'number') {
                  const date = new Date((raw - 25569) * 86400 * 1000);
                  return date.toISOString().split('T')[0];
                }
                
                // If it's a string, try to parse it
                if (typeof raw === 'string') {
                  const trimmed = raw.trim();
                  if (!trimmed) return null;
                  
                  // Already YYYY-MM-DD
                  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) return trimmed;
                  
                  // Try parsing common formats
                  const date = new Date(trimmed);
                  if (!isNaN(date.getTime())) {
                    return date.toISOString().split('T')[0];
                  }
                }
                
                return null;
              };

              return {
                full_name: getVal(row['Nama Lengkap']),
                nik_ktp: String(getVal(row['NIK KTP']) || ''),
                gender: getVal(row['Gender']) || 'Laki-laki',
                religion: getVal(row['Agama']) || 'Islam',
                dob: formatExcelDate(row['Tgl Lahir (YYYY-MM-DD)']),
                address: getVal(row['Alamat']),
                phone: String(getVal(row['No Telepon']) || ''),
                email: getVal(row['Email']),
                marital_status: getVal(row['Status Nikah']) || 'Belum Menikah',
                dependents_count: parseInt(getVal(row['Tanggungan'])) || 0,
                emergency_contact_name: getVal(row['Nama Kontak Darurat']) || '',
                emergency_contact_rel: getVal(row['Hubungan Kontak Darurat']) || '',
                emergency_contact_phone: String(getVal(row['No HP Kontak Darurat']) || ''),
                last_education: getVal(row['Pendidikan Terakhir']) || 'Sarjana',
                major: getVal(row['Jurusan']) || '',
                internal_nik: String(getVal(row['NIK Internal']) || ''),
                position: getVal(row['Jabatan']),
                grade: getVal(row['Golongan']),
                location_id: null,
                location_name: getVal(row['Lokasi Penempatan']),
                schedule_id: null,
                schedule_name: getVal(row['Pilih Jadwal Kerja']),
                employee_type: getVal(row['Jenis Karyawan']) || 'Tetap',
                start_date: formatExcelDate(row['Tgl Mulai (YYYY-MM-DD)']),
                end_date: formatExcelDate(row['Tgl Akhir (YYYY-MM-DD)']),
                schedule_type: getVal(row['Pilih Jadwal Kerja']) || 'Office Hour',
                leave_quota: parseInt(getVal(row['Jatah Cuti Tahunan'])) || 12,
                is_leave_accumulated: getVal(row['Akumulasi Cuti (Ya/Tidak)']) === 'Ya',
                max_carry_over_days: parseInt(getVal(row['Maksimal Carry-over'])) || 0,
                carry_over_quota: parseInt(getVal(row['Jatah Carry-over Saat Ini'])) || 0,
                maternity_leave_quota: parseInt(getVal(row['Jatah Cuti Melahirkan'])) || 0,
                is_presence_limited_checkin: getVal(row['Batasi Check-in Datang (Ya/Tidak)']) !== 'Tidak',
                is_presence_limited_checkout: getVal(row['Batasi Check-out Pulang (Ya/Tidak)']) !== 'Tidak',
                is_presence_limited_ot_in: getVal(row['Batasi Check-in Lembur (Ya/Tidak)']) !== 'Tidak',
                is_presence_limited_ot_out: getVal(row['Batasi Check-out Lembur (Ya/Tidak)']) !== 'Tidak',
                access_code: String(getVal(row['Kode Akses']) || ''),
                password: String(getVal(row['Password']) || '123456'),
                role: 'user'
              };
            });

            const res = await accountService.bulkCreate(formattedData);
            
            setIsSaving(false); // Matikan loading segera setelah proses selesai

            let message = `${res.success} akun berhasil diimpor.`;
            if (res.failed > 0) {
              message += `\n${res.failed} akun gagal diimpor.`;
            }

            await Swal.fire({
              title: res.failed === 0 ? 'Berhasil!' : 'Selesai dengan Catatan',
              text: message,
              icon: res.failed === 0 ? 'success' : 'warning',
              footer: res.failed > 0 ? `<div style="max-height: 150px; overflow-y: auto; text-align: left; font-size: 10px; color: #ef4444; background: #fef2f2; padding: 8px; border-radius: 4px; border: 1px solid #fee2e2;">${res.errors.join('<br/>')}</div>` : null
            });

            fetchAccounts();
          } catch (error) {
            setIsSaving(false);
            console.error(error);
            Swal.fire('Gagal', 'Terjadi kesalahan sistem saat impor', 'error');
          } finally {
            if (fileInputRef.current) fileInputRef.current.value = '';
          }
        } else {
          if (fileInputRef.current) fileInputRef.current.value = '';
        }
      } catch (error) {
        console.error(error);
        Swal.fire('Gagal', 'Gagal membaca file Excel', 'error');
        if (fileInputRef.current) fileInputRef.current.value = '';
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleBulkImageUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const confirm = await Swal.fire({
      title: 'Konfirmasi Bulk Upload',
      text: `Anda memilih ${files.length} file. Sistem akan mencocokkan file berdasarkan Nama atau NIK (Format: Nama_photo.jpg). Lanjutkan?`,
      icon: 'question',
      showCancelButton: true,
      confirmButtonColor: '#006E62',
      confirmButtonText: 'Ya, Unggah',
      cancelButtonText: 'Batal'
    });

    if (!confirm.isConfirmed) {
      if (bulkImageInputRef.current) bulkImageInputRef.current.value = '';
      return;
    }

    setIsSaving(true);
    const results = { success: 0, failed: 0, errors: [] as string[] };

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      const fileName = file.name.split('.').slice(0, -1).join('.'); // Remove extension
      
      try {
        // Parse filename: [Identifier]_[Type]
        const parts = fileName.split('_');
        if (parts.length < 2) {
          throw new Error(`Format nama file tidak valid. Gunakan: Nama_photo.jpg`);
        }

        const typeStr = parts[parts.length - 1].toLowerCase();
        const identifier = parts.slice(0, -1).join('_'); // Handle names with underscores

        let type: 'photo' | 'ktp' | 'ijazah' | 'sk' | 'mcu' | 'kontrak';
        let folderId: string;

        if (typeStr.includes('photo') || typeStr.includes('foto')) {
          type = 'photo';
          folderId = 'photos';
        } else if (typeStr.includes('ktp')) {
          type = 'ktp';
          folderId = 'ktp';
        } else if (typeStr.includes('ijazah') || typeStr.includes('diploma')) {
          type = 'ijazah';
          folderId = 'ijazah';
        } else if (typeStr.includes('sk')) {
          type = 'sk';
          folderId = 'sk';
        } else if (typeStr.includes('mcu')) {
          type = 'mcu';
          folderId = 'mcu';
        } else if (typeStr.includes('kontrak') || typeStr.includes('contract')) {
          type = 'kontrak';
          folderId = 'contracts';
        } else {
          throw new Error(`Tipe dokumen "${typeStr}" tidak dikenal. Gunakan: photo, ktp, ijazah, sk, mcu, atau kontrak.`);
        }

        // 1. Upload to Google Drive
        const driveFileId = await googleDriveService.uploadFile(file, folderId);

        // 2. Update Database
        await accountService.updateImageByNikOrName(identifier, type, driveFileId);
        results.success++;
      } catch (err: any) {
        results.failed++;
        results.errors.push(`${file.name}: ${err.message}`);
      }
    }

    setIsSaving(false);
    if (bulkImageInputRef.current) bulkImageInputRef.current.value = '';

    let message = `${results.success} file berhasil diunggah & dicocokkan.`;
    if (results.failed > 0) {
      message += `\n${results.failed} file gagal.`;
    }

    await Swal.fire({
      title: results.failed === 0 ? 'Bulk Upload Berhasil' : 'Selesai dengan Catatan',
      text: message,
      icon: results.failed === 0 ? 'success' : 'warning',
      footer: results.failed > 0 ? `<div style="max-height: 150px; overflow-y: auto; text-align: left; font-size: 10px; color: #ef4444; background: #fef2f2; padding: 8px; border-radius: 4px; border: 1px solid #fee2e2;">${results.errors.join('<br/>')}</div>` : null
    });

    fetchAccounts();
  };

  const handleUpdate = async (id: string, input: Partial<AccountInput>) => {
    setIsSaving(true);
    try {
      const updated = await accountService.update(id, input);
      setAccounts(prev => prev.map(acc => acc.id === id ? updated : acc));
      setEditingAccount(null);
      setShowForm(false);
      Swal.fire({
        title: 'Terupdate!',
        text: 'Data akun berhasil diperbarui.',
        icon: 'success',
        timer: 1500,
        showConfirmButton: false
      });
    } catch (error) {
      Swal.fire('Gagal', 'Gagal memperbarui data', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleDelete = async (id: string) => {
    const result = await Swal.fire({
      title: 'Apakah Anda yakin?',
      text: "Akun ini akan dihapus permanen.",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#006E62',
      cancelButtonColor: '#ef4444',
      confirmButtonText: 'Ya, hapus!',
      cancelButtonText: 'Batal'
    });

    if (result.isConfirmed) {
      setIsSaving(true);
      try {
        await accountService.delete(id);
        setAccounts(prev => prev.filter(acc => acc.id !== id));
        setSelectedIds(prev => prev.filter(sid => sid !== id));
        setSelectedAccountId(null);
        Swal.fire('Terhapus!', 'Akun telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Gagal menghapus data', 'error');
      } finally {
        setIsSaving(false);
      }
    }
  };

  const handleBulkDelete = async () => {
    if (selectedIds.length === 0) return;

    const result = await Swal.fire({
      title: 'Hapus Masal?',
      text: `Apakah Anda yakin ingin menghapus ${selectedIds.length} akun terpilih secara permanen?`,
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#ef4444',
      cancelButtonColor: '#6b7280',
      confirmButtonText: 'Ya, Hapus Semua',
      cancelButtonText: 'Batal'
    });

    if (result.isConfirmed) {
      setIsSaving(true);
      try {
        for (const id of selectedIds) {
          await accountService.delete(id);
        }
        setAccounts(prev => prev.filter(acc => !selectedIds.includes(acc.id)));
        setSelectedIds([]);
        Swal.fire('Berhasil!', 'Data terpilih telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Beberapa data mungkin gagal dihapus.', 'error');
        fetchAccounts();
      } finally {
        setIsSaving(false);
      }
    }
  };

  const toggleSelect = (id: string) => {
    setSelectedIds(prev => 
      prev.includes(id) ? prev.filter(i => i !== id) : [...prev, id]
    );
  };

  const toggleSelectAll = () => {
    if (selectedIds.length === filteredAccounts.length) {
      setSelectedIds([]);
    } else {
      setSelectedIds(filteredAccounts.map(acc => acc.id));
    }
  };

  const today = new Date().toISOString().split('T')[0];
  
  const activeAccounts = accounts.filter(acc => !acc.end_date || acc.end_date > today);
  const inactiveAccounts = accounts.filter(acc => acc.end_date && acc.end_date <= today);

  const filteredAccounts = (statusFilter === 'aktif' ? activeAccounts : inactiveAccounts).filter(acc => {
    const searchStr = (acc.search_all || `${acc.full_name} ${acc.internal_nik} ${acc.position}`).toLowerCase();
    return searchStr.includes(searchTerm.toLowerCase());
  });

  // Halaman Detail
  if (selectedAccountId) {
    return (
      <div className="animate-in fade-in slide-in-from-right duration-300">
        <button 
          onClick={() => setSelectedAccountId(null)}
          className="mb-4 flex items-center gap-2 text-gray-500 hover:text-[#006E62] transition-colors font-bold text-xs uppercase"
        >
          <ArrowLeft size={16} /> Kembali ke Daftar
        </button>
        <AccountDetail
          id={selectedAccountId}
          onClose={() => setSelectedAccountId(null)}
          onEdit={(acc) => { setSelectedAccountId(null); setEditingAccount(acc); setShowForm(true); }}
          onDelete={(id) => handleDelete(id)}
        />
      </div>
    );
  }

  const DetailRow = ({ label, value }: { label: string, value: any }) => (
    <div>
      <p className="text-[9px] font-bold text-gray-400 uppercase tracking-tighter mb-0.5">{label}</p>
      <p className="text-xs text-gray-700 font-medium leading-tight">{value || '-'}</p>
    </div>
  );

  // Self Profile View
  if (isSelfProfile) {
    if (isLoading) return <div className="flex justify-center py-20"><div className="w-8 h-8 border-4 border-[#006E62] border-t-transparent rounded-full animate-spin"></div></div>;
    if (!selfAccount) return <div className="text-center py-20 text-gray-400">Data profil tidak ditemukan.</div>;

    if (showForm) {
      return (
        <AccountForm 
          onClose={() => setShowForm(false)} 
          onSubmit={async (data) => {
            setIsSaving(true);
            try {
              const updated = await accountService.update(selfAccount.id, data);
              setSelfAccount(updated);
              setShowForm(false);
              Swal.fire('Berhasil', 'Profil Anda telah diperbarui.', 'success');
            } catch (error) {
              Swal.fire('Gagal', 'Gagal memperbarui profil.', 'error');
            } finally {
              setIsSaving(false);
            }
          }}
          initialData={selfAccount}
          isSelfEdit={true}
        />
      );
    }

    return (
      <div className="space-y-6 animate-in fade-in duration-300 pb-20">
        {isSaving && <LoadingSpinner />}
        
        {/* Header Profile */}
        <div className="bg-white rounded-md border border-gray-100 p-6 flex flex-col md:flex-row gap-6 items-start shadow-sm">
          <div className="w-32 h-32 rounded-md border-4 border-gray-50 overflow-hidden shrink-0 shadow-inner">
            {selfAccount.photo_google_id ? (
              <img src={googleDriveService.getFileUrl(selfAccount.photo_google_id)} className="w-full h-full object-cover" />
            ) : (
              <div className="w-full h-full flex items-center justify-center bg-gray-100 text-gray-400"><UserCircle size={48} /></div>
            )}
          </div>
          
          <div className="flex-1 space-y-2">
            <div className="flex items-center gap-3">
               <h2 className="text-2xl font-bold text-gray-800 tracking-tight">{selfAccount.full_name}</h2>
               <span className="px-2 py-0.5 text-[10px] font-bold uppercase rounded bg-[#006E62]/10 text-[#006E62]">
                 {selfAccount.employee_type}
               </span>
            </div>
            <p className="text-sm font-bold text-gray-500 uppercase tracking-widest">{selfAccount.position} • {selfAccount.grade} • {selfAccount.internal_nik}</p>
            <div className="flex flex-wrap gap-4 pt-2">
               <div className="flex items-center gap-1.5 text-xs text-gray-600"><MapPin size={14} className="text-gray-400" /> {selfAccount.location?.name || '-'}</div>
               <div className="flex items-center gap-1.5 text-xs text-gray-600"><Mail size={14} className="text-gray-400" /> {selfAccount.email || '-'}</div>
               <div className="flex items-center gap-1.5 text-xs text-gray-600"><Phone size={14} className="text-gray-400" /> {selfAccount.phone || '-'}</div>
            </div>
          </div>

          <div className="flex flex-col gap-2">
             <button 
               onClick={() => setShowForm(true)} 
               className="flex items-center gap-2 bg-[#006E62] text-white px-4 py-2 rounded shadow-sm hover:bg-[#005a50] transition-all text-xs font-bold uppercase"
             >
               <Edit2 size={14} /> Edit Profil
             </button>
             <button 
               onClick={async () => {
                 const res = await Swal.fire({
                   title: 'Logout?',
                   text: 'Anda akan keluar dari aplikasi.',
                   icon: 'question',
                   showCancelButton: true,
                   confirmButtonColor: '#ef4444',
                   confirmButtonText: 'Ya, Logout',
                   cancelButtonText: 'Batal'
                 });
                 if (res.isConfirmed && setUser) {
                   authService.logout();
                   setUser(null);
                 }
               }} 
               className="flex items-center gap-2 bg-red-50 text-red-600 px-4 py-2 rounded border border-red-100 hover:bg-red-100 transition-all text-xs font-bold uppercase"
             >
               <LogOut size={14} /> Keluar
             </button>
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          {/* Identitas Lengkap */}
          <div className="bg-white border border-gray-100 p-5 rounded-md shadow-sm space-y-4">
            <div className="flex items-center gap-2 border-b border-gray-50 pb-3 mb-2">
              <Users size={16} className="text-[#006E62]" />
              <h4 className="text-[11px] font-bold uppercase tracking-widest text-gray-400">Identitas Lengkap</h4>
            </div>
            <div className="grid grid-cols-2 gap-4">
               <DetailRow label="NIK KTP" value={selfAccount.nik_ktp} />
               <DetailRow label="Tanggal Lahir" value={selfAccount.dob} />
               <DetailRow label="Gender" value={selfAccount.gender} />
               <DetailRow label="Agama" value={selfAccount.religion} />
               <DetailRow label="Status Nikah" value={selfAccount.marital_status} />
               <DetailRow label="Tanggungan" value={selfAccount.dependents_count} />
            </div>
            <DetailRow label="Alamat Domisili" value={selfAccount.address} />
            <div className="grid grid-cols-2 gap-4">
               <DetailRow label="Pendidikan" value={selfAccount.last_education} />
               <DetailRow label="Jurusan" value={selfAccount.major} />
            </div>
            <div className="pt-2 border-t border-gray-50">
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-tighter mb-2">Kontak Darurat</p>
              <div className="grid grid-cols-2 gap-4">
                <DetailRow label="Nama" value={selfAccount.emergency_contact_name} />
                <DetailRow label="Hubungan" value={selfAccount.emergency_contact_rel} />
                <DetailRow label="No HP" value={selfAccount.emergency_contact_phone} />
              </div>
            </div>
          </div>

          {/* Karier & Penempatan */}
          <div className="bg-white border border-gray-100 p-5 rounded-md shadow-sm space-y-4">
            <div className="flex items-center gap-2 border-b border-gray-50 pb-3 mb-2">
              <Briefcase size={16} className="text-[#006E62]" />
              <h4 className="text-[11px] font-bold uppercase tracking-widest text-gray-400">Karier & Penempatan</h4>
            </div>
            <div className="grid grid-cols-2 gap-4">
              <DetailRow label="Jabatan" value={selfAccount.position} />
              <DetailRow label="Golongan" value={selfAccount.grade} />
              <DetailRow label="NIK Internal" value={selfAccount.internal_nik} />
              <DetailRow label="Tipe Karyawan" value={selfAccount.employee_type} />
              <DetailRow label="Jadwal" value={selfAccount.schedule_type} />
              <DetailRow label="Mulai Kerja" value={selfAccount.start_date} />
            </div>
            <DetailRow label="Lokasi Penempatan" value={selfAccount.location?.name} />
          </div>

          {/* Keamanan & Akses */}
          <div className="bg-white border border-gray-100 p-5 rounded-md shadow-sm space-y-4">
            <div className="flex items-center gap-2 border-b border-gray-50 pb-3 mb-2">
              <Shield size={16} className="text-[#006E62]" />
              <h4 className="text-[11px] font-bold uppercase tracking-widest text-gray-400">Keamanan & Akses</h4>
            </div>
            <div className="space-y-4">
              <div className="p-3 bg-gray-50 rounded border border-gray-100">
                <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Kode Akses</p>
                <p className="text-lg font-mono font-bold text-[#006E62] tracking-[0.2em]">{selfAccount.access_code}</p>
              </div>
              <div className="p-3 bg-gray-50 rounded border border-gray-100">
                <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Password</p>
                <p className="text-sm font-mono text-gray-600">********</p>
                <p className="text-[9px] text-gray-400 italic mt-1">Gunakan tombol Edit Profil untuk mengubah password.</p>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  const SubTab = ({ id, label, icon: Icon }: { id: any, label: string, icon: any }) => (
    <button
      onClick={() => setActiveSubTab(id)}
      className={`flex items-center gap-2 px-5 py-3 text-[11px] font-bold uppercase tracking-widest border-b-2 transition-all whitespace-nowrap ${
        activeSubTab === id ? 'border-[#006E62] text-[#006E62] bg-emerald-50/30' : 'border-transparent text-gray-400 hover:text-gray-600'
      }`}
    >
      <Icon size={14} />
      {label}
    </button>
  );

  return (
    <div className="space-y-6">
      {isSaving && <LoadingSpinner />}

      {/* Internal Sub-Tabs Navigation */}
      <div className="flex border-b border-gray-100 overflow-x-auto scrollbar-none bg-white -mt-4 mb-6">
        <SubTab id="data" label="Data Akun" icon={Users} />
        <SubTab id="career" label="Log Karir" icon={History} />
        <SubTab id="contract" label="Kontrak Kerja" icon={FileBadge} />
        <SubTab id="cert" label="Sertifikasi" icon={Award} />
        <SubTab id="health" label="Log Kesehatan" icon={Activity} />
        <SubTab id="discipline" label="Peringatan & Keluar" icon={ShieldAlert} />
      </div>

      {activeSubTab === 'data' ? (
        <div className="space-y-6 animate-in fade-in duration-300">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div className="relative flex-1 max-w-md">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
              <input
                type="text"
                placeholder="Cari (Nama, NIK, Jabatan)..."
                className="w-full pl-10 pr-4 py-2 bg-gray-50 border border-gray-200 rounded-md focus:outline-none focus:ring-2 focus:ring-[#006E62] focus:border-transparent transition-all text-sm"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
            </div>
            
            <div className="flex items-center gap-2">
              {selectedIds.length > 0 && (
                <button 
                  onClick={handleBulkDelete}
                  className="flex items-center gap-2 bg-red-50 text-red-600 px-3 py-2 rounded-md border border-red-100 hover:bg-red-100 transition-all text-xs font-bold uppercase mr-2"
                >
                  <Trash2 size={14} />
                  Hapus ({selectedIds.length})
                </button>
              )}
              <div className="flex items-center gap-1 mr-2">
                <button 
                  onClick={downloadTemplate}
                  className="flex items-center gap-1.5 text-gray-500 hover:text-[#006E62] px-2 py-1.5 text-[10px] font-bold uppercase transition-colors"
                  title="Unduh Template CSV"
                >
                  <Download size={14} />
                  <span className="hidden sm:inline">Template</span>
                </button>
                <button 
                  onClick={() => fileInputRef.current?.click()}
                  className="flex items-center gap-1.5 text-gray-500 hover:text-[#006E62] px-2 py-1.5 text-[10px] font-bold uppercase transition-colors"
                  title="Impor Akun dari CSV"
                >
                  <Upload size={14} />
                  <span className="hidden sm:inline">Impor</span>
                </button>
                <button 
                  onClick={() => bulkImageInputRef.current?.click()}
                  className="flex items-center gap-1.5 text-gray-500 hover:text-[#006E62] px-2 py-1.5 text-[10px] font-bold uppercase transition-colors"
                  title="Bulk Upload Foto/Dokumen"
                >
                  <ImageIcon size={14} />
                  <span className="hidden sm:inline">Bulk Foto</span>
                </button>
                <input 
                  type="file" 
                  ref={fileInputRef} 
                  onChange={handleImport} 
                  accept=".xlsx,.csv" 
                  className="hidden" 
                />
                <input 
                  type="file" 
                  ref={bulkImageInputRef} 
                  onChange={handleBulkImageUpload} 
                  multiple 
                  accept="image/*,application/pdf" 
                  className="hidden" 
                />
              </div>

              <div className="flex border border-gray-200 rounded-md overflow-hidden bg-white">
                <button 
                  onClick={() => setViewMode('grid')}
                  className={`p-2 transition-colors ${viewMode === 'grid' ? 'bg-gray-100 text-[#006E62]' : 'text-gray-400 hover:text-gray-600'}`}
                >
                  <Grid size={18} />
                </button>
                <button 
                  onClick={() => setViewMode('list')}
                  className={`p-2 transition-colors ${viewMode === 'list' ? 'bg-gray-100 text-[#006E62]' : 'text-gray-400 hover:text-gray-600'}`}
                >
                  <ListIcon size={18} />
                </button>
              </div>
              <button 
                onClick={() => { setEditingAccount(null); setShowForm(true); }}
                className="flex items-center gap-2 bg-[#006E62] text-white px-4 py-2 rounded-md hover:bg-[#005a50] transition-colors shadow-sm"
              >
                <Plus size={18} />
                <span className="font-medium text-sm">Tambah Akun</span>
              </button>
            </div>
          </div>

          <div className="flex border-b border-gray-100">
            <button
              onClick={() => setStatusFilter('aktif')}
              className={`px-6 py-3 text-[11px] font-bold uppercase tracking-widest border-b-2 transition-all flex items-center gap-2 ${
                statusFilter === 'aktif' ? 'border-[#006E62] text-[#006E62]' : 'border-transparent text-gray-400 hover:text-gray-600'
              }`}
            >
              <UserCheck size={14} />
              Karyawan Aktif <span className={`ml-1 px-1.5 py-0.5 rounded-full text-[9px] ${statusFilter === 'aktif' ? 'bg-[#006E62] text-white' : 'bg-gray-100 text-gray-400'}`}>{activeAccounts.length}</span>
            </button>
            <button
              onClick={() => setStatusFilter('non-aktif')}
              className={`px-6 py-3 text-[11px] font-bold uppercase tracking-widest border-b-2 transition-all flex items-center gap-2 ${
                statusFilter === 'non-aktif' ? 'border-red-500 text-red-600' : 'border-transparent text-gray-400 hover:text-gray-600'
              }`}
            >
              <UserX size={14} />
              Karyawan Non-Aktif <span className={`ml-1 px-1.5 py-0.5 rounded-full text-[9px] ${statusFilter === 'non-aktif' ? 'bg-red-50 text-white' : 'bg-gray-100 text-gray-400'}`}>{inactiveAccounts.length}</span>
            </button>
          </div>

          {isLoading ? (
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
              {[1, 2, 3, 4, 5, 6].map(i => <CardSkeleton key={i} />)}
            </div>
          ) : filteredAccounts.length === 0 ? (
            <div className="flex flex-col items-center justify-center py-20 text-gray-400">
              <Users size={48} strokeWidth={1} className="mb-4" />
              <p className="text-lg">Data akun {statusFilter === 'aktif' ? 'aktif' : 'non-aktif'} tidak ditemukan.</p>
            </div>
          ) : viewMode === 'grid' ? (
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
              {filteredAccounts.map(account => {
                const isInactive = account.end_date && account.end_date <= today;
                const isSelected = selectedIds.includes(account.id);
                return (
                  <div 
                    key={account.id} 
                    className={`group bg-white border p-4 rounded-md shadow-sm hover:shadow-md transition-all cursor-pointer border-l-4 relative ${isSelected ? 'border-[#006E62] bg-emerald-50/10' : 'border-gray-100 border-l-transparent'} ${isInactive && !isSelected ? 'hover:border-l-red-500' : (!isSelected ? 'hover:border-l-[#006E62]' : '')}`}
                  >
                    <div className="absolute top-3 right-3 flex items-center gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                      <button 
                        onClick={(e) => { e.stopPropagation(); toggleSelect(account.id); }}
                        className={`p-1.5 rounded-full border transition-colors ${isSelected ? 'bg-[#006E62] text-white border-[#006E62]' : 'bg-white text-gray-400 border-gray-200 hover:border-[#006E62] hover:text-[#006E62]'}`}
                      >
                        <UserCheck size={12} />
                      </button>
                      <button 
                        onClick={(e) => { e.stopPropagation(); handleDelete(account.id); }}
                        className="p-1.5 rounded-full bg-white text-red-400 border border-gray-200 hover:border-red-500 hover:text-red-500 transition-colors"
                      >
                        <Trash2 size={12} />
                      </button>
                    </div>

                    <div onClick={() => setSelectedAccountId(account.id)}>
                      <div className="flex items-center gap-3 mb-3">
                        <div className="w-10 h-10 rounded-full bg-gray-100 overflow-hidden border border-gray-200 shrink-0 flex items-center justify-center">
                          {account.photo_google_id ? (
                            <img src={googleDriveService.getFileUrl(account.photo_google_id)} alt="" className="w-full h-full object-cover" />
                          ) : (
                            <UserCircle size={24} className="text-gray-400" />
                          )}
                        </div>
                        <div className="min-w-0 flex-1">
                          <h3 className="font-bold text-[#006E62] group-hover:text-[#005a50] line-clamp-1 text-sm">{account.full_name}</h3>
                          <p className="text-[10px] text-gray-400 font-bold uppercase">{account.position} • {account.internal_nik}</p>
                        </div>
                        {isInactive && <span className="text-[8px] font-bold px-1 py-0.5 bg-red-50 text-red-600 rounded uppercase">Exit</span>}
                      </div>
                      <div className="flex items-center justify-between pt-2 border-t border-gray-50">
                         <span className="text-[10px] text-gray-500 font-medium">{(account as any).location?.name || 'Tanpa Lokasi'}</span>
                         <span className={`text-[9px] font-bold px-1.5 py-0.5 rounded uppercase ${
                           isInactive ? 'bg-gray-100 text-gray-400' : (account.employee_type === 'Tetap' ? 'bg-green-50 text-green-600' : 'bg-orange-50 text-orange-600')
                         }`}>
                           {account.employee_type}
                         </span>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          ) : (
            <div className="bg-white border border-gray-100 rounded-md overflow-hidden shadow-sm overflow-x-auto">
              <table className="w-full text-left min-w-[700px]">
                <thead className="bg-gray-50 text-[10px] font-bold text-gray-500 uppercase">
                  <tr>
                    <th className="px-6 py-3 w-10">
                      <input 
                        type="checkbox" 
                        className="rounded border-gray-300 text-[#006E62] focus:ring-[#006E62]"
                        checked={selectedIds.length === filteredAccounts.length && filteredAccounts.length > 0}
                        onChange={toggleSelectAll}
                      />
                    </th>
                    <th className="px-6 py-3">Nama & Posisi</th>
                    <th className="px-6 py-3">NIK Internal</th>
                    <th className="px-6 py-3">Lokasi Penempatan</th>
                    <th className="px-6 py-3">Status</th>
                    <th className="px-6 py-3 text-right">Aksi</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {filteredAccounts.map(account => {
                    const isInactive = account.end_date && account.end_date <= today;
                    const isSelected = selectedIds.includes(account.id);
                    return (
                      <tr 
                        key={account.id} 
                        className={`hover:bg-gray-50 transition-colors ${isSelected ? 'bg-emerald-50/20' : ''}`}
                      >
                        <td className="px-6 py-4">
                          <input 
                            type="checkbox" 
                            className="rounded border-gray-300 text-[#006E62] focus:ring-[#006E62]"
                            checked={isSelected}
                            onChange={() => toggleSelect(account.id)}
                            onClick={(e) => e.stopPropagation()}
                          />
                        </td>
                        <td className="px-6 py-4" onClick={() => setSelectedAccountId(account.id)}>
                          <div className="flex items-center gap-3">
                          <div className="w-8 h-8 rounded-full bg-gray-100 overflow-hidden shrink-0 border border-gray-200 flex items-center justify-center">
                            {account.photo_google_id ? (
                              <img src={googleDriveService.getFileUrl(account.photo_google_id)} className="w-full h-full object-cover" />
                            ) : (
                              <UserCircle size={20} className="text-gray-400" />
                            )}
                          </div>
                            <div>
                              <div className="font-bold text-[#006E62] text-xs">{account.full_name}</div>
                              <div className="text-[9px] text-gray-400 uppercase font-bold">{account.position}</div>
                            </div>
                          </div>
                        </td>
                        <td className="px-6 py-4 text-xs font-mono text-gray-500" onClick={() => setSelectedAccountId(account.id)}>{account.internal_nik}</td>
                        <td className="px-6 py-4 text-xs text-gray-500" onClick={() => setSelectedAccountId(account.id)}>{(account as any).location?.name || '-'}</td>
                        <td className="px-6 py-4" onClick={() => setSelectedAccountId(account.id)}>
                          <span className={`text-[9px] font-bold px-1.5 py-0.5 border border-gray-100 rounded uppercase ${isInactive ? 'bg-red-50 text-red-600' : 'bg-gray-50 text-gray-500'}`}>
                            {isInactive ? 'NON-AKTIF' : account.employee_type}
                          </span>
                        </td>
                        <td className="px-6 py-4 text-right">
                          <button 
                            onClick={(e) => { e.stopPropagation(); handleDelete(account.id); }}
                            className="p-1.5 text-red-400 hover:text-red-600 transition-colors"
                            title="Hapus Akun"
                          >
                            <Trash2 size={14} />
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      ) : activeSubTab === 'career' ? (
        <div className="animate-in fade-in duration-300">
          <CareerLogMain />
        </div>
      ) : activeSubTab === 'contract' ? (
        <div className="animate-in fade-in duration-300">
          <ContractMain />
        </div>
      ) : activeSubTab === 'cert' ? (
        <div className="animate-in fade-in duration-300">
          <CertificationMain />
        </div>
      ) : activeSubTab === 'health' ? (
        <div className="animate-in fade-in duration-300">
          <HealthLogMain />
        </div>
      ) : activeSubTab === 'discipline' ? (
        <div className="animate-in fade-in duration-300">
          <DisciplineMain />
        </div>
      ) : null}

      {showForm && (
        <AccountForm 
          onClose={() => { setShowForm(false); setEditingAccount(null); }}
          onSubmit={editingAccount ? (data) => handleUpdate(editingAccount.id, data) : handleCreate}
          initialData={editingAccount || undefined}
        />
      )}
    </div>
  );
};

export default AccountMain;
