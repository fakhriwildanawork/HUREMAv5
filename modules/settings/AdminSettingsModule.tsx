import React, { useState, useEffect } from 'react';
import { Shield, Users, Search, Save, Check, X, UserCheck, ShieldAlert, Wallet, Target } from 'lucide-react';
import { accountService } from '../../services/accountService';
import { settingsService } from '../../services/settingsService';
import { Account } from '../../types';
import Swal from 'sweetalert2';

const AdminSettingsModule: React.FC = () => {
  const [accounts, setAccounts] = useState<Account[]>([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  
  // Admin Role States (Arrays of Account IDs)
  const [hrAdmins, setHrAdmins] = useState<string[]>([]);
  const [performanceAdmins, setPerformanceAdmins] = useState<string[]>([]);
  const [financeAdmins, setFinanceAdmins] = useState<string[]>([]);

  useEffect(() => {
    fetchData();
  }, []);

  const fetchData = async () => {
    setLoading(true);
    try {
      const [accData, hrData, perfData, finData] = await Promise.all([
        accountService.getAll(),
        settingsService.getSetting('admin_hr_ids', []),
        settingsService.getSetting('admin_performance_ids', []),
        settingsService.getSetting('admin_finance_ids', [])
      ]);
      
      setAccounts(accData);
      setHrAdmins(hrData);
      setPerformanceAdmins(perfData);
      setFinanceAdmins(finData);
    } catch (error) {
      console.error('Error fetching admin settings:', error);
      Swal.fire('Error', 'Gagal memuat data pengaturan admin', 'error');
    } finally {
      setLoading(false);
    }
  };

  const handleSave = async () => {
    setSaving(true);
    try {
      await Promise.all([
        settingsService.updateSetting('admin_hr_ids', hrAdmins, 'Daftar ID Pegawai dengan akses Admin Kepegawaian'),
        settingsService.updateSetting('admin_performance_ids', performanceAdmins, 'Daftar ID Pegawai dengan akses Admin Performa'),
        settingsService.updateSetting('admin_finance_ids', financeAdmins, 'Daftar ID Pegawai dengan akses Admin Finance')
      ]);
      
      Swal.fire({
        title: 'Berhasil',
        text: 'Pengaturan admin telah diperbarui',
        icon: 'success',
        timer: 2000,
        showConfirmButton: false
      });
    } catch (error) {
      console.error('Error saving admin settings:', error);
      Swal.fire('Error', 'Gagal menyimpan pengaturan admin', 'error');
    } finally {
      setSaving(false);
    }
  };

  const toggleAdmin = (accountId: string, role: 'hr' | 'performance' | 'finance') => {
    if (role === 'hr') {
      setHrAdmins(prev => 
        prev.includes(accountId) ? prev.filter(id => id !== accountId) : [...prev, accountId]
      );
    } else if (role === 'performance') {
      setPerformanceAdmins(prev => 
        prev.includes(accountId) ? prev.filter(id => id !== accountId) : [...prev, accountId]
      );
    } else if (role === 'finance') {
      setFinanceAdmins(prev => 
        prev.includes(accountId) ? prev.filter(id => id !== accountId) : [...prev, accountId]
      );
    }
  };

  const getSortedAccounts = (adminIds: string[]) => {
    return accounts
      .filter(acc => acc.full_name !== 'Superadmin')
      .filter(acc => 
        acc.full_name.toLowerCase().includes(searchTerm.toLowerCase()) ||
        acc.internal_nik.toLowerCase().includes(searchTerm.toLowerCase())
      )
      .sort((a, b) => {
        const aSelected = adminIds.includes(a.id);
        const bSelected = adminIds.includes(b.id);
        if (aSelected && !bSelected) return -1;
        if (!aSelected && bSelected) return 1;
        return 0;
      });
  };

  if (loading) {
    return (
      <div className="flex flex-col items-center justify-center h-64">
        <div className="w-10 h-10 border-4 border-[#006E62] border-t-transparent rounded-full animate-spin"></div>
        <p className="mt-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Memuat Pengaturan...</p>
      </div>
    );
  }

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-2xl font-black text-gray-800 tracking-tight uppercase">Pengaturan Admin</h2>
          <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Kelola hak akses admin khusus untuk pegawai</p>
        </div>
        <button
          onClick={handleSave}
          disabled={saving}
          className="flex items-center gap-2 bg-[#006E62] text-white px-6 py-3 rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-[#005a50] transition-all shadow-lg shadow-[#006E62]/20 active:scale-95 disabled:opacity-50"
        >
          {saving ? <div className="w-3 h-3 border-2 border-white border-t-transparent rounded-full animate-spin"></div> : <Save size={14} />}
          Simpan Pengaturan
        </button>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Admin Kepegawaian */}
        <div className="bg-white rounded-2xl border border-gray-100 shadow-sm overflow-hidden flex flex-col">
          <div className="p-6 border-b border-gray-50 bg-emerald-50/50">
            <div className="flex items-center gap-3 mb-2">
              <div className="p-2 bg-emerald-100 text-emerald-600 rounded-lg">
                <UserCheck size={20} />
              </div>
              <h3 className="font-bold text-gray-800 text-sm uppercase tracking-wider">Admin Kepegawaian</h3>
            </div>
            <p className="text-[10px] text-gray-500 leading-relaxed">Akses untuk mengelola data karyawan, absensi, dan dokumen digital.</p>
            <div className="mt-4 flex items-center justify-between">
              <span className="text-[10px] font-black text-emerald-600 bg-emerald-100 px-2 py-1 rounded-md uppercase">
                {hrAdmins.length} Terpilih
              </span>
            </div>
          </div>
          <div className="p-4 border-b border-gray-50">
            <div className="relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={14} />
              <input
                type="text"
                placeholder="Cari Pegawai..."
                className="w-full pl-9 pr-4 py-2 bg-gray-50 border-none rounded-lg text-xs focus:ring-2 focus:ring-emerald-500/20"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
            </div>
          </div>
          <div className="flex-1 overflow-y-auto max-h-[400px] p-2 space-y-1 scrollbar-thin">
            {getSortedAccounts(hrAdmins).map(acc => (
              <button
                key={acc.id}
                onClick={() => toggleAdmin(acc.id, 'hr')}
                className={`w-full flex items-center justify-between p-3 rounded-xl transition-all ${
                  hrAdmins.includes(acc.id) ? 'bg-emerald-50 border-emerald-100' : 'hover:bg-gray-50 border-transparent'
                } border`}
              >
                <div className="flex items-center gap-3 text-left">
                  <div className={`w-8 h-8 rounded-full flex items-center justify-center text-xs font-bold ${
                    hrAdmins.includes(acc.id) ? 'bg-emerald-600 text-white' : 'bg-gray-100 text-gray-400'
                  }`}>
                    {acc.full_name.charAt(0)}
                  </div>
                  <div>
                    <p className="text-xs font-bold text-gray-700">{acc.full_name}</p>
                    <p className="text-[10px] text-gray-400">{acc.internal_nik}</p>
                  </div>
                </div>
                {hrAdmins.includes(acc.id) && <Check size={16} className="text-emerald-600" />}
              </button>
            ))}
          </div>
        </div>

        {/* Admin Performa */}
        <div className="bg-white rounded-2xl border border-gray-100 shadow-sm overflow-hidden flex flex-col">
          <div className="p-6 border-b border-gray-50 bg-blue-50/50">
            <div className="flex items-center gap-3 mb-2">
              <div className="p-2 bg-blue-100 text-blue-600 rounded-lg">
                <Target size={20} />
              </div>
              <h3 className="font-bold text-gray-800 text-sm uppercase tracking-wider">Admin Performa</h3>
            </div>
            <p className="text-[10px] text-gray-500 leading-relaxed">Akses untuk mengelola KPI, Key Activities, dan Laporan Sales.</p>
            <div className="mt-4 flex items-center justify-between">
              <span className="text-[10px] font-black text-blue-600 bg-blue-100 px-2 py-1 rounded-md uppercase">
                {performanceAdmins.length} Terpilih
              </span>
            </div>
          </div>
          <div className="p-4 border-b border-gray-50">
            <div className="relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={14} />
              <input
                type="text"
                placeholder="Cari Pegawai..."
                className="w-full pl-9 pr-4 py-2 bg-gray-50 border-none rounded-lg text-xs focus:ring-2 focus:ring-blue-500/20"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
            </div>
          </div>
          <div className="flex-1 overflow-y-auto max-h-[400px] p-2 space-y-1 scrollbar-thin">
            {getSortedAccounts(performanceAdmins).map(acc => (
              <button
                key={acc.id}
                onClick={() => toggleAdmin(acc.id, 'performance')}
                className={`w-full flex items-center justify-between p-3 rounded-xl transition-all ${
                  performanceAdmins.includes(acc.id) ? 'bg-blue-50 border-blue-100' : 'hover:bg-gray-50 border-transparent'
                } border`}
              >
                <div className="flex items-center gap-3 text-left">
                  <div className={`w-8 h-8 rounded-full flex items-center justify-center text-xs font-bold ${
                    performanceAdmins.includes(acc.id) ? 'bg-blue-600 text-white' : 'bg-gray-100 text-gray-400'
                  }`}>
                    {acc.full_name.charAt(0)}
                  </div>
                  <div>
                    <p className="text-xs font-bold text-gray-700">{acc.full_name}</p>
                    <p className="text-[10px] text-gray-400">{acc.internal_nik}</p>
                  </div>
                </div>
                {performanceAdmins.includes(acc.id) && <Check size={16} className="text-blue-600" />}
              </button>
            ))}
          </div>
        </div>

        {/* Admin Finance */}
        <div className="bg-white rounded-2xl border border-gray-100 shadow-sm overflow-hidden flex flex-col">
          <div className="p-6 border-b border-gray-50 bg-indigo-50/50">
            <div className="flex items-center gap-3 mb-2">
              <div className="p-2 bg-indigo-100 text-indigo-600 rounded-lg">
                <Wallet size={20} />
              </div>
              <h3 className="font-bold text-gray-800 text-sm uppercase tracking-wider">Admin Finance</h3>
            </div>
            <p className="text-[10px] text-gray-500 leading-relaxed">Akses untuk mengelola Payroll, Reimburse, dan Skema Gaji.</p>
            <div className="mt-4 flex items-center justify-between">
              <span className="text-[10px] font-black text-indigo-600 bg-indigo-100 px-2 py-1 rounded-md uppercase">
                {financeAdmins.length} Terpilih
              </span>
            </div>
          </div>
          <div className="p-4 border-b border-gray-50">
            <div className="relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={14} />
              <input
                type="text"
                placeholder="Cari Pegawai..."
                className="w-full pl-9 pr-4 py-2 bg-gray-50 border-none rounded-lg text-xs focus:ring-2 focus:ring-indigo-500/20"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
            </div>
          </div>
          <div className="flex-1 overflow-y-auto max-h-[400px] p-2 space-y-1 scrollbar-thin">
            {getSortedAccounts(financeAdmins).map(acc => (
              <button
                key={acc.id}
                onClick={() => toggleAdmin(acc.id, 'finance')}
                className={`w-full flex items-center justify-between p-3 rounded-xl transition-all ${
                  financeAdmins.includes(acc.id) ? 'bg-indigo-50 border-indigo-100' : 'hover:bg-gray-50 border-transparent'
                } border`}
              >
                <div className="flex items-center gap-3 text-left">
                  <div className={`w-8 h-8 rounded-full flex items-center justify-center text-xs font-bold ${
                    financeAdmins.includes(acc.id) ? 'bg-indigo-600 text-white' : 'bg-gray-100 text-gray-400'
                  }`}>
                    {acc.full_name.charAt(0)}
                  </div>
                  <div>
                    <p className="text-xs font-bold text-gray-700">{acc.full_name}</p>
                    <p className="text-[10px] text-gray-400">{acc.internal_nik}</p>
                  </div>
                </div>
                {financeAdmins.includes(acc.id) && <Check size={16} className="text-indigo-600" />}
              </button>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};

export default AdminSettingsModule;
