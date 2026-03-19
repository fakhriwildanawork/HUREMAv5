
import React, { useState, useEffect } from 'react';
import { X, Save, Search, User, Calendar, FileText, Loader2, CheckCircle2 } from 'lucide-react';
import { accountService } from '../../../services/accountService';
import { awardService } from '../../../services/awardService';
import { Account, EmployeeOfThePeriodInput } from '../../../types';
import Swal from 'sweetalert2';

interface EmployeeOfThePeriodFormProps {
  onClose: () => void;
  onSuccess: () => void;
}

const EmployeeOfThePeriodForm: React.FC<EmployeeOfThePeriodFormProps> = ({ onClose, onSuccess }) => {
  const [accounts, setAccounts] = useState<Account[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedAccountIds, setSelectedAccountIds] = useState<string[]>([]);
  const [formData, setFormData] = useState({
    month: new Date().getMonth() + 1,
    year: new Date().getFullYear(),
    reason: ''
  });
  const [isSaving, setIsSaving] = useState(false);

  useEffect(() => {
    const fetchAccounts = async () => {
      try {
        const data = await accountService.getAll();
        setAccounts(data);
      } catch (error) {
        console.error('Error fetching accounts:', error);
      }
    };
    fetchAccounts();
  }, []);

  const filteredAccounts = accounts.filter(acc => 
    acc.full_name.toLowerCase().includes(searchQuery.toLowerCase()) ||
    acc.internal_nik.toLowerCase().includes(searchQuery.toLowerCase())
  );

  const toggleAccount = (id: string) => {
    setSelectedAccountIds(prev => 
      prev.includes(id) ? prev.filter(i => i !== id) : [...prev, id]
    );
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (selectedAccountIds.length === 0) {
      Swal.fire('Peringatan', 'Pilih minimal satu pegawai.', 'warning');
      return;
    }

    try {
      setIsSaving(true);
      const input: EmployeeOfThePeriodInput = {
        account_ids: selectedAccountIds,
        month: formData.month,
        year: formData.year,
        reason: formData.reason
      };
      await awardService.createEmployeeOfThePeriod(input);
      Swal.fire({ title: 'Berhasil!', text: 'Penghargaan telah ditambahkan.', icon: 'success', timer: 1500, showConfirmButton: false });
      onSuccess();
    } catch (error) {
      console.error('Error saving award:', error);
      Swal.fire('Gagal', 'Terjadi kesalahan saat menyimpan data.', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const months = [
    { value: 1, label: 'Januari' }, { value: 2, label: 'Februari' }, { value: 3, label: 'Maret' },
    { value: 4, label: 'April' }, { value: 5, label: 'Mei' }, { value: 6, label: 'Juni' },
    { value: 7, label: 'Juli' }, { value: 8, label: 'Agustus' }, { value: 9, label: 'September' },
    { value: 10, label: 'Oktober' }, { value: 11, label: 'November' }, { value: 12, label: 'Desember' }
  ];

  return (
    <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black bg-opacity-40 p-4">
      <div className="bg-white rounded-xl shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in duration-200">
        <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
          <div>
            <h3 className="text-base font-bold text-gray-800">Kelola Employee of The Period</h3>
            <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Pilih Bintang HUREMA Bulan Ini</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors"><X size={20} /></button>
        </div>

        <form onSubmit={handleSubmit} className="p-6 space-y-6 max-h-[85vh] overflow-y-auto scrollbar-none">
          <div className="grid grid-cols-2 gap-4">
            <div className="space-y-1">
              <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Bulan</label>
              <div className="relative">
                <Calendar className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={14} />
                <select 
                  value={formData.month}
                  onChange={(e) => setFormData({...formData, month: Number(e.target.value)})}
                  className="w-full pl-9 pr-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-amber-500 outline-none bg-white"
                >
                  {months.map(m => <option key={m.value} value={m.value}>{m.label}</option>)}
                </select>
              </div>
            </div>
            <div className="space-y-1">
              <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Tahun</label>
              <div className="relative">
                <Calendar className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={14} />
                <input 
                  type="number"
                  value={formData.year}
                  onChange={(e) => setFormData({...formData, year: Number(e.target.value)})}
                  className="w-full pl-9 pr-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-amber-500 outline-none"
                />
              </div>
            </div>
          </div>

          <div className="space-y-2">
            <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Pilih Pegawai ({selectedAccountIds.length} terpilih)</label>
            <div className="relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={14} />
              <input 
                type="text"
                placeholder="Cari nama atau NIK..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full pl-9 pr-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-amber-500 outline-none"
              />
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 max-h-48 overflow-y-auto p-1 scrollbar-thin scrollbar-thumb-gray-200">
              {filteredAccounts.map(acc => (
                <div 
                  key={acc.id}
                  onClick={() => toggleAccount(acc.id)}
                  className={`flex items-center gap-3 p-2 rounded-lg border cursor-pointer transition-all ${
                    selectedAccountIds.includes(acc.id) 
                      ? 'border-amber-500 bg-amber-50' 
                      : 'border-gray-100 hover:border-gray-300'
                  }`}
                >
                  <div className="w-8 h-8 rounded-full bg-gray-100 flex items-center justify-center text-gray-400 overflow-hidden">
                    <User size={16} />
                  </div>
                  <div className="flex-1 min-w-0">
                    <p className="text-[10px] font-bold text-gray-800 truncate">{acc.full_name}</p>
                    <p className="text-[8px] text-gray-400 uppercase">{acc.internal_nik}</p>
                  </div>
                  {selectedAccountIds.includes(acc.id) && <CheckCircle2 size={14} className="text-amber-500" />}
                </div>
              ))}
            </div>
          </div>

          <div className="space-y-1">
            <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Alasan / Pesan Apresiasi</label>
            <div className="relative">
              <FileText className="absolute left-3 top-3 text-gray-400" size={14} />
              <textarea 
                rows={4}
                value={formData.reason}
                onChange={(e) => setFormData({...formData, reason: e.target.value})}
                placeholder="Tuliskan alasan mengapa pegawai ini terpilih..."
                className="w-full pl-9 pr-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-amber-500 outline-none resize-none"
              />
            </div>
          </div>

          <div className="flex justify-end gap-3 pt-4">
            <button type="button" onClick={onClose} className="px-4 py-2 text-xs font-bold text-gray-500 uppercase">Batal</button>
            <button 
              disabled={isSaving}
              className="flex items-center gap-2 bg-amber-500 text-white px-8 py-2 rounded shadow-md hover:bg-amber-600 transition-all text-xs font-bold uppercase disabled:opacity-50"
            >
              {isSaving ? <Loader2 size={14} className="animate-spin" /> : <Save size={14} />} Simpan Penghargaan
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};

export default EmployeeOfThePeriodForm;
