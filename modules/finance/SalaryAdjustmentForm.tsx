import React, { useState, useEffect } from 'react';
import { ArrowLeft, Save, Users, DollarSign, Calendar, Info, Search, Check } from 'lucide-react';
import { financeService } from '../../services/financeService';
import { accountService } from '../../services/accountService';
import { Account, SalaryAdjustmentInput } from '../../types';
import Swal from 'sweetalert2';

interface SalaryAdjustmentFormProps {
  adjustment?: any | null;
  onBack: () => void;
  onSuccess: () => void;
}

const formatNumber = (num: number) => {
  return num.toLocaleString('id-ID');
};

const parseNumber = (str: string) => {
  if (typeof str !== 'string') return 0;
  return Number(str.replace(/\./g, '')) || 0;
};

const SalaryAdjustmentForm: React.FC<SalaryAdjustmentFormProps> = ({ adjustment, onBack, onSuccess }) => {
  const [loading, setLoading] = useState(false);
  const [accounts, setAccounts] = useState<Account[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedAccountIds, setSelectedAccountIds] = useState<string[]>(adjustment ? [adjustment.account_id] : []);
  
  const [formData, setFormData] = useState({
    type: adjustment?.type || 'Addition' as 'Addition' | 'Deduction',
    amount: adjustment?.amount || 0,
    month: adjustment?.month || new Date().getMonth() + 1,
    year: adjustment?.year || new Date().getFullYear(),
    description: adjustment?.description || ''
  });

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

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (selectedAccountIds.length === 0) {
      Swal.fire('Peringatan', 'Silakan pilih minimal satu karyawan.', 'warning');
      return;
    }

    setLoading(true);
    try {
      // Check payroll status first
      const payrollStatus = await financeService.getPayrollStatus(formData.month, formData.year);
      if (payrollStatus && payrollStatus.status !== 'Draft') {
        Swal.fire('Gagal!', `Periode gaji ${formData.month}/${formData.year} sudah di-ACC (${payrollStatus.status}). Data tidak dapat diubah.`, 'error');
        setLoading(false);
        return;
      }

      if (adjustment) {
        await financeService.updateSalaryAdjustment(adjustment.id, {
          ...formData,
          account_id: selectedAccountIds[0]
        });
        Swal.fire('Berhasil!', 'Penyesuaian gaji berhasil diperbarui.', 'success');
      } else {
        const adjustments: SalaryAdjustmentInput[] = selectedAccountIds.map(accountId => ({
          ...formData,
          account_id: accountId
        }));
        await financeService.createSalaryAdjustment(adjustments);
        Swal.fire('Berhasil!', `${adjustments.length} penyesuaian gaji berhasil dibuat.`, 'success');
      }
      onSuccess();
    } catch (error) {
      console.error('Error saving adjustment:', error);
      Swal.fire('Gagal!', 'Terjadi kesalahan saat menyimpan data.', 'error');
    } finally {
      setLoading(false);
    }
  };

  const filteredAccounts = accounts.filter(acc => 
    acc.full_name.toLowerCase().includes(searchQuery.toLowerCase()) ||
    acc.internal_nik.toLowerCase().includes(searchQuery.toLowerCase())
  );

  const toggleAccount = (id: string) => {
    if (adjustment) {
      setSelectedAccountIds([id]);
    } else {
      setSelectedAccountIds(prev => 
        prev.includes(id) ? prev.filter(i => i !== id) : [...prev, id]
      );
    }
  };

  return (
    <div className="max-w-4xl mx-auto space-y-6">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-4">
          <button
            onClick={onBack}
            className="p-2 hover:bg-gray-100 rounded-full transition-colors text-gray-500"
          >
            <ArrowLeft size={20} />
          </button>
          <div>
            <h2 className="text-2xl font-bold text-gray-800">{adjustment ? 'Edit Kustom Gaji' : 'Buat Kustom Gaji Baru'}</h2>
            <p className="text-sm text-gray-500">Tambahkan pendapatan atau potongan insidental untuk karyawan.</p>
          </div>
        </div>
      </div>

      <form onSubmit={handleSubmit} className="grid grid-cols-1 md:grid-cols-3 gap-6">
        {/* Left Column: Form Fields */}
        <div className="md:col-span-2 space-y-6">
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 space-y-6">
            <div className="flex items-center gap-2 text-[#006E62] mb-2">
              <Info size={18} />
              <h3 className="font-bold text-sm uppercase tracking-wider">Detail Penyesuaian</h3>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Jenis Penyesuaian</label>
                <div className="flex items-center gap-2 p-1 bg-gray-100 rounded-lg">
                  <button
                    type="button"
                    onClick={() => setFormData({ ...formData, type: 'Addition' })}
                    className={`flex-1 py-2 text-xs font-bold rounded-md transition-all ${
                      formData.type === 'Addition' ? 'bg-emerald-600 text-white shadow-sm' : 'text-gray-500 hover:text-gray-700'
                    }`}
                  >
                    TAMBAHAN
                  </button>
                  <button
                    type="button"
                    onClick={() => setFormData({ ...formData, type: 'Deduction' })}
                    className={`flex-1 py-2 text-xs font-bold rounded-md transition-all ${
                      formData.type === 'Deduction' ? 'bg-rose-600 text-white shadow-sm' : 'text-gray-500 hover:text-gray-700'
                    }`}
                  >
                    POTONGAN
                  </button>
                </div>
              </div>

              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
                  <DollarSign size={12} />
                  Nominal
                </label>
                <div className="relative">
                  <div className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400 text-sm font-medium">Rp</div>
                  <input
                    type="text"
                    value={formatNumber(formData.amount)}
                    onChange={(e) => setFormData({ ...formData, amount: parseNumber(e.target.value) })}
                    className="w-full pl-10 pr-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium"
                    required
                  />
                </div>
              </div>

              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
                  <Calendar size={12} />
                  Bulan
                </label>
                <select
                  value={formData.month}
                  onChange={(e) => setFormData({ ...formData, month: Number(e.target.value) })}
                  className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium"
                >
                  {Array.from({ length: 12 }, (_, i) => (
                    <option key={i + 1} value={i + 1}>
                      {new Date(0, i).toLocaleString('id-ID', { month: 'long' })}
                    </option>
                  ))}
                </select>
              </div>

              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
                  <Calendar size={12} />
                  Tahun
                </label>
                <select
                  value={formData.year}
                  onChange={(e) => setFormData({ ...formData, year: Number(e.target.value) })}
                  className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium"
                >
                  {Array.from({ length: 5 }, (_, i) => (
                    <option key={i} value={new Date().getFullYear() - 2 + i}>
                      {new Date().getFullYear() - 2 + i}
                    </option>
                  ))}
                </select>
              </div>

              <div className="md:col-span-2 space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Keterangan</label>
                <textarea
                  value={formData.description}
                  onChange={(e) => setFormData({ ...formData, description: e.target.value })}
                  placeholder="Contoh: Bonus Lembur Proyek A atau Denda Keterlambatan Meeting"
                  className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium min-h-[100px]"
                  required
                />
              </div>
            </div>
          </div>

          <div className="flex justify-end gap-3">
            <button
              type="button"
              onClick={onBack}
              className="px-6 py-2.5 border border-gray-200 text-gray-600 rounded-lg hover:bg-gray-50 transition-all text-sm font-bold uppercase tracking-wider"
            >
              Batal
            </button>
            <button
              type="submit"
              disabled={loading}
              className="flex items-center gap-2 px-8 py-2.5 bg-[#006E62] text-white rounded-lg hover:bg-[#005a50] transition-all text-sm font-bold uppercase tracking-wider shadow-lg disabled:opacity-50"
            >
              {loading ? (
                <div className="w-5 h-5 border-2 border-white border-t-transparent rounded-full animate-spin"></div>
              ) : (
                <Save size={18} />
              )}
              {adjustment ? 'Simpan Perubahan' : 'Simpan Penyesuaian'}
            </button>
          </div>
        </div>

        {/* Right Column: Employee Selection */}
        <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 flex flex-col h-[600px]">
          <div className="flex items-center gap-2 text-[#006E62] mb-4">
            <Users size={18} />
            <h3 className="font-bold text-sm uppercase tracking-wider">Pilih Karyawan</h3>
          </div>

          <div className="relative mb-4">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
            <input
              type="text"
              placeholder="Cari nama/NIK..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full pl-10 pr-4 py-2 bg-gray-50 border-none rounded-lg text-xs focus:ring-2 focus:ring-[#006E62] transition-all"
            />
          </div>

          <div className="flex-1 overflow-y-auto space-y-2 pr-2 custom-scrollbar">
            {filteredAccounts.map((acc) => (
              <button
                key={acc.id}
                type="button"
                onClick={() => toggleAccount(acc.id)}
                className={`w-full flex items-center justify-between p-3 rounded-lg border transition-all text-left ${
                  selectedAccountIds.includes(acc.id)
                    ? 'border-[#006E62] bg-emerald-50'
                    : 'border-gray-100 hover:border-gray-200'
                }`}
              >
                <div>
                  <div className="text-sm font-bold text-gray-800">{acc.full_name}</div>
                  <div className="text-[10px] text-gray-500 uppercase tracking-wider font-medium">{acc.internal_nik} • {acc.position}</div>
                </div>
                {selectedAccountIds.includes(acc.id) && (
                  <div className="bg-[#006E62] text-white p-1 rounded-full">
                    <Check size={12} />
                  </div>
                )}
              </button>
            ))}
          </div>

          <div className="mt-4 pt-4 border-t border-gray-100">
            <div className="flex justify-between items-center text-xs">
              <span className="text-gray-500 font-medium">Terpilih:</span>
              <span className="font-bold text-[#006E62]">{selectedAccountIds.length} Karyawan</span>
            </div>
          </div>
        </div>
      </form>
    </div>
  );
};

export default SalaryAdjustmentForm;
