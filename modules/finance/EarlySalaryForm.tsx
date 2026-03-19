import React, { useState, useEffect } from 'react';
import { ArrowLeft, Save, AlertCircle, Wallet, Calendar } from 'lucide-react';
import { financeService } from '../../services/financeService';
import { authService } from '../../services/authService';
import { EarlySalaryRequestInput } from '../../types';
import Swal from 'sweetalert2';

interface EarlySalaryFormProps {
  onBack: () => void;
}

const EarlySalaryForm: React.FC<EarlySalaryFormProps> = ({ onBack }) => {
  const [loading, setLoading] = useState(false);
  const [basicSalary, setBasicSalary] = useState<number>(0);
  const [formData, setFormData] = useState<EarlySalaryRequestInput>({
    month: new Date().getMonth() + 1,
    year: new Date().getFullYear(),
    amount: 0,
    reason: ''
  });

  const user = authService.getCurrentUser();

  useEffect(() => {
    const fetchBasicSalary = async () => {
      if (!user) return;
      try {
        // Fetch assignments to get scheme and basic salary
        const assignments = await financeService.getAssignments();
        const userAssignment = assignments.find(a => a.account_id === user.id);
        if (userAssignment && userAssignment.scheme) {
          setBasicSalary(userAssignment.scheme.basic_salary);
        }
      } catch (error) {
        console.error('Error fetching basic salary:', error);
      }
    };
    fetchBasicSalary();
  }, [user]);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!user) return;

    if (formData.amount <= 0) {
      Swal.fire('Error', 'Nominal pengajuan harus lebih dari 0.', 'error');
      return;
    }

    if (formData.amount > basicSalary) {
      Swal.fire('Error', `Nominal pengajuan maksimal adalah Rp ${basicSalary.toLocaleString('id-ID')} (Gaji Pokok).`, 'error');
      return;
    }

    if (!formData.reason.trim()) {
      Swal.fire('Error', 'Alasan pengajuan harus diisi.', 'error');
      return;
    }

    setLoading(true);
    try {
      // Check if payroll for this period is already processed
      const payrollStatus = await financeService.getPayrollStatus(formData.month, formData.year);
      if (payrollStatus && (payrollStatus.status === 'Approved' || payrollStatus.status === 'Paid')) {
        Swal.fire('Gagal', 'Tidak dapat mengajukan gaji awal pada periode yang sudah diverifikasi/dibayar.', 'error');
        setLoading(false);
        return;
      }

      await financeService.createEarlySalaryRequest({
        ...formData,
        account_id: user.id
      });

      await Swal.fire({
        title: 'Berhasil!',
        text: 'Pengajuan gaji awal telah dikirim.',
        icon: 'success',
        confirmButtonColor: '#006E62'
      });
      onBack();
    } catch (error) {
      console.error('Error creating early salary request:', error);
      Swal.fire('Error', 'Gagal mengirim pengajuan.', 'error');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="max-w-2xl mx-auto">
      <div className="flex items-center gap-4 mb-8">
        <button 
          onClick={onBack}
          className="p-2 hover:bg-gray-100 rounded-full transition-colors"
        >
          <ArrowLeft size={20} className="text-gray-600" />
        </button>
        <div>
          <h2 className="text-2xl font-bold text-gray-800">Ajukan Gaji Awal</h2>
          <p className="text-sm text-gray-500">Isi formulir di bawah untuk mengajukan kasbon gaji.</p>
        </div>
      </div>

      <form onSubmit={handleSubmit} className="space-y-6">
        <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6 space-y-6">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="space-y-2">
              <label className="text-xs font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
                <Calendar size={14} />
                Bulan Periode
              </label>
              <select
                required
                value={formData.month}
                onChange={(e) => setFormData({ ...formData, month: Number(e.target.value) })}
                className="w-full bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all"
              >
                {Array.from({ length: 12 }, (_, i) => (
                  <option key={i + 1} value={i + 1}>
                    {new Date(0, i).toLocaleString('id-ID', { month: 'long' })}
                  </option>
                ))}
              </select>
            </div>

            <div className="space-y-2">
              <label className="text-xs font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
                <Calendar size={14} />
                Tahun
              </label>
              <select
                required
                value={formData.year}
                onChange={(e) => setFormData({ ...formData, year: Number(e.target.value) })}
                className="w-full bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all"
              >
                {Array.from({ length: 3 }, (_, i) => (
                  <option key={i} value={new Date().getFullYear() + i}>
                    {new Date().getFullYear() + i}
                  </option>
                ))}
              </select>
            </div>
          </div>

          <div className="space-y-2">
            <label className="text-xs font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
              <Wallet size={14} />
              Nominal Pengajuan
            </label>
            <div className="relative">
              <span className="absolute left-4 top-1/2 -translate-y-1/2 text-gray-400 font-bold text-sm">Rp</span>
              <input
                type="number"
                required
                min="0"
                max={basicSalary}
                value={formData.amount}
                onChange={(e) => setFormData({ ...formData, amount: Number(e.target.value) })}
                className="w-full bg-gray-50 border-none rounded-lg pl-12 pr-4 py-3 text-sm font-mono font-bold focus:ring-2 focus:ring-[#006E62] transition-all"
                placeholder="0"
              />
            </div>
            <p className="text-[10px] text-gray-400 italic">
              * Maksimal pengajuan: Rp {basicSalary.toLocaleString('id-ID')} (Gaji Pokok Anda)
            </p>
          </div>

          <div className="space-y-2">
            <label className="text-xs font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
              <AlertCircle size={14} />
              Alasan Pengajuan
            </label>
            <textarea
              required
              rows={4}
              value={formData.reason}
              onChange={(e) => setFormData({ ...formData, reason: e.target.value })}
              className="w-full bg-gray-50 border-none rounded-lg p-4 text-sm focus:ring-2 focus:ring-[#006E62] transition-all resize-none"
              placeholder="Jelaskan alasan Anda mengajukan gaji awal..."
            />
          </div>
        </div>

        <div className="flex justify-end gap-3">
          <button
            type="button"
            onClick={onBack}
            className="px-6 py-2.5 text-sm font-bold text-gray-500 hover:text-gray-700 transition-colors"
          >
            BATAL
          </button>
          <button
            type="submit"
            disabled={loading}
            className="flex items-center gap-2 px-8 py-2.5 bg-[#006E62] text-white rounded-lg hover:bg-[#005a50] disabled:opacity-50 disabled:cursor-not-allowed transition-all text-sm font-bold shadow-lg shadow-emerald-100"
          >
            {loading ? (
              <div className="w-5 h-5 border-2 border-white border-t-transparent rounded-full animate-spin"></div>
            ) : (
              <>
                <Save size={18} />
                KIRIM PENGAJUAN
              </>
            )}
          </button>
        </div>
      </form>
    </div>
  );
};

export default EarlySalaryForm;
