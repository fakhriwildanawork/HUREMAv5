import React, { useState, useEffect } from 'react';
import { Plus, Search, Filter, Trash2, Edit2, DollarSign, Calendar, Info, Eye, AlertCircle, Check } from 'lucide-react';
import { financeService } from '../../services/financeService';
import { SalaryAdjustment, PayrollStatus } from '../../types';
import SalaryAdjustmentForm from './SalaryAdjustmentForm';
import Swal from 'sweetalert2';

const SalaryAdjustmentMain: React.FC = () => {
  const [adjustments, setAdjustments] = useState<SalaryAdjustment[]>([]);
  const [loading, setLoading] = useState(true);
  const [searchQuery, setSearchQuery] = useState('');
  const [filterMonth, setFilterMonth] = useState(new Date().getMonth() + 1);
  const [filterYear, setFilterYear] = useState(new Date().getFullYear());
  const [activeView, setActiveView] = useState<'list' | 'form' | 'detail'>('list');
  const [selectedAdjustment, setSelectedAdjustment] = useState<SalaryAdjustment | null>(null);
  const [payrollStatus, setPayrollStatus] = useState<PayrollStatus | null>(null);

  const fetchAdjustments = async () => {
    setLoading(true);
    try {
      const [aData, pStatus] = await Promise.all([
        financeService.getSalaryAdjustments({ month: filterMonth, year: filterYear }),
        financeService.getPayrollStatus(filterMonth, filterYear)
      ]);
      setAdjustments(aData);
      setPayrollStatus(pStatus);
    } catch (error) {
      console.error('Error fetching adjustments:', error);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchAdjustments();
  }, [filterMonth, filterYear]);

  const handleDelete = async (id: string) => {
    if (payrollStatus && payrollStatus.status !== 'Draft') {
      Swal.fire('Gagal!', `Periode gaji ini sudah di-ACC (${payrollStatus.status}). Data tidak dapat dihapus.`, 'error');
      return;
    }

    const result = await Swal.fire({
      title: 'Hapus Penyesuaian?',
      text: "Data ini akan dihapus permanen dari sistem.",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
      confirmButtonText: 'Ya, Hapus'
    });

    if (result.isConfirmed) {
      try {
        await financeService.deleteSalaryAdjustment(id);
        Swal.fire('Berhasil!', 'Penyesuaian gaji berhasil dihapus.', 'success');
        fetchAdjustments();
      } catch (error) {
        Swal.fire('Gagal!', 'Terjadi kesalahan saat menghapus data.', 'error');
      }
    }
  };

  const filteredAdjustments = adjustments.filter(as => 
    as.account?.full_name.toLowerCase().includes(searchQuery.toLowerCase()) ||
    as.account?.internal_nik.toLowerCase().includes(searchQuery.toLowerCase()) ||
    as.description.toLowerCase().includes(searchQuery.toLowerCase())
  );

  const isLocked = payrollStatus && payrollStatus.status !== 'Draft';

  if (activeView === 'form') {
    return (
      <SalaryAdjustmentForm 
        adjustment={selectedAdjustment}
        onBack={() => setActiveView('list')}
        onSuccess={() => {
          setActiveView('list');
          fetchAdjustments();
        }}
      />
    );
  }

  if (activeView === 'detail' && selectedAdjustment) {
    return (
      <div className="max-w-2xl mx-auto space-y-6">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-4">
            <button
              onClick={() => setActiveView('list')}
              className="p-2 hover:bg-gray-100 rounded-full transition-colors text-gray-500"
            >
              <Calendar size={20} />
            </button>
            <div>
              <h2 className="text-2xl font-bold text-gray-800">Detail Kustom Gaji</h2>
              <p className="text-sm text-gray-500">Informasi lengkap penyesuaian gaji insidental.</p>
            </div>
          </div>
        </div>

        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
          <div className={`p-6 flex items-center justify-between ${
            selectedAdjustment.type === 'Addition' ? 'bg-emerald-50' : 'bg-rose-50'
          }`}>
            <div className="flex items-center gap-4">
              <div className={`p-3 rounded-xl ${
                selectedAdjustment.type === 'Addition' ? 'bg-emerald-600 text-white' : 'bg-rose-600 text-white'
              }`}>
                <DollarSign size={24} />
              </div>
              <div>
                <div className="text-xs font-bold uppercase tracking-widest opacity-70">
                  {selectedAdjustment.type === 'Addition' ? 'Tambahan Gaji' : 'Potongan Gaji'}
                </div>
                <div className="text-2xl font-bold">
                  Rp {selectedAdjustment.amount.toLocaleString('id-ID')}
                </div>
              </div>
            </div>
            <div className={`px-4 py-1.5 rounded-full text-[10px] font-bold uppercase tracking-wider ${
              selectedAdjustment.type === 'Addition' ? 'bg-emerald-100 text-emerald-700' : 'bg-rose-100 text-rose-700'
            }`}>
              {selectedAdjustment.type === 'Addition' ? 'Addition' : 'Deduction'}
            </div>
          </div>

          <div className="p-8 space-y-8">
            <div className="grid grid-cols-2 gap-8">
              <div className="space-y-1">
                <div className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Karyawan</div>
                <div className="font-bold text-gray-800">{selectedAdjustment.account?.full_name}</div>
                <div className="text-xs text-gray-500">{selectedAdjustment.account?.internal_nik}</div>
              </div>
              <div className="space-y-1">
                <div className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Periode Gaji</div>
                <div className="font-bold text-gray-800">
                  {new Date(0, selectedAdjustment.month - 1).toLocaleString('id-ID', { month: 'long' })} {selectedAdjustment.year}
                </div>
              </div>
            </div>

            <div className="space-y-1">
              <div className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Keterangan</div>
              <div className="p-4 bg-gray-50 rounded-xl text-sm text-gray-700 leading-relaxed italic">
                "{selectedAdjustment.description}"
              </div>
            </div>

            <div className="pt-6 border-t border-gray-100 flex items-center justify-between text-[10px] text-gray-400 font-bold uppercase tracking-widest">
              <div>Dibuat Pada: {new Date(selectedAdjustment.created_at).toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' })}</div>
              <div>ID: {selectedAdjustment.id.split('-')[0]}</div>
            </div>
          </div>
        </div>

        <div className="flex justify-center gap-3">
          <button
            onClick={() => setActiveView('list')}
            className="px-8 py-2.5 bg-gray-100 text-gray-600 rounded-xl hover:bg-gray-200 transition-all text-xs font-bold uppercase tracking-widest"
          >
            Kembali
          </button>
          {!isLocked && (
            <button
              onClick={() => {
                setSelectedAdjustment(selectedAdjustment);
                setActiveView('form');
              }}
              className="px-8 py-2.5 bg-[#006E62] text-white rounded-xl hover:bg-[#005a50] transition-all text-xs font-bold uppercase tracking-widest shadow-lg"
            >
              Edit Data
            </button>
          )}
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-2xl font-bold text-gray-800">Kustom Gaji</h2>
          <p className="text-sm text-gray-500">Kelola tambahan dan potongan gaji insidental per periode.</p>
        </div>
        <button
          onClick={() => {
            setSelectedAdjustment(null);
            setActiveView('form');
          }}
          className="flex items-center gap-2 px-6 py-2.5 bg-[#006E62] text-white rounded-lg hover:bg-[#005a50] transition-all text-sm font-bold uppercase tracking-wider shadow-md"
        >
          <Plus size={18} />
          Tambah Kustom Gaji
        </button>
      </div>

      {/* Filters & Status */}
      <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
        <div className="bg-white p-4 rounded-xl shadow-sm border border-gray-100 flex items-center gap-4">
          <div className="p-2 bg-emerald-50 text-emerald-600 rounded-lg">
            <Calendar size={20} />
          </div>
          <div className="flex-1">
            <div className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Bulan</div>
            <select
              value={filterMonth}
              onChange={(e) => setFilterMonth(Number(e.target.value))}
              className="w-full bg-transparent border-none p-0 text-sm font-bold text-gray-800 focus:ring-0"
            >
              {Array.from({ length: 12 }, (_, i) => (
                <option key={i + 1} value={i + 1}>
                  {new Date(0, i).toLocaleString('id-ID', { month: 'long' })}
                </option>
              ))}
            </select>
          </div>
        </div>

        <div className="bg-white p-4 rounded-xl shadow-sm border border-gray-100 flex items-center gap-4">
          <div className="p-2 bg-blue-50 text-blue-600 rounded-lg">
            <Calendar size={20} />
          </div>
          <div className="flex-1">
            <div className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tahun</div>
            <select
              value={filterYear}
              onChange={(e) => setFilterYear(Number(e.target.value))}
              className="w-full bg-transparent border-none p-0 text-sm font-bold text-gray-800 focus:ring-0"
            >
              {Array.from({ length: 5 }, (_, i) => (
                <option key={i} value={new Date().getFullYear() - 2 + i}>
                  {new Date().getFullYear() - 2 + i}
                </option>
              ))}
            </select>
          </div>
        </div>

        <div className="md:col-span-2 bg-white p-4 rounded-xl shadow-sm border border-gray-100 flex items-center justify-between px-6">
          <div className="flex items-center gap-4">
            <div className={`p-2 rounded-lg ${
              isLocked ? 'bg-rose-50 text-rose-600' : 'bg-emerald-50 text-emerald-600'
            }`}>
              {isLocked ? <AlertCircle size={20} /> : <Check size={20} />}
            </div>
            <div>
              <div className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Status Payroll</div>
              <div className={`text-sm font-bold ${isLocked ? 'text-rose-600' : 'text-emerald-600'}`}>
                {payrollStatus ? payrollStatus.status.toUpperCase() : 'DRAFT / BELUM ADA'}
              </div>
            </div>
          </div>
          {isLocked && (
            <div className="text-[10px] font-bold text-rose-400 uppercase tracking-widest flex items-center gap-1">
              <Info size={12} />
              Data Terkunci
            </div>
          )}
        </div>
      </div>

      <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="p-4 border-b border-gray-50 flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="relative flex-1 max-w-md">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
            <input
              type="text"
              placeholder="Cari nama, NIK, atau keterangan..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full pl-10 pr-4 py-2 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all"
            />
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead>
              <tr className="bg-gray-50 text-gray-500 text-[10px] uppercase tracking-widest font-bold">
                <th className="px-6 py-4">Karyawan</th>
                <th className="px-6 py-4">Jenis</th>
                <th className="px-6 py-4">Nominal</th>
                <th className="px-6 py-4">Keterangan</th>
                <th className="px-6 py-4 text-right">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {loading ? (
                <tr>
                  <td colSpan={5} className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center gap-2">
                      <div className="w-8 h-8 border-4 border-[#006E62] border-t-transparent rounded-full animate-spin"></div>
                      <span className="text-xs font-bold text-gray-400 uppercase tracking-widest">Memuat Data...</span>
                    </div>
                  </td>
                </tr>
              ) : filteredAdjustments.length === 0 ? (
                <tr>
                  <td colSpan={5} className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center gap-2 text-gray-400">
                      <DollarSign size={48} strokeWidth={1} />
                      <p className="text-sm font-medium">Belum ada data kustom gaji di periode ini.</p>
                    </div>
                  </td>
                </tr>
              ) : (
                filteredAdjustments.map((as) => (
                  <tr key={as.id} className="hover:bg-gray-50 transition-colors group">
                    <td className="px-6 py-4">
                      <div className="font-semibold text-gray-900">{as.account?.full_name}</div>
                      <div className="text-xs text-gray-500">{as.account?.internal_nik}</div>
                    </td>
                    <td className="px-6 py-4">
                      <span className={`px-2 py-1 rounded-full text-[10px] font-bold uppercase ${
                        as.type === 'Addition' ? 'bg-emerald-50 text-emerald-600' : 'bg-rose-50 text-rose-600'
                      }`}>
                        {as.type === 'Addition' ? 'TAMBAHAN' : 'POTONGAN'}
                      </span>
                    </td>
                    <td className={`px-6 py-4 font-mono text-sm font-bold ${
                      as.type === 'Addition' ? 'text-emerald-600' : 'text-rose-600'
                    }`}>
                      {as.type === 'Addition' ? '+' : '-'} Rp {as.amount.toLocaleString('id-ID')}
                    </td>
                    <td className="px-6 py-4">
                      <div className="text-xs text-gray-600 truncate max-w-xs italic">"{as.description}"</div>
                    </td>
                    <td className="px-6 py-4 text-right">
                      <div className="flex items-center justify-end gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                        <button
                          onClick={() => {
                            setSelectedAdjustment(as);
                            setActiveView('detail');
                          }}
                          className="p-2 text-gray-400 hover:text-[#006E62] hover:bg-emerald-50 rounded-lg transition-colors"
                          title="Detail"
                        >
                          <Eye size={18} />
                        </button>
                        {!isLocked && (
                          <>
                            <button
                              onClick={() => {
                                setSelectedAdjustment(as);
                                setActiveView('form');
                              }}
                              className="p-2 text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                              title="Edit"
                            >
                              <Edit2 size={18} />
                            </button>
                            <button
                              onClick={() => handleDelete(as.id)}
                              className="p-2 text-red-600 hover:bg-red-50 rounded-lg transition-colors"
                              title="Hapus"
                            >
                              <Trash2 size={18} />
                            </button>
                          </>
                        )}
                      </div>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

export default SalaryAdjustmentMain;
