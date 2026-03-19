import React, { useState, useEffect } from 'react';
import { Plus, Search, Calendar, Trash2, Eye, FileText, CheckCircle, Clock, AlertCircle, Settings } from 'lucide-react';
import { financeService } from '../../services/financeService';
import { Payroll } from '../../types';
import PayrollProcess from './PayrollProcess';
import PayrollSettings from './PayrollSettings';
import PayslipDetail from './PayslipDetail';
import Swal from 'sweetalert2';

const PayrollMain: React.FC = () => {
  const [payrolls, setPayrolls] = useState<Payroll[]>([]);
  const [loading, setLoading] = useState(true);
  const [searchQuery, setSearchQuery] = useState('');
  const [activeView, setActiveView] = useState<'list' | 'process' | 'settings' | 'payslips'>('list');
  const [selectedPayroll, setSelectedPayroll] = useState<Payroll | null>(null);
  const [selectedItems, setSelectedItems] = useState<string[]>([]);

  const fetchPayrolls = async () => {
    setLoading(true);
    try {
      const data = await financeService.getPayrolls();
      setPayrolls(data);
    } catch (error) {
      console.error('Error fetching payrolls:', error);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchPayrolls();
  }, []);

  const handleDelete = async (id: string) => {
    const result = await Swal.fire({
      title: 'Hapus Payroll?',
      text: "Seluruh data payslip pada periode ini akan dihapus permanen.",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
      confirmButtonText: 'Ya, Hapus'
    });

    if (result.isConfirmed) {
      try {
        await financeService.deletePayroll(id);
        Swal.fire('Berhasil!', 'Data payroll berhasil dihapus.', 'success');
        fetchPayrolls();
      } catch (error) {
        Swal.fire('Gagal!', 'Terjadi kesalahan saat menghapus data.', 'error');
      }
    }
  };

  const handleBulkDelete = async () => {
    if (selectedItems.length === 0) return;

    const result = await Swal.fire({
      title: 'Hapus Terpilih?',
      text: `${selectedItems.length} data payroll akan dihapus permanen.`,
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
      confirmButtonText: 'Ya, Hapus Semua'
    });

    if (result.isConfirmed) {
      try {
        await Promise.all(selectedItems.map(id => financeService.deletePayroll(id)));
        Swal.fire('Berhasil!', 'Data terpilih berhasil dihapus.', 'success');
        setSelectedItems([]);
        fetchPayrolls();
      } catch (error) {
        Swal.fire('Gagal!', 'Terjadi kesalahan saat menghapus data.', 'error');
      }
    }
  };

  const toggleSelect = (id: string) => {
    setSelectedItems(prev => 
      prev.includes(id) ? prev.filter(i => i !== id) : [...prev, id]
    );
  };

  const getStatusBadge = (status: Payroll['status']) => {
    switch (status) {
      case 'Draft':
        return <span className="px-2 py-1 bg-gray-100 text-gray-600 rounded-full text-[10px] font-bold uppercase tracking-wider flex items-center gap-1"><Clock size={10} /> Draft</span>;
      case 'Pending':
        return <span className="px-2 py-1 bg-amber-100 text-amber-600 rounded-full text-[10px] font-bold uppercase tracking-wider flex items-center gap-1"><AlertCircle size={10} /> Menunggu Verifikasi</span>;
      case 'Approved':
        return <span className="px-2 py-1 bg-emerald-100 text-emerald-600 rounded-full text-[10px] font-bold uppercase tracking-wider flex items-center gap-1"><CheckCircle size={10} /> Disetujui</span>;
      case 'Paid':
        return <span className="px-2 py-1 bg-blue-100 text-blue-600 rounded-full text-[10px] font-bold uppercase tracking-wider flex items-center gap-1"><FileText size={10} /> Terbit (Sent)</span>;
      default:
        return null;
    }
  };

  const filteredPayrolls = payrolls.filter(p => 
    new Date(0, p.month - 1).toLocaleString('id-ID', { month: 'long' }).toLowerCase().includes(searchQuery.toLowerCase()) ||
    p.year.toString().includes(searchQuery)
  );

  if (activeView === 'process') {
    return (
      <PayrollProcess 
        payroll={selectedPayroll} 
        onBack={() => {
          setActiveView('list');
          fetchPayrolls();
        }} 
      />
    );
  }

  if (activeView === 'settings') {
    return <PayrollSettings onBack={() => setActiveView('list')} />;
  }

  if (activeView === 'payslips' && selectedPayroll) {
    return <PayslipDetail payroll={selectedPayroll} onBack={() => setActiveView('list')} />;
  }

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-2xl font-bold text-gray-800">Payroll Management</h2>
          <p className="text-sm text-gray-500">Kelola perhitungan gaji, verifikasi, dan penerbitan payslip.</p>
        </div>
        <div className="flex items-center gap-3">
          <button
            onClick={() => setActiveView('settings')}
            className="p-2.5 bg-white border border-gray-200 text-gray-600 rounded-lg hover:bg-gray-50 transition-all shadow-sm"
            title="Pengaturan Kop Payslip"
          >
            <Settings size={20} />
          </button>
          <button
            onClick={() => {
              setSelectedPayroll(null);
              setActiveView('process');
            }}
            className="flex items-center gap-2 px-6 py-2.5 bg-[#006E62] text-white rounded-lg hover:bg-[#005a50] transition-all text-sm font-bold uppercase tracking-wider shadow-md"
          >
            <Plus size={18} />
            Buat Payroll Baru
          </button>
        </div>
      </div>

      <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="p-4 border-b border-gray-50 flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="relative flex-1 max-w-md">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
            <input
              type="text"
              placeholder="Cari periode (Bulan/Tahun)..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full pl-10 pr-4 py-2 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all"
            />
          </div>
          {selectedItems.length > 0 && (
            <button
              onClick={handleBulkDelete}
              className="flex items-center gap-2 px-4 py-2 bg-rose-50 text-rose-600 rounded-lg hover:bg-rose-100 transition-all text-xs font-bold uppercase tracking-wider"
            >
              <Trash2 size={16} />
              Hapus Terpilih ({selectedItems.length})
            </button>
          )}
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead>
              <tr className="bg-gray-50 text-gray-500 text-[10px] uppercase tracking-widest font-bold">
                <th className="px-6 py-4 w-10">
                  <input 
                    type="checkbox" 
                    checked={selectedItems.length === filteredPayrolls.length && filteredPayrolls.length > 0}
                    onChange={(e) => {
                      if (e.target.checked) setSelectedItems(filteredPayrolls.map(p => p.id));
                      else setSelectedItems([]);
                    }}
                    className="rounded border-gray-300 text-[#006E62] focus:ring-[#006E62]"
                  />
                </th>
                <th className="px-6 py-4">Periode</th>
                <th className="px-6 py-4">Rentang Analisa</th>
                <th className="px-6 py-4">Status</th>
                <th className="px-6 py-4">Dibuat Oleh</th>
                <th className="px-6 py-4">Verifikator</th>
                <th className="px-6 py-4 text-right">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {loading ? (
                <tr>
                  <td colSpan={7} className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center gap-2">
                      <div className="w-8 h-8 border-4 border-[#006E62] border-t-transparent rounded-full animate-spin"></div>
                      <span className="text-xs font-bold text-gray-400 uppercase tracking-widest">Memuat Data...</span>
                    </div>
                  </td>
                </tr>
              ) : filteredPayrolls.length === 0 ? (
                <tr>
                  <td colSpan={7} className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center gap-2 text-gray-400">
                      <Calendar size={48} strokeWidth={1} />
                      <p className="text-sm font-medium">Belum ada data payroll.</p>
                    </div>
                  </td>
                </tr>
              ) : (
                filteredPayrolls.map((p) => (
                  <tr key={p.id} className="hover:bg-gray-50 transition-colors group">
                    <td className="px-6 py-4">
                      <input 
                        type="checkbox" 
                        checked={selectedItems.includes(p.id)}
                        onChange={() => toggleSelect(p.id)}
                        className="rounded border-gray-300 text-[#006E62] focus:ring-[#006E62]"
                      />
                    </td>
                    <td className="px-6 py-4">
                      <div className="font-bold text-gray-800">
                        {new Date(0, p.month - 1).toLocaleString('id-ID', { month: 'long' })} {p.year}
                      </div>
                      <div className="text-[10px] text-gray-400 uppercase tracking-widest font-medium">ID: {p.id.split('-')[0]}</div>
                    </td>
                    <td className="px-6 py-4">
                      <div className="text-xs font-medium text-gray-600">
                        {new Date(p.start_date).toLocaleDateString('id-ID', { day: '2-digit', month: 'short' })} - {new Date(p.end_date).toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' })}
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      {getStatusBadge(p.status)}
                    </td>
                    <td className="px-6 py-4">
                      <div className="text-xs font-medium text-gray-700">{p.creator?.full_name || '-'}</div>
                      <div className="text-[10px] text-gray-400">
                        {new Date(p.created_at).toLocaleDateString('id-ID')}
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <div className="text-xs font-medium text-gray-700">{p.verifier?.full_name || '-'}</div>
                      {p.verified_at && (
                        <div className="text-[10px] text-gray-400">
                          {new Date(p.verified_at).toLocaleDateString('id-ID')}
                        </div>
                      )}
                    </td>
                    <td className="px-6 py-4 text-right">
                      <div className="flex items-center justify-end gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                        <button
                          onClick={() => {
                            setSelectedPayroll(p);
                            setActiveView('payslips');
                          }}
                          className="p-2 text-gray-400 hover:text-[#006E62] hover:bg-emerald-50 rounded-lg transition-colors"
                          title="Lihat Payslip"
                        >
                          <Eye size={18} />
                        </button>
                        <button
                          onClick={() => {
                            setSelectedPayroll(p);
                            setActiveView('process');
                          }}
                          className="p-2 text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                          title="Edit / Proses"
                        >
                          <FileText size={18} />
                        </button>
                        <button
                          onClick={() => handleDelete(p.id)}
                          className="p-2 text-red-600 hover:bg-red-50 rounded-lg transition-colors"
                          title="Hapus"
                        >
                          <Trash2 size={18} />
                        </button>
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

export default PayrollMain;
