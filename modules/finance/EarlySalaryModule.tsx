import React, { useState, useEffect } from 'react';
import { Plus, Search, Filter, Receipt, ArrowLeft, CheckCircle2, XCircle, Clock, FileText, Download, Wallet } from 'lucide-react';
import { financeService } from '../../services/financeService';
import { authService } from '../../services/authService';
import { EarlySalaryRequest, EarlySalaryStatus } from '../../types';
import EarlySalaryForm from './EarlySalaryForm';
import EarlySalaryDetail from './EarlySalaryDetail';
import Swal from 'sweetalert2';

const EarlySalaryModule: React.FC = () => {
  const [requests, setRequests] = useState<EarlySalaryRequest[]>([]);
  const [loading, setLoading] = useState(true);
  const [activeView, setActiveView] = useState<'list' | 'form' | 'detail'>('list');
  const [selectedRequest, setSelectedRequest] = useState<EarlySalaryRequest | null>(null);
  const [filterMonth, setFilterMonth] = useState(new Date().getMonth() + 1);
  const [filterYear, setFilterYear] = useState(new Date().getFullYear());
  
  const user = authService.getCurrentUser();
  const isAdmin = user?.role === 'admin' || user?.is_finance_admin;

  const fetchRequests = async () => {
    setLoading(true);
    try {
      const filters: any = {};
      if (!isAdmin) {
        filters.account_id = user?.id;
      } else {
        filters.month = filterMonth;
        filters.year = filterYear;
      }
      const data = await financeService.getEarlySalaryRequests(filters);
      setRequests(data);
    } catch (error) {
      console.error('Error fetching early salary requests:', error);
      Swal.fire('Error', 'Gagal memuat data pengajuan gaji awal.', 'error');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchRequests();
  }, [filterMonth, filterYear]);

  const handleViewDetail = (r: EarlySalaryRequest) => {
    setSelectedRequest(r);
    setActiveView('detail');
  };

  const getStatusColor = (status: EarlySalaryStatus) => {
    switch (status) {
      case 'Approved': return 'bg-blue-50 text-blue-600 border-blue-100';
      case 'Paid': return 'bg-emerald-50 text-emerald-600 border-emerald-100';
      case 'Rejected': return 'bg-rose-50 text-rose-600 border-rose-100';
      default: return 'bg-amber-50 text-amber-600 border-amber-100';
    }
  };

  const getStatusLabel = (status: EarlySalaryStatus) => {
    switch (status) {
      case 'Approved': return 'Disetujui';
      case 'Paid': return 'Sudah Dibayar';
      case 'Rejected': return 'Ditolak';
      default: return 'Pending';
    }
  };

  if (activeView === 'form') {
    return (
      <EarlySalaryForm 
        onBack={() => {
          setActiveView('list');
          fetchRequests();
        }} 
      />
    );
  }

  if (activeView === 'detail' && selectedRequest) {
    return (
      <EarlySalaryDetail 
        request={selectedRequest}
        isAdmin={isAdmin}
        onBack={() => {
          setActiveView('list');
          setSelectedRequest(null);
          fetchRequests();
        }}
      />
    );
  }

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-2xl font-bold text-gray-800">Ambil Gaji Awal</h2>
          <p className="text-sm text-gray-500">
            {isAdmin ? 'Verifikasi dan kelola pengajuan gaji awal (kasbon) pegawai.' : 'Ajukan pembayaran gaji awal untuk keperluan mendesak.'}
          </p>
        </div>
        {!isAdmin && (
          <button
            onClick={() => setActiveView('form')}
            className="flex items-center gap-2 px-4 py-2 bg-[#006E62] text-white rounded-lg hover:bg-[#005a50] transition-all text-sm font-medium shadow-md"
          >
            <Plus size={18} />
            Ajukan Gaji Awal
          </button>
        )}
      </div>

      <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="p-4 border-b border-gray-50 flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="flex items-center gap-2">
            <Filter size={18} className="text-gray-400" />
            <select 
              value={filterMonth}
              onChange={(e) => setFilterMonth(Number(e.target.value))}
              className="bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all"
            >
              {Array.from({ length: 12 }, (_, i) => (
                <option key={i + 1} value={i + 1}>
                  {new Date(0, i).toLocaleString('id-ID', { month: 'long' })}
                </option>
              ))}
            </select>
            <select 
              value={filterYear}
              onChange={(e) => setFilterYear(Number(e.target.value))}
              className="bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all"
            >
              {Array.from({ length: 5 }, (_, i) => (
                <option key={i} value={new Date().getFullYear() - i}>
                  {new Date().getFullYear() - i}
                </option>
              ))}
            </select>
          </div>
          {isAdmin && (
            <button
              className="flex items-center gap-2 px-3 py-1.5 text-xs font-bold text-[#006E62] hover:bg-emerald-50 rounded-lg transition-all"
              onClick={() => Swal.fire('Info', 'Fitur ekspor sedang disiapkan.', 'info')}
            >
              <Download size={16} />
              EKSPOR DATA
            </button>
          )}
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead>
              <tr className="bg-gray-50 text-gray-500 text-[10px] uppercase tracking-widest font-bold">
                <th className="px-6 py-4">Periode</th>
                {isAdmin && <th className="px-6 py-4">Pegawai</th>}
                <th className="px-6 py-4">Alasan</th>
                <th className="px-6 py-4">Nominal</th>
                <th className="px-6 py-4">Status</th>
                <th className="px-6 py-4 text-right">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {loading ? (
                <tr>
                  <td colSpan={isAdmin ? 6 : 5} className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center gap-2">
                      <div className="w-8 h-8 border-4 border-[#006E62] border-t-transparent rounded-full animate-spin"></div>
                      <span className="text-xs font-bold text-gray-400 uppercase tracking-widest">Memuat Data...</span>
                    </div>
                  </td>
                </tr>
              ) : requests.length === 0 ? (
                <tr>
                  <td colSpan={isAdmin ? 6 : 5} className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center gap-2 text-gray-400">
                      <Wallet size={48} strokeWidth={1} />
                      <p className="text-sm font-medium">Belum ada pengajuan gaji awal.</p>
                    </div>
                  </td>
                </tr>
              ) : (
                requests.map((r) => (
                  <tr key={r.id} className="hover:bg-gray-50 transition-colors group">
                    <td className="px-6 py-4">
                      <div className="text-sm font-bold text-gray-900">
                        {new Date(r.year, r.month - 1).toLocaleString('id-ID', { month: 'long', year: 'numeric' })}
                      </div>
                      <div className="text-[10px] text-gray-400 uppercase font-bold">
                        Input: {new Date(r.created_at).toLocaleDateString('id-ID')}
                      </div>
                    </td>
                    {isAdmin && (
                      <td className="px-6 py-4">
                        <div className="font-semibold text-gray-900">{r.account?.full_name}</div>
                        <div className="text-xs text-gray-500">{r.account?.internal_nik}</div>
                      </td>
                    )}
                    <td className="px-6 py-4">
                      <div className="text-sm text-gray-600 truncate max-w-xs">{r.reason}</div>
                    </td>
                    <td className="px-6 py-4">
                      <div className="text-sm font-mono font-bold text-gray-900">
                        Rp {r.amount.toLocaleString('id-ID')}
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <span className={`px-2 py-1 rounded-full text-[10px] font-bold uppercase border ${getStatusColor(r.status)}`}>
                        {getStatusLabel(r.status)}
                      </span>
                    </td>
                    <td className="px-6 py-4 text-right">
                      <button
                        onClick={() => handleViewDetail(r)}
                        className="p-2 text-[#006E62] hover:bg-emerald-50 rounded-lg transition-colors"
                        title="Lihat Detail"
                      >
                        <FileText size={18} />
                      </button>
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

export default EarlySalaryModule;
