import React, { useState, useEffect } from 'react';
import { Plus, Search, Filter, Receipt, ArrowLeft, CheckCircle2, XCircle, Clock, FileText, Download } from 'lucide-react';
import { financeService } from '../../services/financeService';
import { authService } from '../../services/authService';
import { Reimbursement } from '../../types';
import ReimbursementForm from './ReimbursementForm';
import ReimbursementDetail from './ReimbursementDetail';
import Swal from 'sweetalert2';

const ReimbursementMain: React.FC = () => {
  const [reimbursements, setReimbursements] = useState<Reimbursement[]>([]);
  const [loading, setLoading] = useState(true);
  const [activeView, setActiveView] = useState<'list' | 'form' | 'detail'>('list');
  const [selectedReimbursement, setSelectedReimbursement] = useState<Reimbursement | null>(null);
  const [filterMonth, setFilterMonth] = useState(new Date().getMonth() + 1);
  const [filterYear, setFilterYear] = useState(new Date().getFullYear());
  
  const user = authService.getCurrentUser();
  const isAdmin = user?.role === 'admin' || user?.is_finance_admin;

  const fetchReimbursements = async () => {
    setLoading(true);
    try {
      const filters: any = {};
      if (!isAdmin) {
        filters.account_id = user?.id;
      } else {
        filters.month = filterMonth;
        filters.year = filterYear;
      }
      const data = await financeService.getReimbursements(filters);
      setReimbursements(data);
    } catch (error) {
      console.error('Error fetching reimbursements:', error);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchReimbursements();
  }, [filterMonth, filterYear]);

  const handleViewDetail = async (r: Reimbursement) => {
    setSelectedReimbursement(r);
    setActiveView('detail');
    if (isAdmin && !r.is_read) {
      try {
        await financeService.markAsRead(r.id);
        // Update local state to remove NEW label
        setReimbursements(prev => prev.map(item => item.id === r.id ? { ...item, is_read: true } : item));
      } catch (error) {
        console.error('Error marking as read:', error);
      }
    }
  };

  const getStatusColor = (status: string) => {
    switch (status) {
      case 'Approved': return 'bg-emerald-50 text-emerald-600 border-emerald-100';
      case 'Partially Approved': return 'bg-blue-50 text-blue-600 border-blue-100';
      case 'Rejected': return 'bg-rose-50 text-rose-600 border-rose-100';
      default: return 'bg-amber-50 text-amber-600 border-amber-100';
    }
  };

  if (activeView === 'form') {
    return (
      <ReimbursementForm 
        onBack={() => {
          setActiveView('list');
          fetchReimbursements();
        }} 
      />
    );
  }

  if (activeView === 'detail' && selectedReimbursement) {
    return (
      <ReimbursementDetail 
        reimbursement={selectedReimbursement}
        isAdmin={isAdmin}
        onBack={() => {
          setActiveView('list');
          setSelectedReimbursement(null);
          fetchReimbursements();
        }}
      />
    );
  }

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-2xl font-bold text-gray-800">Reimburse</h2>
          <p className="text-sm text-gray-500">
            {isAdmin ? 'Verifikasi dan kelola pengajuan reimburse pegawai.' : 'Ajukan dan pantau status reimburse Anda.'}
          </p>
        </div>
        {!isAdmin && (
          <button
            onClick={() => setActiveView('form')}
            className="flex items-center gap-2 px-4 py-2 bg-[#006E62] text-white rounded-lg hover:bg-[#005a50] transition-all text-sm font-medium shadow-md"
          >
            <Plus size={18} />
            Ajukan Reimburse
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
              {Array.from({ length: 10 }, (_, i) => (
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
                <th className="px-6 py-4">Tanggal</th>
                {isAdmin && <th className="px-6 py-4">Pegawai</th>}
                <th className="px-6 py-4">Kategori & Keterangan</th>
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
              ) : reimbursements.length === 0 ? (
                <tr>
                  <td colSpan={isAdmin ? 6 : 5} className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center gap-2 text-gray-400">
                      <Receipt size={48} strokeWidth={1} />
                      <p className="text-sm font-medium">Belum ada pengajuan reimburse.</p>
                    </div>
                  </td>
                </tr>
              ) : (
                reimbursements.map((r) => (
                  <tr key={r.id} className="hover:bg-gray-50 transition-colors group">
                    <td className="px-6 py-4">
                      <div className="text-sm font-medium text-gray-900">
                        {new Date(r.transaction_date).toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' })}
                      </div>
                      <div className="text-[10px] text-gray-400 uppercase font-bold">
                        Input: {new Date(r.created_at).toLocaleDateString('id-ID')}
                      </div>
                    </td>
                    {isAdmin && (
                      <td className="px-6 py-4">
                        <div className="flex items-center gap-2">
                          <div className="font-semibold text-gray-900">{r.account?.full_name}</div>
                          {isAdmin && !r.is_read && (
                            <span className="px-1.5 py-0.5 bg-red-500 text-white text-[8px] font-black rounded uppercase tracking-tighter animate-pulse">NEW</span>
                          )}
                        </div>
                        <div className="text-xs text-gray-500">{r.account?.internal_nik}</div>
                      </td>
                    )}
                    <td className="px-6 py-4">
                      <div className="text-xs font-bold text-[#006E62] uppercase tracking-wider mb-1">{r.category}</div>
                      <div className="text-sm text-gray-600 truncate max-w-xs">{r.description}</div>
                    </td>
                    <td className="px-6 py-4">
                      <div className="text-sm font-mono font-bold text-gray-900">
                        Rp {r.amount_requested.toLocaleString('id-ID')}
                      </div>
                      {r.amount_approved !== null && r.amount_approved !== r.amount_requested && (
                        <div className="text-[10px] font-bold text-emerald-600 uppercase">
                          Disetujui: Rp {r.amount_approved.toLocaleString('id-ID')}
                        </div>
                      )}
                    </td>
                    <td className="px-6 py-4">
                      <span className={`px-2 py-1 rounded-full text-[10px] font-bold uppercase border ${getStatusColor(r.status)}`}>
                        {r.status === 'Partially Approved' ? 'Setuju Sebagian' : 
                         r.status === 'Approved' ? 'Disetujui' : 
                         r.status === 'Rejected' ? 'Ditolak' : 'Pending'}
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

export default ReimbursementMain;
