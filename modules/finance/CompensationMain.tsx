
import React, { useState, useEffect } from 'react';
import { Search, Filter, Receipt, FileText, Download, CheckCircle2, Clock, History, AlertCircle } from 'lucide-react';
import { financeService } from '../../services/financeService';
import { authService } from '../../services/authService';
import { Compensation, CompensationStatus } from '../../types';
import CompensationDetail from './CompensationDetail';
import Swal from 'sweetalert2';

const CompensationMain: React.FC = () => {
  const [compensations, setCompensations] = useState<Compensation[]>([]);
  const [loading, setLoading] = useState(true);
  const [activeTab, setActiveTab] = useState<CompensationStatus>('Pending');
  const [selectedCompensation, setSelectedCompensation] = useState<Compensation | null>(null);
  const [isDetailOpen, setIsDetailOpen] = useState(false);
  
  const user = authService.getCurrentUser();
  const isAdmin = user?.role === 'admin' || user?.is_finance_admin;

  const fetchCompensations = async () => {
    setLoading(true);
    try {
      const data = await financeService.getCompensations({ status: activeTab });
      setCompensations(data);
    } catch (error) {
      console.error('Error fetching compensations:', error);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchCompensations();
  }, [activeTab]);

  const handleViewDetail = async (c: Compensation) => {
    setSelectedCompensation(c);
    setIsDetailOpen(true);
    if (isAdmin && !c.is_read) {
      try {
        await financeService.markCompensationAsRead(c.id);
        // Update local state to remove NEW label
        setCompensations(prev => prev.map(item => item.id === c.id ? { ...item, is_read: true } : item));
      } catch (error) {
        console.error('Error marking as read:', error);
      }
    }
  };

  const getStatusColor = (status: string) => {
    switch (status) {
      case 'Completed': return 'bg-emerald-50 text-emerald-600 border-emerald-100';
      default: return 'bg-amber-50 text-amber-600 border-amber-100';
    }
  };

  if (!isAdmin) {
    return (
      <div className="flex flex-col items-center justify-center h-64 text-gray-400 bg-gray-50 rounded-lg border border-dashed border-gray-200">
        <AlertCircle size={48} strokeWidth={1} className="mb-2" />
        <p className="font-medium text-sm">Hanya Admin Finance yang dapat mengakses modul ini.</p>
      </div>
    );
  }

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-2xl font-bold text-gray-800">Kompensasi</h2>
          <p className="text-sm text-gray-500">Kelola pesangon dan pinalti karyawan yang resign atau dipecat.</p>
        </div>
      </div>

      {/* Tabs */}
      <div className="flex border-b border-gray-100">
        <button
          onClick={() => setActiveTab('Pending')}
          className={`px-6 py-3 text-sm font-bold uppercase tracking-widest transition-all border-b-2 ${
            activeTab === 'Pending' 
              ? 'border-[#006E62] text-[#006E62]' 
              : 'border-transparent text-gray-400 hover:text-gray-600'
          }`}
        >
          <div className="flex items-center gap-2">
            <Clock size={16} />
            Antrean
            {compensations.filter(c => !c.is_read && activeTab === 'Pending').length > 0 && (
              <span className="w-2 h-2 bg-red-500 rounded-full"></span>
            )}
          </div>
        </button>
        <button
          onClick={() => setActiveTab('Completed')}
          className={`px-6 py-3 text-sm font-bold uppercase tracking-widest transition-all border-b-2 ${
            activeTab === 'Completed' 
              ? 'border-[#006E62] text-[#006E62]' 
              : 'border-transparent text-gray-400 hover:text-gray-600'
          }`}
        >
          <div className="flex items-center gap-2">
            <History size={16} />
            Riwayat
          </div>
        </button>
      </div>

      <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead>
              <tr className="bg-gray-50 text-gray-500 text-[10px] uppercase tracking-widest font-bold">
                <th className="px-6 py-4">Pegawai</th>
                <th className="px-6 py-4">Tipe & Tanggal Exit</th>
                <th className="px-6 py-4">Jenis & Nominal</th>
                <th className="px-6 py-4">Status</th>
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
              ) : compensations.length === 0 ? (
                <tr>
                  <td colSpan={5} className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center gap-2 text-gray-400">
                      <Receipt size={48} strokeWidth={1} />
                      <p className="text-sm font-medium">Tidak ada data kompensasi.</p>
                    </div>
                  </td>
                </tr>
              ) : (
                compensations.map((c) => (
                  <tr key={c.id} className="hover:bg-gray-50 transition-colors group">
                    <td className="px-6 py-4">
                      <div className="flex items-center gap-2">
                        <div className="font-semibold text-gray-900">{c.account?.full_name}</div>
                        {!c.is_read && c.status === 'Pending' && (
                          <span className="px-1.5 py-0.5 bg-red-500 text-white text-[8px] font-black rounded uppercase tracking-tighter animate-pulse">NEW</span>
                        )}
                      </div>
                      <div className="text-xs text-gray-500">{c.account?.internal_nik}</div>
                    </td>
                    <td className="px-6 py-4">
                      <div className={`text-xs font-bold uppercase tracking-wider mb-1 ${c.termination_type === 'Pemecatan' ? 'text-red-600' : 'text-orange-600'}`}>
                        {c.termination_type}
                      </div>
                      <div className="text-sm text-gray-600">
                        {new Date(c.termination_date).toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' })}
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <div className="text-xs font-bold text-gray-500 uppercase tracking-wider mb-1">{c.type === 'Severance' ? 'Pesangon' : 'Pinalti'}</div>
                      <div className="text-sm font-mono font-bold text-gray-900">
                        Rp {c.amount.toLocaleString('id-ID')}
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <span className={`px-2 py-1 rounded-full text-[10px] font-bold uppercase border ${getStatusColor(c.status)}`}>
                        {c.status === 'Completed' ? 'Selesai' : 'Menunggu'}
                      </span>
                    </td>
                    <td className="px-6 py-4 text-right">
                      <button
                        onClick={() => handleViewDetail(c)}
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

      {isDetailOpen && selectedCompensation && (
        <CompensationDetail 
          compensation={selectedCompensation}
          onClose={() => {
            setIsDetailOpen(false);
            setSelectedCompensation(null);
          }}
          onSuccess={() => {
            setIsDetailOpen(false);
            setSelectedCompensation(null);
            fetchCompensations();
          }}
        />
      )}
    </div>
  );
};

export default CompensationMain;
