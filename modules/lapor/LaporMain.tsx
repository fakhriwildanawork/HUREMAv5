import React, { useState, useEffect } from 'react';
import { AlertTriangle, Plus, Search, Calendar, User, Shield, CheckCircle2, Clock, Eye, Trash2, Filter, AlertCircle, Users } from 'lucide-react';
import Swal from 'sweetalert2';
import { whistleblowingService } from '../../services/whistleblowingService';
import { authService } from '../../services/authService';
import { accountService } from '../../services/accountService';
import { Whistleblowing, AuthUser, Account } from '../../types';
import LaporForm from './LaporForm';
import LaporDetail from './LaporDetail';
import LoadingSpinner from '../../components/Common/LoadingSpinner';
import { CardSkeleton } from '../../components/Common/Skeleton';

const LaporMain: React.FC = () => {
  const [reports, setReports] = useState<Whistleblowing[]>([]);
  const [accounts, setAccounts] = useState<Account[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [categoryFilter, setCategoryFilter] = useState('All');
  const [showForm, setShowForm] = useState(false);
  const [selectedReport, setSelectedReport] = useState<Whistleblowing | null>(null);
  const [showDetail, setShowDetail] = useState(false);
  const [user, setUser] = useState<AuthUser | null>(null);

  const isAdmin = user?.role === 'admin' || user?.is_hr_admin || user?.is_performance_admin || user?.is_finance_admin;

  useEffect(() => {
    const currentUser = authService.getCurrentUser();
    setUser(currentUser);
    fetchData(currentUser);
    fetchAccounts();
  }, []);

  const fetchData = async (currentUser: AuthUser | null) => {
    if (!currentUser) return;
    const isUserAdmin = currentUser.role === 'admin' || currentUser.is_hr_admin || currentUser.is_performance_admin || currentUser.is_finance_admin;
    try {
      setIsLoading(true);
      const data = isUserAdmin 
        ? await whistleblowingService.getAll() 
        : await whistleblowingService.getByAccountId(currentUser.id);
      setReports(data);
    } catch (error) {
      console.error(error);
    } finally {
      setIsLoading(false);
    }
  };

  const fetchAccounts = async () => {
    try {
      const data = await accountService.getAll();
      setAccounts(data || []);
    } catch (error) {
      console.error('Failed to fetch accounts:', error);
    }
  };

  const handleSubmitReport = async (input: any) => {
    if (!user) return;
    try {
      setIsSaving(true);
      await whistleblowingService.submitReport({
        ...input,
        account_id: user.id
      });
      await fetchData(user);
      setShowForm(false);
      Swal.fire({
        title: 'Laporan Terkirim!',
        text: 'Laporan Anda telah diterima oleh manajemen untuk ditindaklanjuti.',
        icon: 'success',
        timer: 2500,
        showConfirmButton: false
      });
    } catch (error) {
      Swal.fire('Gagal', 'Terjadi kesalahan saat mengirim laporan.', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleDelete = async (id: string) => {
    const result = await Swal.fire({
      title: 'Hapus Laporan?',
      text: "Data laporan akan dihapus permanen.",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
      confirmButtonText: 'Ya, Hapus!'
    });

    if (result.isConfirmed) {
      try {
        setIsSaving(true);
        await whistleblowingService.delete(id);
        await fetchData(user);
        Swal.fire('Terhapus!', 'Laporan telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Terjadi kesalahan saat menghapus laporan.', 'error');
      } finally {
        setIsSaving(false);
      }
    }
  };

  const handleViewDetail = async (report: Whistleblowing) => {
    setSelectedReport(report);
    setShowDetail(true);
    
    // Auto mark as read for admin
    if (isAdmin && report.status === 'Unread') {
      try {
        await whistleblowingService.markAsRead(report.id);
        // Update local state
        setReports(prev => prev.map(r => r.id === report.id ? { ...r, status: 'Read' } : r));
      } catch (error) {
        console.error('Failed to mark as read:', error);
      }
    }
  };

  const filteredReports = reports.filter(r => {
    const matchesSearch = r.description.toLowerCase().includes(searchTerm.toLowerCase()) ||
                         r.category.toLowerCase().includes(searchTerm.toLowerCase()) ||
                         r.account?.full_name?.toLowerCase().includes(searchTerm.toLowerCase());
    const matchesCategory = categoryFilter === 'All' || r.category === categoryFilter;
    return matchesSearch && matchesCategory;
  });

  const categories = ['All', 'Pencurian', 'Perusakan', 'Bullying', 'Fraud', 'Pelanggaran SOP', 'Lainnya'];

  return (
    <div className="space-y-6">
      {isSaving && <LoadingSpinner message="Memproses Data..." />}

      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-xl font-bold text-gray-800 tracking-tight">Laporan Pelanggaran</h2>
          <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Whistleblowing System - Jaga Integritas Perusahaan</p>
        </div>
        
        <div className="flex flex-wrap items-center gap-3">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
            <input
              type="text"
              placeholder="Cari laporan..."
              className="pl-9 pr-4 py-2 bg-gray-50 border border-gray-100 rounded-md focus:outline-none focus:ring-2 focus:ring-[#006E62] text-xs w-48 md:w-64"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
          
          <select 
            value={categoryFilter}
            onChange={(e) => setCategoryFilter(e.target.value)}
            className="px-3 py-2 bg-gray-50 border border-gray-100 rounded-md focus:outline-none focus:ring-2 focus:ring-[#006E62] text-xs font-bold text-gray-600"
          >
            {categories.map(cat => <option key={cat} value={cat}>{cat}</option>)}
          </select>

          {!isAdmin && (
            <button 
              onClick={() => setShowForm(true)}
              className="flex items-center gap-2 bg-rose-600 text-white px-4 py-2 rounded-md hover:bg-rose-700 transition-colors shadow-sm text-xs font-bold uppercase"
            >
              <Plus size={16} />
              Buat Laporan
            </button>
          )}
        </div>
      </div>

      {isLoading ? (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {[1, 2, 3].map(i => <CardSkeleton key={i} />)}
        </div>
      ) : (
        <div className="space-y-4">
          {isAdmin ? (
            <div className="bg-white rounded-2xl border border-gray-100 shadow-sm overflow-hidden">
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-gray-50/50 border-b border-gray-100">
                      <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Status</th>
                      <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Kategori</th>
                      <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Pelapor</th>
                      <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Terlapor</th>
                      <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tanggal</th>
                      <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">Aksi</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-50">
                    {filteredReports.map(report => (
                      <tr key={report.id} className={`hover:bg-gray-50/50 transition-colors ${report.status === 'Unread' ? 'bg-rose-50/20' : ''}`}>
                        <td className="px-6 py-4">
                          {report.status === 'Unread' ? (
                            <span className="flex items-center gap-1.5 text-[10px] font-bold text-rose-600 uppercase">
                              <div className="w-1.5 h-1.5 bg-rose-600 rounded-full animate-pulse"></div>
                              Baru
                            </span>
                          ) : (
                            <span className="flex items-center gap-1.5 text-[10px] font-bold text-gray-400 uppercase">
                              <CheckCircle2 size={12} />
                              Dibaca
                            </span>
                          )}
                        </td>
                        <td className="px-6 py-4">
                          <span className="px-2 py-1 bg-rose-50 text-rose-600 rounded text-[9px] font-bold uppercase">
                            {report.category}
                          </span>
                        </td>
                        <td className="px-6 py-4">
                          <div className="flex items-center gap-2">
                            <div className="w-6 h-6 rounded-full bg-gray-100 flex items-center justify-center text-[10px] font-bold text-gray-500">
                              {report.account?.full_name?.charAt(0)}
                            </div>
                            <span className="text-xs font-bold text-gray-700">{report.account?.full_name}</span>
                          </div>
                        </td>
                        <td className="px-6 py-4">
                          <div className="flex items-center gap-1 text-[10px] font-bold text-gray-500">
                            <Users size={12} />
                            <span>{report.reported_account_ids?.length || 0} Orang</span>
                          </div>
                        </td>
                        <td className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase">
                          {new Date(report.created_at).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' })}
                        </td>
                        <td className="px-6 py-4 text-right">
                          <div className="flex items-center justify-end gap-2">
                            <button 
                              onClick={() => handleViewDetail(report)}
                              className="p-2 text-gray-400 hover:text-[#006E62] hover:bg-emerald-50 rounded-lg transition-all"
                              title="Lihat Detail"
                            >
                              <Eye size={16} />
                            </button>
                            <button 
                              onClick={() => handleDelete(report.id)}
                              className="p-2 text-gray-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-all"
                              title="Hapus"
                            >
                              <Trash2 size={16} />
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          ) : (
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
              {filteredReports.map(report => (
                <div key={report.id} className="bg-white border border-gray-100 p-5 rounded-2xl shadow-sm hover:shadow-md transition-all border-l-4 border-l-transparent hover:border-l-rose-600 group flex flex-col h-full">
                  <div className="flex justify-between items-start mb-4">
                    <span className="px-2 py-1 bg-rose-50 text-rose-600 rounded text-[9px] font-bold uppercase">
                      {report.category}
                    </span>
                    <div className="flex items-center gap-1">
                      <button 
                        onClick={() => handleViewDetail(report)}
                        className="p-1.5 text-gray-400 hover:text-[#006E62] hover:bg-emerald-50 rounded-lg transition-all"
                      >
                        <Eye size={14} />
                      </button>
                    </div>
                  </div>
                  
                  <div className="mb-3">
                    <p className="text-xs text-gray-600 line-clamp-3 italic">"{report.description}"</p>
                  </div>

                  <div className="mt-auto pt-4 border-t border-gray-50 flex items-center justify-between">
                    <div className="flex items-center gap-2 text-[10px] font-bold text-gray-400 uppercase">
                      <Calendar size={12} />
                      <span>{new Date(report.created_at).toLocaleDateString('id-ID', { day: 'numeric', month: 'short' })}</span>
                    </div>
                    <div className="flex items-center gap-1.5">
                      {report.status === 'Read' ? (
                        <span className="text-[9px] font-bold text-emerald-600 uppercase flex items-center gap-1">
                          <CheckCircle2 size={10} />
                          Ditinjau
                        </span>
                      ) : (
                        <span className="text-[9px] font-bold text-amber-500 uppercase flex items-center gap-1">
                          <Clock size={10} />
                          Pending
                        </span>
                      )}
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}

          {filteredReports.length === 0 && (
            <div className="flex flex-col items-center justify-center py-20 text-gray-400 bg-gray-50 rounded-xl border border-dashed border-gray-200">
              <AlertTriangle size={48} strokeWidth={1} className="mb-4 opacity-20" />
              <p className="text-sm font-bold uppercase tracking-widest">Belum ada laporan ditemukan</p>
            </div>
          )}
        </div>
      )}

      {showForm && (
        <LaporForm 
          onClose={() => setShowForm(false)} 
          onSubmit={handleSubmitReport}
          accounts={accounts.filter(a => a.id !== user?.id)}
        />
      )}

      {showDetail && selectedReport && (
        <LaporDetail 
          report={selectedReport}
          onClose={() => { setShowDetail(false); setSelectedReport(null); }}
          isAdmin={isAdmin}
        />
      )}
    </div>
  );
};

export default LaporMain;
