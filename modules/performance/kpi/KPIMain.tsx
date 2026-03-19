import React, { useState, useEffect } from 'react';
import { Target, Plus, Search, Calendar, History, MessageSquare, CheckCircle2, XCircle, Clock, ArrowRight, FileText, User, AlertTriangle, Info, BarChart3, Trash2, Eye, Edit3 } from 'lucide-react';
import Swal from 'sweetalert2';
import { kpiService } from '../../../services/kpiService';
import { authService } from '../../../services/authService';
import { KPI, AuthUser } from '../../../types';
import KPIForm from './KPIForm';
import KPIReportForm from './KPIReportForm';
import KPIVerifyForm from './KPIVerifyForm';
import KPIDetail from './KPIDetail';
import LoadingSpinner from '../../../components/Common/LoadingSpinner';
import { CardSkeleton } from '../../../components/Common/Skeleton';

const KPIMain: React.FC = () => {
  const [kpis, setKpis] = useState<KPI[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [showForm, setShowForm] = useState(false);
  const [selectedKPI, setSelectedKPI] = useState<KPI | null>(null);
  const [showReportForm, setShowReportForm] = useState(false);
  const [showVerifyForm, setShowVerifyForm] = useState(false);
  const [showDetail, setShowDetail] = useState(false);
  const [user, setUser] = useState<AuthUser | null>(null);

  const isAdmin = user?.role === 'admin' || user?.is_performance_admin;

  useEffect(() => {
    const currentUser = authService.getCurrentUser();
    setUser(currentUser);
    fetchKPIs(currentUser);
  }, []);

  const fetchKPIs = async (currentUser: AuthUser | null) => {
    if (!currentUser) return;
    try {
      setIsLoading(true);
      const data = (currentUser.role === 'admin' || currentUser.is_performance_admin)
        ? await kpiService.getAll() 
        : await kpiService.getByAccountId(currentUser.id);
      
      // Auto-update status to Unreported if deadline passed and still Active
      const today = new Date().toLocaleDateString('en-CA');
      const updatedData = data.map(k => {
        if (k.status === 'Active' && k.deadline < today) {
          return { ...k, status: 'Unreported' as const };
        }
        return k;
      });
      
      setKpis(updatedData);
    } catch (error) {
      console.error(error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleCreate = async (input: any) => {
    try {
      setIsSaving(true);
      if (selectedKPI) {
        await kpiService.update(selectedKPI.id, input);
        Swal.fire({
          title: 'Berhasil!',
          text: 'KPI telah diperbarui.',
          icon: 'success',
          timer: 1500,
          showConfirmButton: false
        });
      } else {
        await kpiService.create(input);
        Swal.fire({
          title: 'Berhasil!',
          text: 'KPI baru telah dibuat.',
          icon: 'success',
          timer: 1500,
          showConfirmButton: false
        });
      }
      await fetchKPIs(user);
      setShowForm(false);
      setSelectedKPI(null);
    } catch (error) {
      Swal.fire('Gagal', 'Terjadi kesalahan saat menyimpan KPI.', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleReport = async (reportData: any) => {
    if (!selectedKPI) return;
    try {
      setIsSaving(true);
      await kpiService.report(selectedKPI.id, reportData);
      await fetchKPIs(user);
      setShowReportForm(false);
      setSelectedKPI(null);
      Swal.fire({
        title: 'Laporan Terkirim!',
        text: 'Laporan Anda sedang menunggu verifikasi Admin.',
        icon: 'success',
        timer: 1500,
        showConfirmButton: false
      });
    } catch (error) {
      Swal.fire('Gagal', 'Terjadi kesalahan saat mengirim laporan.', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleVerify = async (verificationData: any) => {
    if (!selectedKPI || !user) return;
    try {
      setIsSaving(true);
      await kpiService.verify(selectedKPI.id, user.id, verificationData);
      await fetchKPIs(user);
      setShowVerifyForm(false);
      setSelectedKPI(null);
      Swal.fire({
        title: 'Terverifikasi!',
        text: 'KPI telah berhasil diverifikasi dan dinilai.',
        icon: 'success',
        timer: 1500,
        showConfirmButton: false
      });
    } catch (error) {
      Swal.fire('Gagal', 'Terjadi kesalahan saat verifikasi.', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleDelete = async (id: string) => {
    const result = await Swal.fire({
      title: 'Hapus KPI?',
      text: "Data KPI akan dihapus permanen.",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
      cancelButtonColor: '#3085d6',
      confirmButtonText: 'Ya, Hapus!',
      cancelButtonText: 'Batal'
    });

    if (result.isConfirmed) {
      try {
        setIsSaving(true);
        await kpiService.delete(id);
        await fetchKPIs(user);
        Swal.fire('Terhapus!', 'KPI telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Terjadi kesalahan saat menghapus KPI.', 'error');
      } finally {
        setIsSaving(false);
      }
    }
  };

  const filteredKPIs = kpis.filter(k => {
    const searchStr = `${k.account?.full_name || ''} ${k.title} ${k.description}`.toLowerCase();
    return searchStr.includes(searchTerm.toLowerCase());
  });

  const today = new Date().toLocaleDateString('en-CA');
  const threeDaysLater = new Date(Date.now() + 3 * 24 * 60 * 60 * 1000).toLocaleDateString('en-CA');

  const upcomingKPIs = filteredKPIs.filter(k => k.start_date > today && k.status === 'Active');
  const warningKPIs = filteredKPIs.filter(k => k.deadline <= threeDaysLater && k.deadline > today && k.status === 'Active');
  const alertKPIs = filteredKPIs.filter(k => k.deadline === today && k.status === 'Active');
  const overdueKPIs = filteredKPIs.filter(k => k.status === 'Unreported');
  const historyKPIs = filteredKPIs.filter(k => k.status === 'Unverified' || k.status === 'Verified');

  const getStatusBadge = (status: string, deadline: string) => {
    if (status === 'Unreported') return <span className="px-2 py-0.5 bg-rose-50 text-rose-600 rounded text-[9px] font-bold uppercase">Overdue</span>;
    if (status === 'Unverified') return <span className="px-2 py-0.5 bg-blue-50 text-blue-600 rounded text-[9px] font-bold uppercase">Unverified</span>;
    if (status === 'Verified') return <span className="px-2 py-0.5 bg-emerald-50 text-emerald-600 rounded text-[9px] font-bold uppercase">Verified</span>;
    if (status === 'Pause') return <span className="px-2 py-0.5 bg-gray-100 text-gray-400 rounded text-[9px] font-bold uppercase">Paused</span>;
    
    if (deadline === today) return <span className="px-2 py-0.5 bg-orange-50 text-orange-600 rounded text-[9px] font-bold uppercase animate-pulse">Alert</span>;
    if (deadline <= threeDaysLater) return <span className="px-2 py-0.5 bg-amber-50 text-amber-600 rounded text-[9px] font-bold uppercase">Warning</span>;
    
    return <span className="px-2 py-0.5 bg-emerald-50 text-emerald-600 rounded text-[9px] font-bold uppercase">Active</span>;
  };

  const calculatePerformanceScore = () => {
    const verifiedKPIs = kpis.filter(k => k.status === 'Verified');
    if (verifiedKPIs.length === 0) return 0;
    
    const totalWeightedScore = verifiedKPIs.reduce((sum, k) => sum + (k.verification_data?.score || 0) * k.weight, 0);
    const totalWeight = verifiedKPIs.reduce((sum, k) => sum + k.weight, 0);
    
    return Math.round(totalWeightedScore / totalWeight);
  };

  const performanceScore = calculatePerformanceScore();

  return (
    <div className="space-y-6">
      {isSaving && <LoadingSpinner message="Memproses Data..." />}

      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-xl font-bold text-gray-800 tracking-tight">Key Performance Indicator</h2>
          <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Pengukuran & Penilaian Kinerja Pegawai</p>
        </div>
        
        <div className="flex items-center gap-3">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
            <input
              type="text"
              placeholder="Cari KPI..."
              className="pl-9 pr-4 py-2 bg-gray-50 border border-gray-100 rounded-md focus:outline-none focus:ring-2 focus:ring-[#006E62] text-xs w-64"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
          {isAdmin && (
            <button 
              onClick={() => setShowForm(true)}
              className="flex items-center gap-2 bg-[#006E62] text-white px-4 py-2 rounded-md hover:bg-[#005a50] transition-colors shadow-sm text-xs font-bold uppercase"
            >
              <Plus size={16} />
              Buat KPI
            </button>
          )}
        </div>
      </div>

      {!isAdmin && (
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
          <div className="bg-white p-5 rounded-2xl border border-gray-100 shadow-sm flex items-center gap-4">
            <div className="w-12 h-12 rounded-xl bg-emerald-50 flex items-center justify-center text-[#006E62]">
              <BarChart3 size={24} />
            </div>
            <div>
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-0.5">Performance Score</p>
              <p className="text-2xl font-bold text-[#006E62]">{performanceScore}%</p>
            </div>
          </div>
          <div className="bg-white p-5 rounded-2xl border border-gray-100 shadow-sm flex items-center gap-4">
            <div className="w-12 h-12 rounded-xl bg-rose-50 flex items-center justify-center text-rose-600">
              <AlertTriangle size={24} />
            </div>
            <div>
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-0.5">Overdue KPI</p>
              <p className="text-2xl font-bold text-rose-600">{overdueKPIs.length}</p>
            </div>
          </div>
          <div className="bg-white p-5 rounded-2xl border border-gray-100 shadow-sm flex items-center gap-4">
            <div className="w-12 h-12 rounded-xl bg-amber-50 flex items-center justify-center text-amber-600">
              <Clock size={24} />
            </div>
            <div>
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-0.5">Warning / Alert</p>
              <p className="text-2xl font-bold text-amber-600">{warningKPIs.length + alertKPIs.length}</p>
            </div>
          </div>
          <div className="bg-white p-5 rounded-2xl border border-gray-100 shadow-sm flex items-center gap-4">
            <div className="w-12 h-12 rounded-xl bg-blue-50 flex items-center justify-center text-blue-600">
              <History size={24} />
            </div>
            <div>
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-0.5">Total KPI</p>
              <p className="text-2xl font-bold text-blue-600">{kpis.length}</p>
            </div>
          </div>
        </div>
      )}

      {isLoading ? (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {[1, 2, 3].map(i => <CardSkeleton key={i} />)}
        </div>
      ) : (
        <div className="space-y-8">
          {/* Section: Active & Urgent (User Dashboard) */}
          {!isAdmin && (alertKPIs.length > 0 || warningKPIs.length > 0 || overdueKPIs.length > 0) && (
            <div className="space-y-4">
              <div className="flex items-center gap-2">
                <AlertTriangle size={16} className="text-rose-500" />
                <h3 className="text-xs font-bold text-gray-800 uppercase tracking-widest">Butuh Perhatian Segera</h3>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                {[...alertKPIs, ...warningKPIs, ...overdueKPIs].map(kpi => (
                  <KPICard 
                    key={kpi.id} 
                    kpi={kpi} 
                    isAdmin={isAdmin} 
                    onReport={() => { setSelectedKPI(kpi); setShowReportForm(true); }}
                    onVerify={() => { setSelectedKPI(kpi); setShowVerifyForm(true); }}
                    onEdit={() => { setSelectedKPI(kpi); setShowForm(true); }}
                    onView={() => { setSelectedKPI(kpi); setShowDetail(true); }}
                    onDelete={() => handleDelete(kpi.id)}
                    getStatusBadge={getStatusBadge}
                  />
                ))}
              </div>
            </div>
          )}

          {/* Section: All KPIs (Admin) or Remaining (User) */}
          <div className="space-y-4">
            <div className="flex items-center gap-2">
              <Target size={16} className="text-[#006E62]" />
              <h3 className="text-xs font-bold text-gray-800 uppercase tracking-widest">
                {isAdmin ? 'Semua Daftar KPI Pegawai' : 'Daftar KPI & Riwayat'}
              </h3>
            </div>
            {filteredKPIs.length === 0 ? (
              <div className="flex flex-col items-center justify-center py-20 text-gray-400 bg-gray-50 rounded-xl border border-dashed border-gray-200">
                <Target size={48} strokeWidth={1} className="mb-4 opacity-20" />
                <p className="text-sm font-bold uppercase tracking-widest">Belum ada data KPI</p>
              </div>
            ) : (
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                {(isAdmin ? filteredKPIs : historyKPIs.concat(upcomingKPIs)).map(kpi => (
                  <KPICard 
                    key={kpi.id} 
                    kpi={kpi} 
                    isAdmin={isAdmin} 
                    onReport={() => { setSelectedKPI(kpi); setShowReportForm(true); }}
                    onVerify={() => { setSelectedKPI(kpi); setShowVerifyForm(true); }}
                    onEdit={() => { setSelectedKPI(kpi); setShowForm(true); }}
                    onView={() => { setSelectedKPI(kpi); setShowDetail(true); }}
                    onDelete={() => handleDelete(kpi.id)}
                    getStatusBadge={getStatusBadge}
                  />
                ))}
              </div>
            )}
          </div>
        </div>
      )}

      {showForm && (
        <KPIForm 
          onClose={() => { setShowForm(false); setSelectedKPI(null); }} 
          onSubmit={handleCreate} 
          initialData={selectedKPI}
        />
      )}

      {showReportForm && selectedKPI && (
        <KPIReportForm 
          kpiTitle={selectedKPI.title}
          onClose={() => { setShowReportForm(false); setSelectedKPI(null); }}
          onSubmit={handleReport}
        />
      )}

      {showVerifyForm && selectedKPI && (
        <KPIVerifyForm 
          kpi={selectedKPI}
          onClose={() => { setShowVerifyForm(false); setSelectedKPI(null); }}
          onSubmit={handleVerify}
        />
      )}

      {showDetail && selectedKPI && (
        <KPIDetail 
          kpi={selectedKPI}
          onClose={() => { setShowDetail(false); setSelectedKPI(null); }}
        />
      )}
    </div>
  );
};

interface KPICardProps {
  kpi: KPI;
  isAdmin: boolean;
  onReport: () => void;
  onVerify: () => void;
  onEdit: () => void;
  onView: () => void;
  onDelete: () => void;
  getStatusBadge: (status: string, deadline: string) => React.ReactNode;
}

const KPICard: React.FC<KPICardProps> = ({ kpi, isAdmin, onReport, onVerify, onEdit, onView, onDelete, getStatusBadge }) => {
  const isVerified = kpi.status === 'Verified';
  const isUnverified = kpi.status === 'Unverified';
  const isOverdue = kpi.status === 'Unreported';

  return (
    <div className="bg-white border border-gray-100 p-5 rounded-2xl shadow-sm hover:shadow-md transition-all border-l-4 border-l-transparent hover:border-l-[#006E62] group flex flex-col h-full">
      <div className="flex justify-between items-start mb-4">
        <div className="flex flex-col">
          {isAdmin && (
            <div className="flex items-center gap-1.5 mb-1">
              <User size={10} className="text-gray-400" />
              <span className="text-[10px] font-bold text-[#006E62] uppercase">{kpi.account?.full_name}</span>
            </div>
          )}
          <h4 className="text-sm font-bold text-gray-800 leading-tight group-hover:text-[#006E62] transition-colors">{kpi.title}</h4>
        </div>
        <div className="flex flex-col items-end gap-2">
          {getStatusBadge(kpi.status, kpi.deadline)}
          <div className="flex items-center gap-1">
            {isAdmin && !isVerified && !isUnverified && (
              <button 
                onClick={onEdit}
                className="p-1.5 text-gray-400 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-all"
                title="Edit KPI"
              >
                <Edit3 size={14} />
              </button>
            )}
            <button 
              onClick={onView}
              className="p-1.5 text-gray-400 hover:text-[#006E62] hover:bg-emerald-50 rounded-lg transition-all"
              title="Lihat Detail"
            >
              <Eye size={14} />
            </button>
          </div>
        </div>
      </div>

      <p className="text-[11px] text-gray-500 line-clamp-3 mb-4 flex-1">"{kpi.description}"</p>

      <div className="space-y-3 pt-4 border-t border-gray-50 mt-auto">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2 text-[10px] font-bold text-gray-400 uppercase">
            <Calendar size={12} />
            <span>{new Date(kpi.deadline + 'T00:00:00').toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' })}</span>
          </div>
          <div className="flex items-center gap-1 text-[10px] font-bold text-[#006E62] uppercase">
            <Target size={12} />
            <span>Bobot: {kpi.weight}</span>
          </div>
        </div>

        {isVerified && kpi.verification_data && (
          <div className="p-3 bg-emerald-50 rounded-xl flex items-center justify-between">
            <div className="flex items-center gap-2">
              <CheckCircle2 size={14} className="text-emerald-600" />
              <span className="text-[10px] font-bold text-emerald-600 uppercase">Skor Final</span>
            </div>
            <span className="text-lg font-bold text-emerald-600">{kpi.verification_data.score}%</span>
          </div>
        )}

        {isUnverified && (
          <div className="p-3 bg-blue-50 rounded-xl flex items-center justify-between">
            <div className="flex items-center gap-2">
              <Clock size={14} className="text-blue-600" />
              <span className="text-[10px] font-bold text-blue-600 uppercase">Menunggu Verifikasi</span>
            </div>
            <span className="text-[10px] font-bold text-blue-400 uppercase">{kpi.report_data?.self_assessment}% Klaim</span>
          </div>
        )}

        <div className="flex gap-2 pt-2">
          {!isAdmin && !isVerified && !isUnverified && (
            <button 
              onClick={onReport}
              className="flex-1 py-2 bg-[#006E62] text-white rounded-lg text-[10px] font-bold uppercase hover:bg-[#005a50] transition-all flex items-center justify-center gap-2"
            >
              <FileText size={14} /> Laporkan
            </button>
          )}
          {isAdmin && isUnverified && (
            <button 
              onClick={onVerify}
              className="flex-1 py-2 bg-[#006E62] text-white rounded-lg text-[10px] font-bold uppercase hover:bg-[#005a50] transition-all flex items-center justify-center gap-2"
            >
              <CheckCircle2 size={14} /> Verifikasi
            </button>
          )}
          {isAdmin && !isVerified && !isUnverified && (
            <button 
              onClick={onDelete}
              className="px-3 py-2 border border-gray-100 text-rose-500 rounded-lg hover:bg-rose-50 transition-all"
            >
              <Trash2 size={14} />
            </button>
          )}
          {(isVerified || isUnverified) && (
            <div className="flex-1 py-2 bg-gray-50 text-gray-400 rounded-lg text-[10px] font-bold uppercase flex items-center justify-center gap-2 cursor-default">
              <Info size={14} /> Detail Tersimpan
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default KPIMain;
