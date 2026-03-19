import React, { useState, useEffect } from 'react';
import { MapPin, Plus, Search, Calendar, History, MessageSquare, CheckCircle2, XCircle, Clock, ArrowRight, FileText, User, AlertTriangle, Info, BarChart3, Trash2, Eye, Edit3, CheckSquare, ListTodo, Navigation, Camera } from 'lucide-react';
import Swal from 'sweetalert2';
import { salesReportService } from '../../../services/salesReportService';
import { authService } from '../../../services/authService';
import { SalesReport, AuthUser } from '../../../types';
import SalesReportForm from './SalesReportForm';
import SalesReportDetail from './SalesReportDetail';
import SalesRouteMap from './SalesRouteMap';
import LoadingSpinner from '../../../components/Common/LoadingSpinner';
import { CardSkeleton } from '../../../components/Common/Skeleton';

const SalesReportMain: React.FC = () => {
  const [reports, setReports] = useState<SalesReport[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [showForm, setShowForm] = useState(false);
  const [selectedReport, setSelectedReport] = useState<SalesReport | null>(null);
  const [showDetail, setShowDetail] = useState(false);
  const [showRouteMap, setShowRouteMap] = useState(false);
  const [routeMapReports, setRouteMapReports] = useState<SalesReport[]>([]);
  const [user, setUser] = useState<AuthUser | null>(null);
  const [activeTab, setActiveTab] = useState<'today' | 'history' | 'all'>('today');

  const isAdmin = user?.role === 'admin' || user?.is_performance_admin;

  useEffect(() => {
    const currentUser = authService.getCurrentUser();
    setUser(currentUser);
    if (currentUser?.role === 'admin' || currentUser?.is_performance_admin) {
      setActiveTab('all');
    }
    fetchData(currentUser);
  }, []);

  const fetchData = async (currentUser: AuthUser | null) => {
    if (!currentUser) return;
    try {
      setIsLoading(true);
      const reportsData = (currentUser.role === 'admin' || currentUser.is_performance_admin)
        ? await salesReportService.getAll() 
        : await salesReportService.getByAccountId(currentUser.id);
      setReports(reportsData);
    } catch (error) {
      console.error(error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleSubmitReport = async (input: any) => {
    if (!user) return;
    try {
      setIsSaving(true);
      await salesReportService.submitReport({
        ...input,
        account_id: user.id
      });
      await fetchData(user);
      setShowForm(false);
      Swal.fire({
        title: 'Berhasil!',
        text: 'Laporan kunjungan sales telah dikirim.',
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

  const handleDelete = async (id: string) => {
    const result = await Swal.fire({
      title: 'Hapus Laporan?',
      text: "Data laporan kunjungan akan dihapus permanen.",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
      confirmButtonText: 'Ya, Hapus!'
    });

    if (result.isConfirmed) {
      try {
        setIsSaving(true);
        await salesReportService.delete(id);
        await fetchData(user);
        Swal.fire('Terhapus!', 'Laporan telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Terjadi kesalahan saat menghapus laporan.', 'error');
      } finally {
        setIsSaving(false);
      }
    }
  };

  const today = new Date().toLocaleDateString('en-CA');

  const filteredReports = reports.filter(r => 
    r.customer_name.toLowerCase().includes(searchTerm.toLowerCase()) ||
    r.description.toLowerCase().includes(searchTerm.toLowerCase()) ||
    r.account?.full_name?.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const todayReports = filteredReports.filter(r => r.reported_at.startsWith(today));
  const historyReports = filteredReports.filter(r => !r.reported_at.startsWith(today));

  // Grouping logic for Admin: Date -> Employee -> Reports
  const adminGroupedReports = filteredReports.reduce((acc, report) => {
    const dateKey = new Date(report.reported_at).toLocaleDateString('en-CA');
    const accountId = report.account_id;
    
    if (!acc[dateKey]) {
      acc[dateKey] = {};
    }
    
    if (!acc[dateKey][accountId]) {
      acc[dateKey][accountId] = {
        accountName: report.account?.full_name || 'Unknown',
        reports: []
      };
    }
    
    acc[dateKey][accountId].reports.push(report);
    return acc;
  }, {} as Record<string, Record<string, { accountName: string, reports: SalesReport[] }>>);

  const sortedDates = Object.keys(adminGroupedReports).sort((a, b) => b.localeCompare(a));

  return (
    <div className="space-y-6">
      {isSaving && <LoadingSpinner message="Memproses Data..." />}

      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-xl font-bold text-gray-800 tracking-tight">Sales Report</h2>
          <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Monitoring Kunjungan & Pipeline Sales</p>
        </div>
        
        <div className="flex items-center gap-3">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
            <input
              type="text"
              placeholder="Cari kunjungan..."
              className="pl-9 pr-4 py-2 bg-gray-50 border border-gray-100 rounded-md focus:outline-none focus:ring-2 focus:ring-[#006E62] text-xs w-64"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
          {!isAdmin && (
            <button 
              onClick={() => setShowForm(true)}
              className="flex items-center gap-2 bg-[#006E62] text-white px-4 py-2 rounded-md hover:bg-[#005a50] transition-colors shadow-sm text-xs font-bold uppercase"
            >
              <Plus size={16} />
              Check-in Kunjungan
            </button>
          )}
        </div>
      </div>

      {!isAdmin && (
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div className="bg-white p-5 rounded-2xl border border-gray-100 shadow-sm flex items-center gap-4">
            <div className="w-12 h-12 rounded-xl bg-emerald-50 flex items-center justify-center text-[#006E62]">
              <Navigation size={24} />
            </div>
            <div>
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-0.5">Kunjungan Hari Ini</p>
              <p className="text-2xl font-bold text-[#006E62]">{todayReports.length}</p>
            </div>
          </div>
          <div className="bg-white p-5 rounded-2xl border border-gray-100 shadow-sm flex items-center gap-4">
            <div className="w-12 h-12 rounded-xl bg-blue-50 flex items-center justify-center text-blue-600">
              <History size={24} />
            </div>
            <div>
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-0.5">Total Kunjungan</p>
              <p className="text-2xl font-bold text-blue-600">{reports.length}</p>
            </div>
          </div>
          <div className="bg-white p-5 rounded-2xl border border-gray-100 shadow-sm flex items-center gap-4">
            <div className="w-12 h-12 rounded-xl bg-amber-50 flex items-center justify-center text-amber-600">
              <BarChart3 size={24} />
            </div>
            <div>
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-0.5">Pipeline Aktif</p>
              <p className="text-2xl font-bold text-amber-600">
                {new Set(reports.map(r => r.customer_name)).size} Klien
              </p>
            </div>
          </div>
        </div>
      )}

      {/* Tabs */}
      <div className="flex items-center gap-2 border-b border-gray-100 pb-px">
        {!isAdmin && (
          <>
            <button 
              onClick={() => setActiveTab('today')}
              className={`px-4 py-2 text-[10px] font-bold uppercase tracking-widest transition-all border-b-2 ${activeTab === 'today' ? 'border-[#006E62] text-[#006E62]' : 'border-transparent text-gray-400 hover:text-gray-600'}`}
            >
              Hari Ini ({todayReports.length})
            </button>
            <button 
              onClick={() => setActiveTab('history')}
              className={`px-4 py-2 text-[10px] font-bold uppercase tracking-widest transition-all border-b-2 ${activeTab === 'history' ? 'border-[#006E62] text-[#006E62]' : 'border-transparent text-gray-400 hover:text-gray-600'}`}
            >
              Riwayat Kunjungan
            </button>
          </>
        )}
        {isAdmin && (
          <button 
            onClick={() => setActiveTab('all')}
            className={`px-4 py-2 text-[10px] font-bold uppercase tracking-widest transition-all border-b-2 ${activeTab === 'all' ? 'border-[#006E62] text-[#006E62]' : 'border-transparent text-gray-400 hover:text-gray-600'}`}
          >
            Semua Laporan Sales ({filteredReports.length})
          </button>
        )}
      </div>

      {isLoading ? (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {[1, 2, 3].map(i => <CardSkeleton key={i} />)}
        </div>
      ) : (
        <div className="space-y-8">
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            {activeTab === 'today' && todayReports.map(report => (
              <VisitCard key={report.id} report={report} onView={() => { setSelectedReport(report); setShowDetail(true); }} />
            ))}
            {activeTab === 'history' && historyReports.map(report => (
              <VisitCard key={report.id} report={report} onView={() => { setSelectedReport(report); setShowDetail(true); }} />
            ))}
            {activeTab === 'all' && !isAdmin && filteredReports.map(report => (
              <VisitCard 
                key={report.id} 
                report={report} 
                onView={() => { setSelectedReport(report); setShowDetail(true); }}
                onDelete={() => handleDelete(report.id)}
              />
            ))}
          </div>

          {activeTab === 'all' && isAdmin && (
            <div className="space-y-8">
              {sortedDates.map(dateKey => (
                <div key={dateKey} className="space-y-4">
                  <div className="flex items-center gap-3 px-4 py-2 bg-gray-50 rounded-xl border border-gray-100">
                    <Calendar size={16} className="text-[#006E62]" />
                    <h3 className="text-sm font-bold text-gray-700 uppercase tracking-wider">
                      {new Date(dateKey).toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' })}
                    </h3>
                  </div>
                  
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                    {Object.entries(adminGroupedReports[dateKey] || {}).map(([accountId, data]) => (
                      <div key={accountId} className="bg-white border border-gray-100 rounded-2xl shadow-sm overflow-hidden flex flex-col">
                        <div className="p-4 bg-emerald-50/50 border-b border-gray-100 flex items-center justify-between">
                          <div className="flex items-center gap-2">
                            <div className="w-8 h-8 rounded-full bg-[#006E62] text-white flex items-center justify-center text-xs font-bold">
                              {(data as any).accountName.charAt(0)}
                            </div>
                            <div>
                              <p className="text-xs font-bold text-gray-800 leading-tight">{(data as any).accountName}</p>
                              <p className="text-[9px] text-gray-400 font-bold uppercase tracking-widest">{(data as any).reports.length} Kunjungan</p>
                            </div>
                          </div>
                          <button 
                            onClick={() => {
                              setRouteMapReports((data as any).reports);
                              setShowRouteMap(true);
                            }}
                            className="flex items-center gap-1.5 px-3 py-1.5 bg-[#006E62] text-white rounded-lg text-[9px] font-bold uppercase tracking-wider hover:bg-[#005a50] transition-all shadow-sm"
                          >
                            <Navigation size={12} />
                            Lihat Rute
                          </button>
                        </div>
                        
                        <div className="p-4 space-y-3 flex-1">
                          {((data as any).reports as SalesReport[]).sort((a, b) => new Date(a.reported_at).getTime() - new Date(b.reported_at).getTime()).map((report, idx) => (
                            <div key={report.id} className="flex gap-3 group cursor-pointer" onClick={() => { setSelectedReport(report); setShowDetail(true); }}>
                              <div className="flex flex-col items-center shrink-0">
                                <div className="w-5 h-5 rounded-full border-2 border-[#006E62] bg-white flex items-center justify-center text-[8px] font-bold text-[#006E62] z-10">
                                  {idx + 1}
                                </div>
                                {idx < (data as any).reports.length - 1 && <div className="w-0.5 flex-1 bg-gray-100 my-1"></div>}
                              </div>
                              <div className="pb-2">
                                <div className="flex items-center justify-between gap-2 mb-0.5">
                                  <p className="text-[11px] font-bold text-gray-800 group-hover:text-[#006E62] transition-colors">{report.customer_name}</p>
                                  <span className="text-[9px] text-gray-400 font-medium whitespace-nowrap">
                                    {new Date(report.reported_at).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })}
                                  </span>
                                </div>
                                <p className="text-[10px] text-gray-400 line-clamp-1 italic">"{report.description}"</p>
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          )}

          {((activeTab === 'today' && todayReports.length === 0) || 
            (activeTab === 'history' && historyReports.length === 0) || 
            (activeTab === 'all' && filteredReports.length === 0)) && (
            <div className="flex flex-col items-center justify-center py-20 text-gray-400 bg-gray-50 rounded-xl border border-dashed border-gray-200">
              <MapPin size={48} strokeWidth={1} className="mb-4 opacity-20" />
              <p className="text-sm font-bold uppercase tracking-widest">Belum ada data kunjungan</p>
            </div>
          )}
        </div>
      )}

      {showForm && (
        <SalesReportForm 
          onClose={() => setShowForm(false)} 
          onSubmit={handleSubmitReport}
        />
      )}

      {showDetail && selectedReport && (
        <SalesReportDetail 
          report={selectedReport}
          onClose={() => { setShowDetail(false); setSelectedReport(null); }}
        />
      )}

      {showRouteMap && routeMapReports.length > 0 && (
        <SalesRouteMap 
          reports={routeMapReports}
          onClose={() => { setShowRouteMap(false); setRouteMapReports([]); }}
          onViewDetail={(report) => {
            setSelectedReport(report);
            setShowDetail(true);
          }}
        />
      )}
    </div>
  );
};

const VisitCard: React.FC<{ report: SalesReport, isAdmin?: boolean, onView: () => void, onDelete?: () => void }> = ({ report, isAdmin, onView, onDelete }) => (
  <div className="bg-white border border-gray-100 p-5 rounded-2xl shadow-sm hover:shadow-md transition-all border-l-4 border-l-transparent hover:border-l-[#006E62] group flex flex-col h-full">
    <div className="flex justify-between items-start mb-4">
      <div className="flex flex-col">
        {isAdmin && (
          <div className="flex items-center gap-1.5 mb-1">
            <User size={10} className="text-gray-400" />
            <span className="text-[10px] font-bold text-[#006E62] uppercase">{report.account?.full_name}</span>
          </div>
        )}
        <h4 className="text-sm font-bold text-gray-800 leading-tight group-hover:text-[#006E62] transition-colors">{report.customer_name}</h4>
      </div>
      <div className="flex items-center gap-1">
        <button 
          onClick={onView}
          className="p-1.5 text-gray-400 hover:text-[#006E62] hover:bg-emerald-50 rounded-lg transition-all"
          title="Lihat Detail"
        >
          <Eye size={14} />
        </button>
        {isAdmin && onDelete && (
          <button 
            onClick={onDelete}
            className="p-1.5 text-gray-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-all"
            title="Hapus"
          >
            <Trash2 size={14} />
          </button>
        )}
      </div>
    </div>
    
    <div className="flex items-center gap-2 mb-4">
      <span className="px-2 py-0.5 bg-emerald-50 text-emerald-600 rounded text-[9px] font-bold uppercase">
        {report.activity_type}
      </span>
      <div className="flex items-center gap-1 text-[9px] text-gray-400 font-medium">
        <MapPin size={10} />
        <span>Geotagged</span>
      </div>
    </div>

    <p className="text-[11px] text-gray-500 italic line-clamp-2 mb-4 flex-1">"{report.description}"</p>
    
    <div className="pt-4 border-t border-gray-50 flex items-center justify-between mt-auto">
      <div className="flex items-center gap-2 text-[10px] font-bold text-gray-400 uppercase">
        <Calendar size={12} />
        <span>{new Date(report.reported_at).toLocaleDateString('id-ID', { day: 'numeric', month: 'short' })}</span>
      </div>
      <div className="text-[9px] text-gray-400 font-medium">
        {new Date(report.reported_at).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })} WIB
      </div>
    </div>
  </div>
);

export default SalesReportMain;
