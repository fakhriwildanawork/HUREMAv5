import React, { useState, useEffect } from 'react';
import { Target, Plus, Search, Calendar, History, MessageSquare, CheckCircle2, XCircle, Clock, ArrowRight, FileText, User, AlertTriangle, Info, BarChart3, Trash2, Eye, Edit3, CheckSquare, ListTodo } from 'lucide-react';
import Swal from 'sweetalert2';
import { keyActivityService } from '../../../services/keyActivityService';
import { authService } from '../../../services/authService';
import { KeyActivity, KeyActivityReport, AuthUser } from '../../../types';
import KeyActivityForm from './KeyActivityForm';
import KeyActivityReportForm from './KeyActivityReportForm';
import KeyActivityDetail from './KeyActivityDetail';
import KeyActivityVerifyForm from './KeyActivityVerifyForm';
import LoadingSpinner from '../../../components/Common/LoadingSpinner';
import { CardSkeleton } from '../../../components/Common/Skeleton';

const KeyActivityMain: React.FC = () => {
  const [activities, setActivities] = useState<KeyActivity[]>([]);
  const [reports, setReports] = useState<KeyActivityReport[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [showForm, setShowForm] = useState(false);
  const [selectedActivity, setSelectedActivity] = useState<KeyActivity | null>(null);
  const [selectedDueDate, setSelectedDueDate] = useState<string | null>(null);
  const [showReportForm, setShowReportForm] = useState(false);
  const [showVerifyForm, setShowVerifyForm] = useState(false);
  const [selectedReport, setSelectedReport] = useState<KeyActivityReport | null>(null);
  const [showDetail, setShowDetail] = useState(false);
  const [user, setUser] = useState<AuthUser | null>(null);
  const [activeTab, setActiveTab] = useState<'today' | 'backlog' | 'history' | 'all' | 'verify'>('today');

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
      const [activitiesData, reportsData] = await Promise.all([
        (currentUser.role === 'admin' || currentUser.is_performance_admin) ? keyActivityService.getAll() : keyActivityService.getByAccountId(currentUser.id),
        (currentUser.role === 'admin' || currentUser.is_performance_admin) ? keyActivityService.getAllReports() : keyActivityService.getReportsByAccount(currentUser.id)
      ]);
      setActivities(activitiesData);
      setReports(reportsData);
    } catch (error) {
      console.error(error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleSaveActivity = async (input: any) => {
    try {
      setIsSaving(true);
      if (selectedActivity) {
        await keyActivityService.update(selectedActivity.id, input);
        Swal.fire('Berhasil!', 'Aktivitas telah diperbarui.', 'success');
      } else {
        await keyActivityService.create(input);
        Swal.fire('Berhasil!', 'Aktivitas baru telah dibuat.', 'success');
      }
      await fetchData(user);
      setShowForm(false);
      setSelectedActivity(null);
    } catch (error) {
      Swal.fire('Gagal', 'Terjadi kesalahan saat menyimpan aktivitas.', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleReport = async (reportData: any) => {
    if (!selectedActivity || !selectedDueDate || !user) return;
    try {
      setIsSaving(true);
      await keyActivityService.submitReport({
        ...reportData,
        activity_id: selectedActivity.id,
        account_id: user.id,
        due_date: selectedDueDate
      });
      await fetchData(user);
      setShowReportForm(false);
      setSelectedActivity(null);
      setSelectedDueDate(null);
      Swal.fire('Terkirim!', 'Laporan aktivitas telah dikirim.', 'success');
    } catch (error) {
      Swal.fire('Gagal', 'Terjadi kesalahan saat mengirim laporan.', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleVerify = async (verificationData: any) => {
    if (!selectedReport || !user) return;
    try {
      setIsSaving(true);
      await keyActivityService.verifyReport(selectedReport.id, user.id, verificationData.score, verificationData.notes);
      await fetchData(user);
      setShowVerifyForm(false);
      setSelectedReport(null);
      Swal.fire({
        title: 'Terverifikasi!',
        text: 'Laporan aktivitas telah berhasil diverifikasi.',
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
      title: 'Hapus Aktivitas?',
      text: "Data aktivitas dan semua laporan terkait akan dihapus permanen.",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
      confirmButtonText: 'Ya, Hapus!'
    });

    if (result.isConfirmed) {
      try {
        setIsSaving(true);
        await keyActivityService.delete(id);
        await fetchData(user);
        Swal.fire('Terhapus!', 'Aktivitas telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Terjadi kesalahan saat menghapus aktivitas.', 'error');
      } finally {
        setIsSaving(false);
      }
    }
  };

  const today = new Date().toLocaleDateString('en-CA'); // YYYY-MM-DD in local time

  // Logic to calculate backlog and today's tasks for current user
  const getTasks = () => {
    const backlog: { activity: KeyActivity, dueDate: string }[] = [];
    const todayTasks: { activity: KeyActivity, dueDate: string }[] = [];

    if (isAdmin) return { backlog, todayTasks };

    activities.forEach(activity => {
      const expectedDates = keyActivityService.generateDueDates(activity, today);
      expectedDates.forEach(date => {
        const isReported = reports.some(r => r.activity_id === activity.id && r.due_date === date);
        if (!isReported) {
          if (date < today) {
            backlog.push({ activity, dueDate: date });
          } else if (date === today) {
            todayTasks.push({ activity, dueDate: date });
          }
        }
      });
    });

    return { backlog, todayTasks };
  };

  const { backlog, todayTasks } = getTasks();

  const calculatePerformanceScore = () => {
    const verifiedReports = reports.filter(r => r.status === 'Verified');
    if (verifiedReports.length === 0) return 0;
    
    const totalWeightedScore = verifiedReports.reduce((sum, r) => sum + (r.verification_data?.score || 0) * (r.activity?.weight || 1), 0);
    const totalWeight = verifiedReports.reduce((sum, r) => sum + (r.activity?.weight || 1), 0);
    
    return Math.round(totalWeightedScore / totalWeight);
  };

  const performanceScore = calculatePerformanceScore();

  const filteredActivities = activities.filter(a => 
    a.title.toLowerCase().includes(searchTerm.toLowerCase()) || 
    a.description.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const filteredTodayTasks = todayTasks.filter(t => 
    t.activity.title.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const filteredBacklog = backlog.filter(b => 
    b.activity.title.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const filteredReports = reports.filter(r => 
    r.activity?.title.toLowerCase().includes(searchTerm.toLowerCase()) ||
    r.description.toLowerCase().includes(searchTerm.toLowerCase()) ||
    r.account?.full_name?.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const unverifiedReports = filteredReports.filter(r => r.status === 'Unverified');

  return (
    <div className="space-y-6">
      {isSaving && <LoadingSpinner message="Memproses Data..." />}

      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-xl font-bold text-gray-800 tracking-tight">Key Activities</h2>
          <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Manajemen Aktivitas & Repetisi Pegawai</p>
        </div>
        
        <div className="flex items-center gap-3">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
            <input
              type="text"
              placeholder="Cari aktivitas..."
              className="pl-9 pr-4 py-2 bg-gray-50 border border-gray-100 rounded-md focus:outline-none focus:ring-2 focus:ring-[#006E62] text-xs w-64"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
          {isAdmin && (
            <button 
              onClick={() => { setSelectedActivity(null); setShowForm(true); }}
              className="flex items-center gap-2 bg-[#006E62] text-white px-4 py-2 rounded-md hover:bg-[#005a50] transition-colors shadow-sm text-xs font-bold uppercase"
            >
              <Plus size={16} />
              Buat Aktivitas
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
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-0.5">Backlog Tasks</p>
              <p className="text-2xl font-bold text-rose-600">{backlog.length}</p>
            </div>
          </div>
          <div className="bg-white p-5 rounded-2xl border border-gray-100 shadow-sm flex items-center gap-4">
            <div className="w-12 h-12 rounded-xl bg-amber-50 flex items-center justify-center text-amber-600">
              <Clock size={24} />
            </div>
            <div>
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-0.5">Today Tasks</p>
              <p className="text-2xl font-bold text-amber-600">{todayTasks.length}</p>
            </div>
          </div>
          <div className="bg-white p-5 rounded-2xl border border-gray-100 shadow-sm flex items-center gap-4">
            <div className="w-12 h-12 rounded-xl bg-blue-50 flex items-center justify-center text-blue-600">
              <History size={24} />
            </div>
            <div>
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-0.5">Total Activities</p>
              <p className="text-2xl font-bold text-blue-600">{activities.length}</p>
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
              Hari Ini ({todayTasks.length})
            </button>
            <button 
              onClick={() => setActiveTab('backlog')}
              className={`px-4 py-2 text-[10px] font-bold uppercase tracking-widest transition-all border-b-2 ${activeTab === 'backlog' ? 'border-rose-500 text-rose-500' : 'border-transparent text-gray-400 hover:text-gray-600'}`}
            >
              Backlog ({backlog.length})
            </button>
            <button 
              onClick={() => setActiveTab('history')}
              className={`px-4 py-2 text-[10px] font-bold uppercase tracking-widest transition-all border-b-2 ${activeTab === 'history' ? 'border-[#006E62] text-[#006E62]' : 'border-transparent text-gray-400 hover:text-gray-600'}`}
            >
              Riwayat Laporan
            </button>
          </>
        )}
        {isAdmin && (
          <>
            <button 
              onClick={() => setActiveTab('all')}
              className={`px-4 py-2 text-[10px] font-bold uppercase tracking-widest transition-all border-b-2 ${activeTab === 'all' ? 'border-[#006E62] text-[#006E62]' : 'border-transparent text-gray-400 hover:text-gray-600'}`}
            >
              Semua Aktivitas
            </button>
            <button 
              onClick={() => setActiveTab('verify')}
              className={`px-4 py-2 text-[10px] font-bold uppercase tracking-widest transition-all border-b-2 ${activeTab === 'verify' ? 'border-blue-500 text-blue-500' : 'border-transparent text-gray-400 hover:text-gray-600'}`}
            >
              Verifikasi Laporan ({unverifiedReports.length})
            </button>
            <button 
              onClick={() => setActiveTab('history')}
              className={`px-4 py-2 text-[10px] font-bold uppercase tracking-widest transition-all border-b-2 ${activeTab === 'history' ? 'border-[#006E62] text-[#006E62]' : 'border-transparent text-gray-400 hover:text-gray-600'}`}
            >
              Riwayat Masuk
            </button>
          </>
        )}
      </div>

      {isLoading ? (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {[1, 2, 3].map(i => <CardSkeleton key={i} />)}
        </div>
      ) : (
        <div className="space-y-8">
          {/* Section: Urgent Attention (User) */}
          {!isAdmin && (todayTasks.length > 0 || backlog.length > 0) && (activeTab === 'today' || activeTab === 'backlog') && (
            <div className="space-y-4">
              <div className="flex items-center gap-2">
                <AlertTriangle size={16} className="text-rose-500" />
                <h3 className="text-xs font-bold text-gray-800 uppercase tracking-widest">Butuh Perhatian Segera</h3>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                {activeTab === 'today' && filteredTodayTasks.map(task => (
                  <TaskCard 
                    key={`${task.activity.id}-${task.dueDate}`}
                    activity={task.activity}
                    dueDate={task.dueDate}
                    onReport={() => { setSelectedActivity(task.activity); setSelectedDueDate(task.dueDate); setShowReportForm(true); }}
                    onView={() => { setSelectedActivity(task.activity); setShowDetail(true); }}
                  />
                ))}
                {activeTab === 'backlog' && filteredBacklog.map(task => (
                  <TaskCard 
                    key={`${task.activity.id}-${task.dueDate}`}
                    activity={task.activity}
                    dueDate={task.dueDate}
                    isBacklog
                    onReport={() => { setSelectedActivity(task.activity); setSelectedDueDate(task.dueDate); setShowReportForm(true); }}
                    onView={() => { setSelectedActivity(task.activity); setShowDetail(true); }}
                  />
                ))}
              </div>
            </div>
          )}

          {/* Section: Main Content */}
          <div className="space-y-4">
            <div className="flex items-center gap-2">
              <Target size={16} className="text-[#006E62]" />
              <h3 className="text-xs font-bold text-gray-800 uppercase tracking-widest">
                {isAdmin ? 'Semua Daftar Key Activity Pegawai' : 'Daftar Aktivitas & Riwayat'}
              </h3>
            </div>
            
            {activeTab === 'history' && filteredReports.length === 0 && (
              <div className="flex flex-col items-center justify-center py-20 text-gray-400 bg-gray-50 rounded-xl border border-dashed border-gray-200">
                <History size={48} strokeWidth={1} className="mb-4 opacity-20" />
                <p className="text-sm font-bold uppercase tracking-widest">Belum ada riwayat laporan</p>
              </div>
            )}

            {activeTab === 'all' && filteredActivities.length === 0 && (
              <div className="flex flex-col items-center justify-center py-20 text-gray-400 bg-gray-50 rounded-xl border border-dashed border-gray-200">
                <Target size={48} strokeWidth={1} className="mb-4 opacity-20" />
                <p className="text-sm font-bold uppercase tracking-widest">Belum ada data aktivitas</p>
              </div>
            )}

            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
              {activeTab === 'history' && filteredReports.map(report => (
                <ReportCard key={report.id} report={report} isAdmin={isAdmin} onVerify={() => { setSelectedReport(report); setShowVerifyForm(true); }} />
              ))}
              {activeTab === 'verify' && unverifiedReports.map(report => (
                <ReportCard key={report.id} report={report} isAdmin={isAdmin} onVerify={() => { setSelectedReport(report); setShowVerifyForm(true); }} />
              ))}
              {activeTab === 'all' && filteredActivities.map(activity => (
                <ActivityCard 
                  key={activity.id} 
                  activity={activity} 
                  onEdit={() => { setSelectedActivity(activity); setShowForm(true); }}
                  onView={() => { setSelectedActivity(activity); setShowDetail(true); }}
                  onDelete={() => handleDelete(activity.id)}
                />
              ))}
              {/* Fallback for empty Today/Backlog if not shown in urgent section */}
              {activeTab === 'today' && todayTasks.length === 0 && (
                <div className="col-span-full flex flex-col items-center justify-center py-20 text-gray-400 bg-gray-50 rounded-xl border border-dashed border-gray-200">
                  <CheckCircle2 size={48} strokeWidth={1} className="mb-4 opacity-20" />
                  <p className="text-sm font-bold uppercase tracking-widest">Semua tugas hari ini sudah selesai!</p>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {showForm && (
        <KeyActivityForm 
          onClose={() => { setShowForm(false); setSelectedActivity(null); }} 
          onSubmit={handleSaveActivity}
          initialData={selectedActivity}
        />
      )}

      {showReportForm && selectedActivity && selectedDueDate && (
        <KeyActivityReportForm 
          activityTitle={selectedActivity.title}
          dueDate={selectedDueDate}
          onClose={() => { setShowReportForm(false); setSelectedActivity(null); setSelectedDueDate(null); }}
          onSubmit={handleReport}
        />
      )}

      {showVerifyForm && selectedReport && (
        <KeyActivityVerifyForm 
          report={selectedReport}
          onClose={() => { setShowVerifyForm(false); setSelectedReport(null); }}
          onSubmit={handleVerify}
        />
      )}

      {showDetail && selectedActivity && (
        <KeyActivityDetail 
          activity={selectedActivity}
          reports={reports.filter(r => r.activity_id === selectedActivity.id)}
          isAdmin={isAdmin}
          onVerify={(report) => { setSelectedReport(report); setShowVerifyForm(true); }}
          onClose={() => { setShowDetail(false); setSelectedActivity(null); }}
        />
      )}
    </div>
  );
};

const TaskCard: React.FC<{ activity: KeyActivity, dueDate: string, isBacklog?: boolean, onReport: () => void, onView: () => void }> = ({ activity, dueDate, isBacklog, onReport, onView }) => (
  <div className={`bg-white border p-5 rounded-2xl shadow-sm hover:shadow-md transition-all border-l-4 ${isBacklog ? 'border-l-rose-500 hover:border-l-rose-600' : 'border-l-[#006E62] hover:border-l-[#005a50]'} flex flex-col h-full group`}>
    <div className="flex justify-between items-start mb-4">
      <div className="flex flex-col">
        <div className="flex items-center gap-1.5 mb-1">
          <Clock size={10} className={isBacklog ? 'text-rose-500' : 'text-gray-400'} />
          <span className={`text-[10px] font-bold uppercase ${isBacklog ? 'text-rose-500' : 'text-gray-400'}`}>
            {isBacklog ? 'Terlewat' : 'Hari Ini'}
          </span>
        </div>
        <h4 className="text-sm font-bold text-gray-800 leading-tight group-hover:text-[#006E62] transition-colors">{activity.title}</h4>
      </div>
      <div className="flex flex-col items-end gap-2">
        <span className={`px-2 py-0.5 rounded text-[9px] font-bold uppercase ${isBacklog ? 'bg-rose-50 text-rose-600' : 'bg-emerald-50 text-emerald-600'}`}>
          {activity.recurrence_type}
        </span>
        <button 
          onClick={onView}
          className="p-1.5 text-gray-400 hover:text-[#006E62] hover:bg-emerald-50 rounded-lg transition-all"
          title="Lihat Detail"
        >
          <Eye size={14} />
        </button>
      </div>
    </div>
    <p className="text-[11px] text-gray-500 line-clamp-2 mb-4 flex-1">"{activity.description}"</p>
    <div className="pt-4 border-t border-gray-50 flex items-center justify-between mt-auto">
      <div className="flex items-center gap-2 text-[10px] font-bold text-gray-400 uppercase">
        <Calendar size={12} />
        <span>{new Date(dueDate + 'T00:00:00').toLocaleDateString('id-ID', { day: 'numeric', month: 'short' })}</span>
      </div>
      <button 
        onClick={onReport}
        className="px-4 py-2 bg-[#006E62] text-white rounded-lg text-[10px] font-bold uppercase hover:bg-[#005a50] transition-all flex items-center gap-2"
      >
        <FileText size={14} /> Lapor
      </button>
    </div>
  </div>
);

const ReportCard: React.FC<{ report: KeyActivityReport, isAdmin?: boolean, onVerify?: () => void }> = ({ report, isAdmin, onVerify }) => (
  <div className="bg-white border border-gray-100 p-5 rounded-2xl shadow-sm hover:shadow-md transition-all border-l-4 border-l-transparent hover:border-l-[#006E62] group flex flex-col h-full">
    <div className="flex justify-between items-start mb-4">
      <div className="flex flex-col">
        {isAdmin && (
          <div className="flex items-center gap-1.5 mb-1">
            <User size={10} className="text-gray-400" />
            <span className="text-[10px] font-bold text-[#006E62] uppercase">{report.account?.full_name}</span>
          </div>
        )}
        <h4 className="text-sm font-bold text-gray-800 leading-tight group-hover:text-[#006E62] transition-colors">{report.activity?.title}</h4>
      </div>
      <span className={`px-2 py-0.5 rounded text-[9px] font-bold uppercase ${report.status === 'Verified' ? 'bg-emerald-50 text-emerald-600' : 'bg-blue-50 text-blue-600'}`}>
        {report.status}
      </span>
    </div>
    <p className="text-[11px] text-gray-500 italic line-clamp-2 mb-4 flex-1">"{report.description}"</p>
    
    {report.status === 'Verified' && report.verification_data && (
      <div className="mb-4 p-2 bg-emerald-50 rounded-xl flex items-center justify-between">
        <span className="text-[9px] font-bold text-emerald-600 uppercase">Skor: {report.verification_data.score}%</span>
        <CheckCircle2 size={12} className="text-emerald-600" />
      </div>
    )}

    <div className="pt-4 border-t border-gray-50 flex items-center justify-between mt-auto">
      <div className="flex items-center gap-2 text-[10px] font-bold text-gray-400 uppercase">
        <Calendar size={12} />
        <span>Due: {new Date(report.due_date + 'T00:00:00').toLocaleDateString('id-ID', { day: 'numeric', month: 'short' })}</span>
      </div>
      {isAdmin && report.status === 'Unverified' ? (
        <button 
          onClick={onVerify}
          className="px-3 py-1.5 bg-[#006E62] text-white rounded-lg text-[9px] font-bold uppercase hover:bg-[#005a50] transition-all flex items-center gap-1.5"
        >
          <CheckCircle2 size={12} /> Verifikasi
        </button>
      ) : (
        <div className="text-[9px] text-gray-400 font-medium">
          Dilapor: {new Date(report.reported_at).toLocaleDateString('id-ID')}
        </div>
      )}
    </div>
  </div>
);

const ActivityCard: React.FC<{ activity: KeyActivity, onEdit: () => void, onView: () => void, onDelete: () => void }> = ({ activity, onEdit, onView, onDelete }) => (
  <div className="bg-white border border-gray-100 p-5 rounded-2xl shadow-sm hover:shadow-md transition-all border-l-4 border-l-transparent hover:border-l-[#006E62] group flex flex-col h-full">
    <div className="flex justify-between items-start mb-4">
      <h4 className="text-sm font-bold text-gray-800 leading-tight group-hover:text-[#006E62] transition-colors">{activity.title}</h4>
      <div className="flex items-center gap-1">
        <button onClick={onEdit} className="p-1.5 text-gray-400 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-all" title="Edit"><Edit3 size={14} /></button>
        <button onClick={onView} className="p-1.5 text-gray-400 hover:text-[#006E62] hover:bg-emerald-50 rounded-lg transition-all" title="Lihat Detail"><Eye size={14} /></button>
        <button onClick={onDelete} className="p-1.5 text-gray-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-all" title="Hapus"><Trash2 size={14} /></button>
      </div>
    </div>
    <p className="text-[11px] text-gray-500 line-clamp-2 mb-4 flex-1">"{activity.description}"</p>
    <div className="pt-4 border-t border-gray-50 flex items-center justify-between mt-auto">
      <div className="flex items-center gap-2 text-[10px] font-bold text-[#006E62] uppercase">
        <Target size={12} />
        <span>Bobot: {activity.weight}</span>
      </div>
      <span className="text-[9px] font-bold text-gray-400 uppercase tracking-widest">{activity.recurrence_type}</span>
    </div>
  </div>
);

export default KeyActivityMain;
