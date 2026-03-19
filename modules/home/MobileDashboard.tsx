
import React, { useState, useEffect } from 'react';
import { 
  Fingerprint, ShieldCheck, Timer, Plane, Calendar, Heart, 
  ClipboardList, Target, CheckSquare, MapPin, Video, MessageSquare, 
  AlertTriangle, Receipt, Wallet, UserCircle, Files, Database,
  Users, CalendarClock, BarChart3, Settings, ArrowLeft, Activity,
  Trophy, Megaphone, ClipboardCheck
} from 'lucide-react';
import { AuthUser, Attendance, Overtime } from '../../types';
import { presenceService } from '../../services/presenceService';
import { overtimeService } from '../../services/overtimeService';
import { googleDriveService } from '../../services/googleDriveService';
import { accountService } from '../../services/accountService';

interface MobileDashboardProps {
  user: AuthUser;
  setActiveTab: (tab: any) => void;
}

const MobileDashboard: React.FC<MobileDashboardProps> = ({ user, setActiveTab }) => {
  const [todayAttendance, setTodayAttendance] = useState<Attendance | null>(null);
  const [todayOvertime, setTodayOvertime] = useState<Overtime | null>(null);
  const [serverTime, setServerTime] = useState<Date>(new Date());
  const [accountDetail, setAccountDetail] = useState<any>(null);
  const [isLoading, setIsLoading] = useState(true);

  useEffect(() => {
    const fetchData = async () => {
      try {
        const [attendance, overtime, sTime, detail] = await Promise.all([
          presenceService.getTodayAttendance(user.id),
          overtimeService.getTodayOvertime(user.id),
          presenceService.getServerTime(),
          accountService.getById(user.id)
        ]);
        setTodayAttendance(attendance);
        setTodayOvertime(overtime);
        setServerTime(sTime);
        setAccountDetail(detail);
      } catch (error) {
        console.error('Error fetching dashboard data:', error);
      } finally {
        setIsLoading(false);
      }
    };

    fetchData();

    const timer = setInterval(() => {
      setServerTime(prev => new Date(prev.getTime() + 1000));
    }, 1000);

    return () => clearInterval(timer);
  }, [user.id]);

  const formatDuration = (startTime: string) => {
    const start = new Date(startTime);
    const diffMs = serverTime.getTime() - start.getTime();
    if (diffMs < 0) return "00:00:00";
    const hours = Math.floor(diffMs / 3600000);
    const minutes = Math.floor((diffMs % 3600000) / 60000);
    const seconds = Math.floor((diffMs % 60000) / 1000);
    return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
  };

  const [viewMode, setViewMode] = useState<'user' | 'admin'>('user');
  const isAdmin = user.role === 'admin' || user.is_hr_admin || user.is_performance_admin || user.is_finance_admin;

  const menuItems = isAdmin ? [
    { id: 'kpi', label: 'Key Performance Indicator', icon: Target, color: 'bg-purple-50 text-purple-600' },
    { id: 'key_activity', label: 'Key Activities', icon: CheckSquare, color: 'bg-teal-50 text-teal-600' },
    { id: 'sales_report', label: 'Sales Report', icon: MapPin, color: 'bg-cyan-50 text-cyan-600' },
    { id: 'rapat', label: 'Rapat', icon: Video, color: 'bg-slate-50 text-slate-600' },
    { id: 'feedback', label: 'Feedback', icon: MessageSquare, color: 'bg-violet-50 text-violet-600' },
    { id: 'lapor', label: 'Laporan Pelanggaran', icon: AlertTriangle, color: 'bg-red-50 text-red-600' },
    { id: 'document', label: 'Dokumen Digital', icon: Files, color: 'bg-blue-50 text-blue-600' },
    { id: 'master_menu', label: 'Dashboard Admin', icon: Database, color: 'bg-gray-100 text-gray-700' },
  ] : [
    { id: 'presence', label: 'Presensi Reguler', icon: Fingerprint, color: 'bg-emerald-50 text-[#006E62]' },
    { id: 'dispensation', label: 'Dispensasi Presensi', icon: ShieldCheck, color: 'bg-blue-50 text-blue-600' },
    { id: 'overtime', label: 'Presensi Lembur', icon: Timer, color: 'bg-orange-50 text-orange-600' },
    { id: 'leave', label: 'Libur Mandiri', icon: Plane, color: 'bg-sky-50 text-sky-600' },
    { id: 'annual_leave', label: 'Cuti Tahunan', icon: Calendar, color: 'bg-indigo-50 text-indigo-600' },
    ...(user.gender === 'Perempuan' ? [{ id: 'maternity_leave', label: 'Cuti Melahirkan', icon: Heart, color: 'bg-rose-50 text-rose-600' }] : []),
    { id: 'permission', label: 'Izin', icon: ClipboardList, color: 'bg-amber-50 text-amber-600' },
    { id: 'kpi', label: 'Key Performance Indicator', icon: Target, color: 'bg-purple-50 text-purple-600' },
    { id: 'key_activity', label: 'Key Activities', icon: CheckSquare, color: 'bg-teal-50 text-teal-600' },
    { id: 'sales_report', label: 'Sales Report', icon: MapPin, color: 'bg-cyan-50 text-cyan-600' },
    { id: 'rapat', label: 'Rapat', icon: Video, color: 'bg-slate-50 text-slate-600' },
    { id: 'feedback', label: 'Feedback', icon: MessageSquare, color: 'bg-violet-50 text-violet-600' },
    { id: 'lapor', label: 'Laporan Pelanggaran', icon: AlertTriangle, color: 'bg-red-50 text-red-600' },
    { id: 'my_payslip', label: 'Slip Gaji Saya', icon: Receipt, color: 'bg-green-50 text-green-600' },
    { id: 'early_salary', label: 'Ambil Gaji Awal', icon: Wallet, color: 'bg-yellow-50 text-yellow-600' },
    { id: 'reimbursement', label: 'Reimburse', icon: Receipt, color: 'bg-pink-50 text-pink-600' },
    { id: 'document', label: 'Dokumen Digital', icon: Files, color: 'bg-blue-50 text-blue-600' },
    ...(isAdmin ? [{ id: 'master_menu', label: 'Dashboard Admin', icon: Database, color: 'bg-gray-100 text-gray-700' }] : []),
  ];

  const adminMenuItems = [
    { id: 'daily_monitoring', label: 'Pemantauan Harian', icon: Activity, color: 'bg-rose-50 text-rose-600' },
    { id: 'admin_dispensation', label: 'Dispensasi sisi admin', icon: ShieldCheck, color: 'bg-cyan-50 text-cyan-600' },
    { id: 'attendance_report', label: 'Laporan Kehadiran', icon: BarChart3, color: 'bg-slate-50 text-slate-600' },
    { id: 'leave', label: 'Libur Mandiri', icon: Plane, color: 'bg-blue-50 text-blue-600' },
    { id: 'overtime', label: 'Presensi Lembur', icon: Timer, color: 'bg-orange-50 text-orange-600' },
    { id: 'permission', label: 'Izin', icon: ClipboardList, color: 'bg-amber-50 text-amber-600' },
    { id: 'annual_leave', label: 'Cuti Tahunan', icon: Calendar, color: 'bg-indigo-50 text-indigo-600' },
    { id: 'reimbursement', label: 'Reimburse', icon: Receipt, color: 'bg-pink-50 text-pink-600' },
    { id: 'early_salary', label: 'Ambil Gaji Awal', icon: Wallet, color: 'bg-yellow-50 text-yellow-600' },
    { id: 'compensation', label: 'Kompensasi', icon: Receipt, color: 'bg-rose-50 text-rose-600' },
    { id: 'finance_report', label: 'Laporan Finance', icon: Wallet, color: 'bg-green-50 text-green-600' },
    { id: 'kpi', label: 'KPI', icon: Target, color: 'bg-purple-50 text-purple-600' },
    { id: 'key_activity', label: 'Key Activities', icon: CheckSquare, color: 'bg-teal-50 text-teal-600' },
    { id: 'sales_report', label: 'Sales report', icon: MapPin, color: 'bg-cyan-50 text-cyan-600' },
    { id: 'feedback', label: 'Feedback Pegawai', icon: MessageSquare, color: 'bg-violet-50 text-violet-600' },
    { id: 'lapor', label: 'Laporan Pelanggaran', icon: AlertTriangle, color: 'bg-red-50 text-red-600' },
    { id: 'employee_of_the_period', label: 'Employee of The Month', icon: Trophy, color: 'bg-orange-50 text-orange-600' },
    { id: 'pengumuman', label: 'Pengumuman', icon: Megaphone, color: 'bg-blue-50 text-blue-600' },
  ];

  const activeWorkSession = todayAttendance?.check_in && !todayAttendance.check_out;
  const activeOvertimeSession = todayOvertime?.check_in && !todayOvertime.check_out;

  return (
    <div className="space-y-6 pb-24">
      {/* Banner Section */}
      <div className="bg-[#006E62] rounded-3xl p-6 text-white shadow-xl relative overflow-hidden">
        <div className="absolute top-0 right-0 w-32 h-32 bg-white/10 rounded-full -mr-16 -mt-16 blur-2xl"></div>
        <div className="absolute bottom-0 left-0 w-24 h-24 bg-[#00FFE4]/10 rounded-full -ml-12 -mb-12 blur-xl"></div>
        
        <div className="relative z-10 flex items-center gap-4">
          <div className="w-16 h-16 rounded-2xl bg-white/20 backdrop-blur-md border border-white/30 overflow-hidden flex items-center justify-center">
            {user.photo_google_id ? (
              <img 
                src={googleDriveService.getFileUrl(user.photo_google_id)} 
                alt={user.full_name} 
                className="w-full h-full object-cover"
                referrerPolicy="no-referrer"
              />
            ) : (
              <UserCircle size={40} className="text-white/70" />
            )}
          </div>
          <div className="flex-1">
            <p className="text-white/70 text-xs font-medium uppercase tracking-wider">Selamat Datang,</p>
            <h2 className="text-xl font-bold truncate">{user.full_name}</h2>
            <div className="flex items-center gap-2 mt-1">
              <span className="text-[10px] bg-white/20 px-2 py-0.5 rounded-full backdrop-blur-sm font-bold uppercase">
                {accountDetail?.position || '-'}
              </span>
              <span className="text-[10px] bg-[#00FFE4]/20 text-[#00FFE4] px-2 py-0.5 rounded-full backdrop-blur-sm font-bold uppercase">
                {accountDetail?.grade || '-'}
              </span>
            </div>
          </div>
        </div>

        {(activeWorkSession || activeOvertimeSession) && (
          <div className="mt-6 pt-6 border-t border-white/10">
            <div className="flex items-center justify-between">
              <div>
                <p className="text-[10px] text-white/60 font-bold uppercase tracking-widest">
                  {activeOvertimeSession ? 'Durasi Lembur Berjalan' : 'Durasi Kerja Berjalan'}
                </p>
                <div className="text-3xl font-mono font-black tracking-tighter mt-1">
                  {activeOvertimeSession 
                    ? formatDuration(todayOvertime!.check_in!) 
                    : formatDuration(todayAttendance!.check_in!)}
                </div>
              </div>
              <div className="w-12 h-12 rounded-full bg-white/10 flex items-center justify-center animate-pulse">
                <Timer size={24} className="text-[#00FFE4]" />
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Grid Menu Section */}
      <div className="space-y-4">
        <div className="flex items-center justify-between">
          <h3 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">
            {viewMode === 'admin' ? 'Dashboard Admin' : 'Menu Layanan'}
          </h3>
          {viewMode === 'admin' && (
            <button 
              onClick={() => setViewMode('user')}
              className="flex items-center gap-1 text-[10px] font-bold text-[#006E62] uppercase tracking-wider"
            >
              <ArrowLeft size={14} /> Kembali
            </button>
          )}
        </div>

        <div className="grid grid-cols-3 sm:grid-cols-4 md:grid-cols-6 gap-4">
          {(viewMode === 'admin' ? adminMenuItems : menuItems).map((item) => (
            <button
              key={item.id}
              onClick={() => {
                if (item.id === 'master_menu') {
                  setViewMode('admin');
                } else {
                  setActiveTab(item.id);
                }
              }}
              className="flex flex-col items-center gap-2 group"
            >
              <div className={`w-14 h-14 ${item.color} rounded-2xl flex items-center justify-center shadow-sm group-active:scale-95 transition-transform duration-200`}>
                <item.icon size={24} />
              </div>
              <span className="text-[10px] font-bold text-gray-600 text-center leading-tight px-1">
                {item.label}
              </span>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
};

export default MobileDashboard;
