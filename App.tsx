import React, { useState, useEffect, Suspense, lazy } from 'react';
import { X, LayoutDashboard, Users, MapPin, CalendarClock, Files, Settings, Database, Fingerprint, Timer, ClipboardCheck, Plane, Calendar, ClipboardList, Heart, Target, CheckSquare, AlertTriangle, Video, Megaphone, Receipt, Trophy, BarChart3, Wallet, AlertCircle, Activity } from 'lucide-react';
import Sidebar from './components/Layout/Sidebar';
import Header from './components/Layout/Header';
import MobileLayout from './components/Layout/MobileLayout';
import MobileDashboard from './modules/home/MobileDashboard';

// Lazy load modules for performance optimization
const LocationMain = lazy(() => import('./modules/location/LocationMain'));
const AccountMain = lazy(() => import('./modules/account/AccountMain'));
const ScheduleMain = lazy(() => import('./modules/schedule/ScheduleMain'));
const DocumentMain = lazy(() => import('./modules/document/DocumentMain'));
const PresenceMain = lazy(() => import('./modules/presence/PresenceMain'));
const OvertimeMain = lazy(() => import('./modules/overtime/OvertimeMain'));
const SubmissionMain = lazy(() => import('./modules/submission/SubmissionMain'));
const LeaveMain = lazy(() => import('./modules/leave/LeaveMain'));
const AnnualLeaveMain = lazy(() => import('./modules/leave/AnnualLeaveMain'));
const PermissionMain = lazy(() => import('./modules/permission/PermissionMain'));
const MaternityLeaveMain = lazy(() => import('./modules/maternity/MaternityLeaveMain'));
const KPIMain = lazy(() => import('./modules/performance/kpi/KPIMain'));
const KeyActivityMain = lazy(() => import('./modules/performance/key-activity/KeyActivityMain'));
const EmployeeOfThePeriodMain = lazy(() => import('./modules/performance/award/EmployeeOfThePeriodMain'));
const SalesReportMain = lazy(() => import('./modules/performance/sales-report/SalesReportMain'));
const FeedbackMain = lazy(() => import('./modules/feedback/FeedbackMain'));
const LaporMain = lazy(() => import('./modules/lapor/LaporMain'));
const RapatMain = lazy(() => import('./modules/rapat/RapatMain'));
const PengumumanMain = lazy(() => import('./modules/pengumuman/PengumumanMain'));
const SalarySchemeMain = lazy(() => import('./modules/finance/SalarySchemeMain'));
const SalaryAdjustmentMain = lazy(() => import('./modules/finance/SalaryAdjustmentMain'));
const PayrollMain = lazy(() => import('./modules/finance/PayrollMain'));
const MyPayslip = lazy(() => import('./modules/finance/MyPayslip'));
const ReimbursementMain = lazy(() => import('./modules/finance/ReimbursementMain'));
const EarlySalaryMain = lazy(() => import('./modules/finance/EarlySalaryModule'));
const CompensationMain = lazy(() => import('./modules/finance/CompensationMain'));
const DispensationMain = lazy(() => import('./modules/dispensation/DispensationMain'));
const AdminDispensationMain = lazy(() => import('./modules/dispensation/AdminDispensationMain'));
const AdminSubmissionModule = lazy(() => import('./modules/admin/AdminSubmissionModule'));
const AttendanceReportMain = lazy(() => import('./modules/report/AttendanceReportMain'));
const EmployeeReportMain = lazy(() => import('./modules/report/EmployeeReportMain'));
const ReportMainModule = lazy(() => import('./modules/report/ReportMainModule'));
const FinanceReportMain = lazy(() => import('./modules/report/FinanceReportMain'));
const MasterMain = lazy(() => import('./modules/settings/MasterMain'));
const AdminSettingsModule = lazy(() => import('./modules/settings/AdminSettingsModule'));
const DailyMonitoring = lazy(() => import('./modules/monitoring/DailyMonitoring'));
const Login = lazy(() => import('./modules/auth/Login'));

import { authService } from './services/authService';
import { AuthUser } from './types';

const App: React.FC = () => {
  const [user, setUser] = useState<AuthUser | null>(null);
  const currentUser = authService.getCurrentUser();
  const isAdminInitial = currentUser?.role === 'admin' || currentUser?.is_hr_admin || currentUser?.is_performance_admin || currentUser?.is_finance_admin;

  const [activeTab, setActiveTab] = useState<'dashboard' | 'location' | 'account' | 'schedule' | 'document' | 'settings' | 'presence' | 'overtime' | 'submission' | 'leave' | 'annual_leave' | 'permission' | 'maternity_leave' | 'master_app' | 'admin_settings' | 'kpi' | 'key_activity' | 'sales_report' | 'feedback' | 'lapor' | 'rapat' | 'pengumuman' | 'salary_scheme' | 'salary_adjustment' | 'payroll' | 'my_payslip' | 'reimbursement' | 'early_salary' | 'compensation' | 'employee_of_the_period' | 'dispensation' | 'admin_dispensation' | 'attendance_report' | 'finance_report' | 'employee_report' | 'daily_monitoring'>(
    (window.innerWidth < 768) 
      ? 'dashboard' 
      : (isAdminInitial ? 'master_app' : 'presence')
  );
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false);
  const [isAuthChecking, setIsAuthChecking] = useState(true);
  const [isMobile, setIsMobile] = useState(window.innerWidth < 768);

  const isAdmin = user?.role === 'admin' || user?.is_hr_admin || user?.is_performance_admin || user?.is_finance_admin;

  useEffect(() => {
    const handleResize = () => setIsMobile(window.innerWidth < 768);
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  useEffect(() => {
    const refreshUserPermissions = async () => {
      const currentUser = authService.getCurrentUser();
      if (currentUser) {
        try {
          // Fetch latest permissions
          const { settingsService } = await import('./services/settingsService');
          const [hrAdmins, perfAdmins, finAdmins] = await Promise.all([
            settingsService.getSetting('admin_hr_ids', []),
            settingsService.getSetting('admin_performance_ids', []),
            settingsService.getSetting('admin_finance_ids', [])
          ]);

          const updatedUser = {
            ...currentUser,
            is_hr_admin: hrAdmins.includes(currentUser.id),
            is_performance_admin: perfAdmins.includes(currentUser.id),
            is_finance_admin: finAdmins.includes(currentUser.id)
          };
          
          setUser(updatedUser);
          // Update localStorage to keep it in sync
          localStorage.setItem('hurema_user_session', JSON.stringify(updatedUser));
        } catch (error) {
          console.error('Error refreshing user permissions:', error);
          setUser(currentUser);
        }
      }
      setIsAuthChecking(false);
    };

    refreshUserPermissions();
  }, []);

  if (isAuthChecking) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-white">
        <div className="w-12 h-12 border-4 border-[#006E62] border-t-transparent rounded-full animate-spin"></div>
      </div>
    );
  }

  if (!user) {
    return (
      <Suspense fallback={
        <div className="min-h-screen flex items-center justify-center bg-white">
          <div className="w-12 h-12 border-4 border-[#006E62] border-t-transparent rounded-full animate-spin"></div>
        </div>
      }>
        <Login onLoginSuccess={(u) => setUser(u)} />
      </Suspense>
    );
  }

  const NavItemMobile = ({ id, icon: Icon, label, indent = false }: { id: any, icon: any, label: string, indent?: boolean }) => (
    <button
      onClick={() => { setActiveTab(id); setIsMobileMenuOpen(false); }}
      className={`flex items-center gap-3 px-4 py-3 rounded-md transition-all duration-200 w-full mb-1 ${
        activeTab === id 
          ? 'bg-[#006E62] text-white shadow-md' 
          : 'text-gray-600 hover:bg-gray-100'
      } ${indent ? 'ml-4 w-[calc(100%-1rem)]' : ''}`}
    >
      <Icon size={20} />
      <span className="font-medium text-sm">{label}</span>
    </button>
  );

  // Mobile Non-Admin Layout
  if (isMobile && !isAdmin) {
    return (
      <MobileLayout activeTab={activeTab} setActiveTab={setActiveTab} user={user}>
        <Suspense fallback={
          <div className="flex flex-col items-center justify-center h-64">
            <div className="w-10 h-10 border-4 border-[#006E62] border-t-transparent rounded-full animate-spin"></div>
            <p className="mt-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Memuat...</p>
          </div>
        }>
          {activeTab === 'dashboard' ? (
            <MobileDashboard user={user} setActiveTab={setActiveTab} />
          ) : activeTab === 'location' ? (
            <LocationMain />
          ) : activeTab === 'account' ? (
            <AccountMain user={user} setUser={setUser} />
          ) : activeTab === 'schedule' ? (
            <ScheduleMain />
          ) : activeTab === 'document' ? (
            <DocumentMain user={user} />
          ) : activeTab === 'presence' ? (
            <PresenceMain />
          ) : activeTab === 'overtime' ? (
            <OvertimeMain />
          ) : activeTab === 'submission' ? (
            <SubmissionMain />
          ) : activeTab === 'leave' ? (
            <LeaveMain />
          ) : activeTab === 'annual_leave' ? (
            <AnnualLeaveMain />
          ) : activeTab === 'permission' ? (
            <PermissionMain />
          ) : activeTab === 'maternity_leave' ? (
            <MaternityLeaveMain />
          ) : activeTab === 'kpi' ? (
            <KPIMain />
          ) : activeTab === 'key_activity' ? (
            <KeyActivityMain />
          ) : activeTab === 'employee_of_the_period' ? (
            <EmployeeOfThePeriodMain />
          ) : activeTab === 'sales_report' ? (
            <SalesReportMain />
          ) : activeTab === 'feedback' ? (
            <FeedbackMain />
          ) : activeTab === 'lapor' ? (
            <LaporMain />
          ) : activeTab === 'rapat' ? (
            <RapatMain />
          ) : activeTab === 'pengumuman' ? (
            <PengumumanMain user={user} />
          ) : activeTab === 'salary_scheme' ? (
            <SalarySchemeMain />
          ) : activeTab === 'salary_adjustment' ? (
            <SalaryAdjustmentMain />
          ) : activeTab === 'payroll' ? (
            <PayrollMain />
          ) : activeTab === 'my_payslip' ? (
            <MyPayslip />
          ) : activeTab === 'reimbursement' ? (
            <ReimbursementMain />
          ) : activeTab === 'early_salary' ? (
            <EarlySalaryMain />
          ) : activeTab === 'compensation' ? (
            <CompensationMain />
          ) : activeTab === 'dispensation' ? (
            <DispensationMain user={user} />
          ) : activeTab === 'admin_dispensation' ? (
            <AdminDispensationMain user={user} />
          ) : activeTab === 'attendance_report' ? (
            <AttendanceReportMain />
          ) : activeTab === 'finance_report' ? (
            <FinanceReportMain />
          ) : activeTab === 'master_app' ? (
            <MasterMain />
          ) : activeTab === 'settings' ? (
            isAdmin ? <AdminSettingsModule /> : <AccountMain user={user} setUser={setUser} isSelfProfile={true} />
          ) : (
            <div className="flex flex-col items-center justify-center h-64 text-gray-400 bg-gray-50 rounded-lg border border-dashed border-gray-200">
              <p className="font-medium text-sm">Modul "{activeTab}" sedang dalam pengembangan.</p>
            </div>
          )}
        </Suspense>
      </MobileLayout>
    );
  }

  // Desktop Non-Admin Restriction
  if (!isMobile && !isAdmin) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-white p-4">
        <div className="text-center max-w-md animate-in fade-in zoom-in duration-500">
          <div className="w-20 h-20 bg-rose-50 rounded-3xl flex items-center justify-center text-rose-500 mx-auto mb-8 shadow-xl shadow-rose-500/10">
            <AlertCircle size={40} />
          </div>
          <h1 className="text-2xl font-black text-gray-800 uppercase tracking-tighter mb-3">Akses Terbatas</h1>
          <p className="text-[10px] font-bold text-gray-400 uppercase tracking-[0.2em] leading-relaxed mb-8">
            Halaman Dashboard Desktop hanya dapat diakses oleh Admin.<br />
            Silakan gunakan aplikasi mobile untuk akses karyawan.
          </p>
          <button 
            onClick={() => authService.logout()}
            className="px-10 py-4 bg-[#006E62] text-white rounded-xl font-black text-[10px] uppercase tracking-[0.2em] hover:bg-[#005a50] shadow-lg shadow-[#006E62]/20 active:scale-95 transition-all"
          >
            Keluar Sesi
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen flex bg-white text-gray-800">
      <Sidebar 
        activeTab={activeTab} 
        setActiveTab={setActiveTab} 
        isCollapsed={isSidebarCollapsed} 
        setIsCollapsed={setIsSidebarCollapsed} 
      />

      {/* Mobile Menu Overlay */}
      {isMobileMenuOpen && (
        <div className="fixed inset-0 bg-black bg-opacity-50 z-40 md:hidden" onClick={() => setIsMobileMenuOpen(false)} />
      )}
      
      {/* Sidebar Mobile */}
      <aside className={`fixed inset-y-0 left-0 w-64 bg-white z-50 transform transition-transform duration-300 md:hidden ${isMobileMenuOpen ? 'translate-x-0' : '-translate-x-full'}`}>
        <div className="p-4 h-full flex flex-col">
          <div className="flex items-center justify-between mb-8 px-2">
            <div className="flex items-center gap-2">
              <div className="w-8 h-8 bg-[#006E62] rounded flex items-center justify-center text-white font-bold italic">H</div>
              <h1 className="text-xl font-bold text-[#006E62]">HUREMA</h1>
            </div>
            <button onClick={() => setIsMobileMenuOpen(false)}><X size={24} className="text-gray-400" /></button>
          </div>
          <nav className="flex-1 space-y-1 overflow-y-auto scrollbar-none">
            {!isAdmin && <NavItemMobile id="dashboard" icon={LayoutDashboard} label="Beranda" />}
            
            <div className="flex items-center gap-3 px-4 py-3 text-gray-400 mt-2">
              <Database size={20} />
              <span className="font-bold text-[10px] uppercase tracking-widest">{isAdmin ? 'Dashboard Admin' : 'Master'}</span>
            </div>
            {!isAdmin ? (
              <>
                <NavItemMobile id="master_app" icon={Database} label="Master Aplikasi" indent />
                <NavItemMobile id="location" icon={MapPin} label="Data Lokasi" indent />
                <NavItemMobile id="schedule" icon={CalendarClock} label="Manajemen Jadwal" indent />
                <NavItemMobile id="account" icon={Users} label="Akun" indent />
              </>
            ) : null}

            <div className="mt-4">
              {!isAdmin && (
                <>
                  <NavItemMobile id="presence" icon={Fingerprint} label="Presensi Reguler" />
                  <NavItemMobile id="overtime" icon={Timer} label="Presensi Lembur" />
                </>
              )}
              <NavItemMobile id="kpi" icon={Target} label={isAdmin ? 'KPI' : 'KPI Performance'} />
              <NavItemMobile id="key_activity" icon={CheckSquare} label="Key Activities" />
              <NavItemMobile id="employee_of_the_period" icon={Trophy} label={isAdmin ? 'Employee of The Month' : 'Employee of The Period'} />
              <NavItemMobile id="sales_report" icon={MapPin} label={isAdmin ? 'Sales report' : 'Sales Report'} />
              <NavItemMobile id="feedback" icon={ClipboardList} label="Feedback Pegawai" />
              <NavItemMobile id="lapor" icon={AlertTriangle} label={isAdmin ? 'Laporan Pelanggaran' : 'Lapor Pelanggaran'} />
              {!isAdmin && <NavItemMobile id="rapat" icon={Video} label="Notulensi Rapat" />}
              <NavItemMobile id="pengumuman" icon={Megaphone} label="Pengumuman" />

              {!isAdmin && (
                <div className="flex items-center gap-3 px-4 py-3 text-gray-400 mt-2">
                  <Receipt size={20} />
                  <span className="font-bold text-[10px] uppercase tracking-widest">Finance</span>
                </div>
              )}
              {!isAdmin && <NavItemMobile id="salary_scheme" icon={Receipt} label="Master Skema Gaji" indent />}
              {isAdmin && false && (
                <NavItemMobile id="salary_adjustment" icon={Receipt} label="Kustom Gaji" indent />
              )}
              <NavItemMobile id="reimbursement" icon={Receipt} label="Reimburse" indent={!isAdmin} />
              <NavItemMobile id="early_salary" icon={Receipt} label="Ambil Gaji Awal" indent={!isAdmin} />
              {isAdmin && (
                <NavItemMobile id="compensation" icon={Receipt} label="Kompensasi" indent={false} />
              )}

              {(isAdmin || user?.is_hr_admin) && (
                <NavItemMobile id="daily_monitoring" icon={Activity} label="Pemantauan Harian" />
              )}
              
              <div className="flex items-center gap-3 px-4 py-3 text-gray-400 mt-2">
                <ClipboardCheck size={20} />
                <span className="font-bold text-[10px] uppercase tracking-widest">Pengajuan</span>
              </div>
              <NavItemMobile id="leave" icon={Plane} label="Libur Mandiri" indent />
              <NavItemMobile id="overtime" icon={Timer} label="Presensi Lembur" indent />
              <NavItemMobile id="permission" icon={ClipboardList} label="Izin" indent />
              <NavItemMobile id="annual_leave" icon={Calendar} label="Cuti Tahunan" indent />
              {user?.gender === 'Perempuan' && (
                <NavItemMobile id="maternity_leave" icon={Heart} label="Cuti Melahirkan" indent />
              )}
              {isAdmin && (
                <NavItemMobile id="admin_dispensation" icon={ClipboardCheck} label="Dispensasi sisi admin" indent />
              )}

              {!isAdmin && <NavItemMobile id="document" icon={Files} label="Dokumen Digital" />}
              {!isAdmin && <NavItemMobile id="employee_report" icon={BarChart3} label="Laporan Karyawan" />}
              <NavItemMobile id="attendance_report" icon={BarChart3} label="Laporan Kehadiran" />
              <NavItemMobile id="finance_report" icon={Wallet} label="Laporan Finance" />
              {!isAdmin && <NavItemMobile id="settings" icon={Settings} label="Pengaturan" />}
            </div>
          </nav>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 min-w-0 flex flex-col">
        <Header activeTab={activeTab} onMenuClick={() => setIsMobileMenuOpen(true)} user={user} />

        <div className="p-4 md:p-8 max-w-6xl mx-auto w-full">
          <Suspense fallback={
            <div className="flex flex-col items-center justify-center h-64">
              <div className="w-10 h-10 border-4 border-[#006E62] border-t-transparent rounded-full animate-spin"></div>
              <p className="mt-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Memuat Modul...</p>
            </div>
          }>
            {activeTab === 'dashboard' ? (
              <MobileDashboard user={user} setActiveTab={setActiveTab} />
            ) : activeTab === 'location' ? (
              <LocationMain />
            ) : activeTab === 'account' ? (
              <AccountMain />
            ) : activeTab === 'schedule' ? (
              <ScheduleMain />
            ) : activeTab === 'document' ? (
              <DocumentMain user={user} />
            ) : activeTab === 'presence' ? (
              <PresenceMain />
            ) : activeTab === 'overtime' ? (
              isAdmin ? (
                <AdminSubmissionModule 
                  user={user!} 
                  type="Lembur" 
                  title="Manajemen Lembur" 
                  subtitle="Verifikasi & Persetujuan Lembur Pegawai"
                  icon={Timer}
                />
              ) : (
                <OvertimeMain />
              )
            ) : activeTab === 'submission' ? (
              <SubmissionMain />
            ) : activeTab === 'leave' ? (
              isAdmin ? (
                <AdminSubmissionModule 
                  user={user!} 
                  type="Libur Mandiri" 
                  title="Manajemen Libur Mandiri" 
                  subtitle="Verifikasi & Persetujuan Libur Mandiri Pegawai"
                  icon={Plane}
                />
              ) : (
                <LeaveMain />
              )
            ) : activeTab === 'annual_leave' ? (
              isAdmin ? (
                <AdminSubmissionModule 
                  user={user!} 
                  type="Cuti Tahunan" 
                  title="Manajemen Cuti Tahunan" 
                  subtitle="Verifikasi & Persetujuan Cuti Tahunan Pegawai"
                  icon={Calendar}
                />
              ) : (
                <AnnualLeaveMain />
              )
            ) : activeTab === 'permission' ? (
              isAdmin ? (
                <AdminSubmissionModule 
                  user={user!} 
                  type="Izin" 
                  title="Manajemen Izin" 
                  subtitle="Verifikasi & Persetujuan Izin Pegawai"
                  icon={ClipboardList}
                />
              ) : (
                <PermissionMain />
              )
            ) : activeTab === 'maternity_leave' ? (
              isAdmin ? (
                <AdminSubmissionModule 
                  user={user!} 
                  type="Cuti Melahirkan" 
                  title="Manajemen Cuti Melahirkan" 
                  subtitle="Verifikasi & Persetujuan Cuti Melahirkan Pegawai"
                  icon={Heart}
                />
              ) : (
                <MaternityLeaveMain />
              )
            ) : activeTab === 'kpi' ? (
              <KPIMain />
            ) : activeTab === 'key_activity' ? (
              <KeyActivityMain />
            ) : activeTab === 'employee_of_the_period' ? (
              <EmployeeOfThePeriodMain />
            ) : activeTab === 'sales_report' ? (
              <SalesReportMain />
            ) : activeTab === 'feedback' ? (
              <FeedbackMain />
            ) : activeTab === 'lapor' ? (
              <LaporMain />
            ) : activeTab === 'rapat' ? (
              <RapatMain />
            ) : activeTab === 'pengumuman' ? (
              <PengumumanMain user={user} />
            ) : activeTab === 'salary_scheme' ? (
              <SalarySchemeMain />
            ) : activeTab === 'salary_adjustment' ? (
              <SalaryAdjustmentMain />
            ) : activeTab === 'payroll' ? (
              <PayrollMain />
            ) : activeTab === 'my_payslip' ? (
              <MyPayslip />
            ) : activeTab === 'reimbursement' ? (
              <ReimbursementMain />
            ) : activeTab === 'early_salary' ? (
              <EarlySalaryMain />
            ) : activeTab === 'compensation' ? (
              <CompensationMain />
            ) : activeTab === 'dispensation' ? (
              <DispensationMain user={user} />
            ) : activeTab === 'admin_dispensation' ? (
              <AdminDispensationMain user={user} />
            ) : activeTab === 'employee_report' ? (
              <EmployeeReportMain />
            ) : activeTab === 'attendance_report' ? (
              <AttendanceReportMain />
            ) : activeTab === 'finance_report' ? (
              <FinanceReportMain />
            ) : activeTab === 'master_app' ? (
              <MasterMain />
            ) : activeTab === 'admin_settings' ? (
              <AdminSettingsModule />
            ) : activeTab === 'daily_monitoring' ? (
              <DailyMonitoring />
            ) : (
              <div className="flex flex-col items-center justify-center h-64 text-gray-400 bg-gray-50 rounded-lg border border-dashed border-gray-200">
                <p className="font-medium text-sm">Modul "{activeTab}" sedang dalam pengembangan.</p>
              </div>
            )}
          </Suspense>
        </div>
      </main>
    </div>
  );
};


export default App;