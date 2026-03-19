
import React, { useState, useEffect, useMemo } from 'react';
import { 
  Calendar, 
  Users, 
  Clock, 
  FileText, 
  Filter, 
  Download, 
  ChevronRight, 
  BarChart3, 
  PieChart as PieChartIcon,
  TrendingUp,
  Timer,
  Plane,
  Heart,
  ClipboardList,
  AlertCircle,
  XCircle,
  Check
} from 'lucide-react';
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  Legend, 
  ResponsiveContainer, 
  LineChart, 
  Line, 
  PieChart, 
  Pie, 
  Cell,
  AreaChart,
  Area
} from 'recharts';
import { reportService } from '../../services/reportService';
import { authService } from '../../services/authService';
import LoadingSpinner from '../../components/Common/LoadingSpinner';
import { format, addDays, isWithinInterval, parseISO, eachDayOfInterval, startOfMonth, endOfMonth, subDays, differenceInDays, isAfter } from 'date-fns';
import { id } from 'date-fns/locale';

const COLORS = ['#006E62', '#10b981', '#f59e0b', '#ef4444', '#6366f1', '#ec4899', '#8b5cf6'];

const AttendanceReportMain: React.FC = () => {
  const [dateRange, setDateRange] = useState({
    start: format(startOfMonth(new Date()), 'yyyy-MM-dd'),
    end: format(new Date(), 'yyyy-MM-dd')
  });
  const [tempDateRange, setTempDateRange] = useState({
    start: format(startOfMonth(new Date()), 'yyyy-MM-dd'),
    end: format(new Date(), 'yyyy-MM-dd')
  });
  const [isLoading, setIsLoading] = useState(true);
  const [reportData, setReportData] = useState<any>(null);
  const [selectedEmployee, setSelectedEmployee] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');

  const user = authService.getCurrentUser();
  const isAdmin = user?.role === 'admin' || user?.is_hr_admin || user?.is_performance_admin || user?.is_finance_admin;

  useEffect(() => {
    fetchReport();
  }, [dateRange]);

  const fetchReport = async () => {
    try {
      setIsLoading(true);
      const data = await reportService.getAttendanceReport(dateRange.start, dateRange.end);
      setReportData(data);
    } catch (error) {
      console.error('Error fetching report:', error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleApplyFilter = () => {
    const start = parseISO(tempDateRange.start);
    const end = parseISO(tempDateRange.end);
    const today = new Date();
    today.setHours(23, 59, 59, 999);

    if (isAfter(start, today) || isAfter(end, today)) {
      alert('Rentang tanggal tidak boleh melebihi tanggal sekarang');
      return;
    }

    if (isAfter(start, end)) {
      alert('Tanggal mulai tidak boleh melebihi tanggal selesai');
      return;
    }

    const diff = differenceInDays(end, start);
    if (diff > 93) {
      alert('Rentang tanggal maksimal adalah 93 hari');
      return;
    }

    setDateRange(tempDateRange);
  };

  const processedData = useMemo(() => {
    if (!reportData) return [];

    const { accounts, attendances, overtimes, leaves, annualLeaves, permissions, maternityLeaves } = reportData;
    const totalDays = eachDayOfInterval({
      start: parseISO(dateRange.start),
      end: parseISO(dateRange.end)
    }).length;

    return accounts.map((acc: any) => {
      const employeeAttendances = attendances.filter((a: any) => a.account_id === acc.id);
      const employeeOvertimes = overtimes.filter((o: any) => o.account_id === acc.id);
      const employeeLeaves = leaves.filter((l: any) => l.account_id === acc.id);
      const employeeAnnualLeaves = annualLeaves.filter((l: any) => l.account_id === acc.id);
      const employeePermissions = permissions.filter((p: any) => p.account_id === acc.id);
      const employeeMaternityLeaves = maternityLeaves.filter((m: any) => m.account_id === acc.id);

      const present = employeeAttendances.length;
      const leave = employeeLeaves.length;
      const annual_leave = employeeAnnualLeaves.length;
      const permission = employeePermissions.length;
      const maternity_leave = employeeMaternityLeaves.length;
      
      // Basic absent logic: total days - (present + leaves/permissions)
      // Note: This is a simplification and doesn't account for weekends/holidays yet
      const absent = Math.max(0, totalDays - (present + leave + annual_leave + permission + maternity_leave));

      return {
        id: acc.id,
        name: acc.full_name,
        nik: acc.internal_nik,
        position: acc.position,
        grade: acc.grade,
        location: acc.location?.name || '-',
        present,
        overtime: employeeOvertimes.length,
        overtime_minutes: employeeOvertimes.reduce((sum: number, o: any) => sum + (o.duration_minutes || 0), 0),
        leave,
        annual_leave,
        permission,
        maternity_leave,
        absent
      };
    });
  }, [reportData, dateRange]);

  const filteredData = useMemo(() => {
    if (!searchTerm) return processedData;
    const lowerSearch = searchTerm.toLowerCase();
    return processedData.filter((emp: any) => 
      emp.name.toLowerCase().includes(lowerSearch) || 
      emp.nik.toLowerCase().includes(lowerSearch) ||
      (emp.position && emp.position.toLowerCase().includes(lowerSearch)) ||
      (emp.grade && emp.grade.toLowerCase().includes(lowerSearch)) ||
      (emp.location && emp.location.toLowerCase().includes(lowerSearch))
    );
  }, [processedData, searchTerm]);

  const overallStats = useMemo(() => {
    if (!processedData.length) return null;

    const present = processedData.reduce((sum, d) => sum + d.present, 0);
    const overtime = processedData.reduce((sum, d) => sum + d.overtime, 0);
    const leave = processedData.reduce((sum, d) => sum + d.leave + d.annual_leave + d.maternity_leave, 0);
    const permission = processedData.reduce((sum, d) => sum + d.permission, 0);
    const absent = processedData.reduce((sum, d) => sum + d.absent, 0);

    return [
      { name: 'Hadir', value: present, color: '#006E62' },
      { name: 'Lembur', value: overtime, color: '#f59e0b' },
      { name: 'Cuti', value: leave, color: '#10b981' },
      { name: 'Izin', value: permission, color: '#6366f1' },
      { name: 'Absen', value: absent, color: '#ef4444' }
    ];
  }, [processedData]);

  const attendancePercentage = useMemo(() => {
    if (!processedData.length) return 0;
    const totalPresent = processedData.reduce((sum, d) => sum + d.present, 0);
    const totalPossible = processedData.length * eachDayOfInterval({
      start: parseISO(dateRange.start),
      end: parseISO(dateRange.end)
    }).length;
    
    return totalPossible > 0 ? ((totalPresent / totalPossible) * 100).toFixed(1) : 0;
  }, [processedData, dateRange]);

  const dailyTrend = useMemo(() => {
    if (!reportData) return [];
    
    const days = eachDayOfInterval({
      start: parseISO(dateRange.start),
      end: parseISO(dateRange.end)
    });

    return days.map(day => {
      const dateStr = format(day, 'yyyy-MM-dd');
      const count = reportData.attendances.filter((a: any) => a.created_at.startsWith(dateStr)).length;
      return {
        date: format(day, 'dd MMM', { locale: id }),
        count
      };
    });
  }, [reportData, dateRange]);

  if (isLoading) return <LoadingSpinner />;

  return (
    <div className="space-y-8 pb-12">
      {/* Header Section */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-2xl font-bold text-gray-800 tracking-tight">Laporan Kehadiran</h2>
          <p className="text-sm text-gray-500">Analisis data presensi, lembur, dan ketidakhadiran karyawan</p>
        </div>
        <div className="flex items-center gap-3 bg-white p-2 rounded-lg border border-gray-100 shadow-sm">
          <div className="flex items-center gap-2 px-3 py-1.5 border-r border-gray-100">
            <Calendar size={16} className="text-[#006E62]" />
            <input 
              type="date" 
              value={tempDateRange.start}
              onChange={(e) => setTempDateRange(prev => ({ ...prev, start: e.target.value }))}
              max={format(new Date(), 'yyyy-MM-dd')}
              className="text-xs font-bold outline-none bg-transparent"
            />
            <span className="text-gray-300 mx-1">-</span>
            <input 
              type="date" 
              value={tempDateRange.end}
              onChange={(e) => setTempDateRange(prev => ({ ...prev, end: e.target.value }))}
              max={format(new Date(), 'yyyy-MM-dd')}
              className="text-xs font-bold outline-none bg-transparent"
            />
            <button 
              onClick={handleApplyFilter}
              className="ml-2 p-1 bg-[#006E62] text-white rounded hover:bg-[#005a50] transition-colors"
              title="Terapkan Filter"
            >
              <Check size={14} />
            </button>
          </div>
          <button className="p-2 text-gray-400 hover:text-[#006E62] transition-colors">
            <Download size={18} />
          </button>
        </div>
      </div>

      {/* Summary Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        <div className="bg-white p-6 rounded-xl border border-gray-100 shadow-sm hover:shadow-md transition-shadow group">
          <div className="flex items-center justify-between mb-4">
            <div className="p-2 bg-[#006E62]/10 rounded-lg text-[#006E62] group-hover:scale-110 transition-transform">
              <Users size={20} />
            </div>
            <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Total Karyawan</span>
          </div>
          <p className="text-3xl font-bold text-gray-800">{processedData.length}</p>
          <p className="text-xs text-gray-400 mt-1">Terdaftar dalam sistem</p>
        </div>

        <div className="bg-white p-6 rounded-xl border border-gray-100 shadow-sm hover:shadow-md transition-shadow group">
          <div className="flex items-center justify-between mb-4">
            <div className="p-2 bg-amber-50 rounded-lg text-amber-600 group-hover:scale-110 transition-transform">
              <Timer size={20} />
            </div>
            <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Total Lembur</span>
          </div>
          <p className="text-3xl font-bold text-gray-800">
            {Math.floor(processedData.reduce((sum, d) => sum + d.overtime_minutes, 0) / 60)}h {processedData.reduce((sum, d) => sum + d.overtime_minutes, 0) % 60}m
          </p>
          <p className="text-xs text-gray-400 mt-1">Akumulasi jam lembur</p>
        </div>

        <div className="bg-white p-6 rounded-xl border border-gray-100 shadow-sm hover:shadow-md transition-shadow group">
          <div className="flex items-center justify-between mb-4">
            <div className="p-2 bg-emerald-50 rounded-lg text-emerald-600 group-hover:scale-110 transition-transform">
              <TrendingUp size={20} />
            </div>
            <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">% Kehadiran</span>
          </div>
          <p className="text-3xl font-bold text-gray-800">{attendancePercentage}%</p>
          <p className="text-xs text-emerald-500 mt-1 font-medium">Berdasarkan data presensi</p>
        </div>

        <div className="bg-white p-6 rounded-xl border border-gray-100 shadow-sm hover:shadow-md transition-shadow group">
          <div className="flex items-center justify-between mb-4">
            <div className="p-2 bg-rose-50 rounded-lg text-rose-600 group-hover:scale-110 transition-transform">
              <AlertCircle size={20} />
            </div>
            <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Ketidakhadiran</span>
          </div>
          <p className="text-3xl font-bold text-gray-800">{processedData.reduce((sum, d) => sum + d.absent, 0)}</p>
          <p className="text-xs text-rose-500 mt-1 font-medium">Total absen periode ini</p>
        </div>
      </div>

      {/* Charts Section */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Daily Trend Chart */}
        <div className="lg:col-span-2 bg-white p-6 rounded-xl border border-gray-100 shadow-sm">
          <div className="flex items-center justify-between mb-6">
            <h3 className="text-sm font-bold text-gray-700 uppercase tracking-wider flex items-center gap-2">
              <BarChart3 size={16} className="text-[#006E62]" />
              Tren Kehadiran Harian
            </h3>
          </div>
          <div className="h-[300px] w-full">
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={dailyTrend}>
                <defs>
                  <linearGradient id="colorCount" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#006E62" stopOpacity={0.1}/>
                    <stop offset="95%" stopColor="#006E62" stopOpacity={0}/>
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f0f0f0" />
                <XAxis 
                  dataKey="date" 
                  axisLine={false} 
                  tickLine={false} 
                  tick={{ fontSize: 10, fill: '#9ca3af' }}
                />
                <YAxis 
                  axisLine={false} 
                  tickLine={false} 
                  tick={{ fontSize: 10, fill: '#9ca3af' }}
                />
                <Tooltip 
                  contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 12px rgba(0,0,0,0.1)' }}
                  labelStyle={{ fontWeight: 'bold', color: '#374151', marginBottom: '4px' }}
                />
                <Area 
                  type="monotone" 
                  dataKey="count" 
                  stroke="#006E62" 
                  strokeWidth={3}
                  fillOpacity={1} 
                  fill="url(#colorCount)" 
                  name="Jumlah Hadir"
                />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Overall Composition Chart */}
        <div className="bg-white p-6 rounded-xl border border-gray-100 shadow-sm">
          <div className="flex items-center justify-between mb-6">
            <h3 className="text-sm font-bold text-gray-700 uppercase tracking-wider flex items-center gap-2">
              <PieChartIcon size={16} className="text-[#006E62]" />
              Komposisi Kehadiran
            </h3>
          </div>
          <div className="h-[300px] w-full relative">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie
                  data={overallStats || []}
                  cx="50%"
                  cy="50%"
                  innerRadius={60}
                  outerRadius={80}
                  paddingAngle={5}
                  dataKey="value"
                >
                  {overallStats?.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={entry.color} />
                  ))}
                </Pie>
                <Tooltip />
              </PieChart>
            </ResponsiveContainer>
            <div className="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
              <span className="text-2xl font-bold text-gray-800">
                {overallStats?.reduce((sum, d) => sum + d.value, 0)}
              </span>
              <span className="text-[10px] text-gray-400 uppercase font-bold tracking-tighter">Total Data</span>
            </div>
          </div>
          <div className="mt-4 space-y-2">
            {overallStats?.map((stat, index) => (
              <div key={index} className="flex items-center justify-between text-xs">
                <div className="flex items-center gap-2">
                  <div className="w-2 h-2 rounded-full" style={{ backgroundColor: stat.color }}></div>
                  <span className="text-gray-500">{stat.name}</span>
                </div>
                <span className="font-bold text-gray-700">{stat.value}</span>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* Overtime & Leave Summary Grid */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Overtime Report */}
        <div className="bg-white p-6 rounded-xl border border-gray-100 shadow-sm">
          <h3 className="text-sm font-bold text-gray-700 uppercase tracking-wider mb-6 flex items-center gap-2">
            <Timer size={16} className="text-amber-500" />
            Laporan Lembur
          </h3>
          <div className="space-y-4">
            {processedData.sort((a, b) => b.overtime_minutes - a.overtime_minutes).slice(0, 5).map((emp, idx) => (
              <div key={idx} className="flex items-center gap-4">
                <div className="flex-1">
                  <div className="flex justify-between mb-1">
                    <span className="text-xs font-bold text-gray-700">{emp.name}</span>
                    <span className="text-xs text-gray-400">{Math.floor(emp.overtime_minutes / 60)}h {emp.overtime_minutes % 60}m</span>
                  </div>
                  <div className="w-full bg-gray-100 rounded-full h-1.5">
                    <div 
                      className="bg-amber-500 h-1.5 rounded-full" 
                      style={{ width: `${Math.min((emp.overtime_minutes / 600) * 100, 100)}%` }}
                    ></div>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Leave & Permission Summary */}
        <div className="bg-white p-6 rounded-xl border border-gray-100 shadow-sm">
          <h3 className="text-sm font-bold text-gray-700 uppercase tracking-wider mb-6 flex items-center gap-2">
            <Plane size={16} className="text-emerald-500" />
            Laporan Cuti & Izin
          </h3>
          <div className="grid grid-cols-2 gap-4">
            <div className="p-4 bg-emerald-50 rounded-xl border border-emerald-100">
              <div className="flex items-center gap-2 text-emerald-600 mb-2">
                <Calendar size={16} />
                <span className="text-[10px] font-bold uppercase tracking-wider">Cuti Tahunan</span>
              </div>
              <p className="text-2xl font-bold text-emerald-700">{processedData.reduce((sum, d) => sum + d.annual_leave, 0)}</p>
              <p className="text-[10px] text-emerald-600/60 mt-1 uppercase font-bold">Total Hari</p>
            </div>
            <div className="p-4 bg-rose-50 rounded-xl border border-rose-100">
              <div className="flex items-center gap-2 text-rose-600 mb-2">
                <Heart size={16} />
                <span className="text-[10px] font-bold uppercase tracking-wider">Melahirkan</span>
              </div>
              <p className="text-2xl font-bold text-rose-700">{processedData.reduce((sum, d) => sum + d.maternity_leave, 0)}</p>
              <p className="text-[10px] text-rose-600/60 mt-1 uppercase font-bold">Total Hari</p>
            </div>
            <div className="p-4 bg-blue-50 rounded-xl border border-blue-100">
              <div className="flex items-center gap-2 text-blue-600 mb-2">
                <ClipboardList size={16} />
                <span className="text-[10px] font-bold uppercase tracking-wider">Izin</span>
              </div>
              <p className="text-2xl font-bold text-blue-700">{processedData.reduce((sum, d) => sum + d.permission, 0)}</p>
              <p className="text-[10px] text-blue-600/60 mt-1 uppercase font-bold">Total Hari</p>
            </div>
            <div className="p-4 bg-indigo-50 rounded-xl border border-indigo-100">
              <div className="flex items-center gap-2 text-indigo-600 mb-2">
                <Plane size={16} />
                <span className="text-[10px] font-bold uppercase tracking-wider">Libur Mandiri</span>
              </div>
              <p className="text-2xl font-bold text-indigo-700">{processedData.reduce((sum, d) => sum + d.leave, 0)}</p>
              <p className="text-[10px] text-indigo-600/60 mt-1 uppercase font-bold">Total Hari</p>
            </div>
          </div>
        </div>
      </div>

      {/* Attendance Table */}
      <div className="bg-white rounded-xl border border-gray-100 shadow-sm overflow-hidden">
        <div className="p-6 border-b border-gray-50 flex flex-col md:flex-row md:items-center justify-between gap-4">
          <h3 className="text-sm font-bold text-gray-700 uppercase tracking-wider flex items-center gap-2">
            <FileText size={16} className="text-[#006E62]" />
            Data Presensi Reguler
          </h3>
          <div className="flex items-center gap-2">
            <div className="relative">
              <Filter className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={14} />
              <input 
                type="text" 
                placeholder="Cari karyawan..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="pl-9 pr-4 py-1.5 text-xs border border-gray-200 rounded-md outline-none focus:ring-2 focus:ring-[#006E62]/10"
              />
            </div>
          </div>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-gray-50/50">
                <th className="py-3 px-6 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Karyawan</th>
                <th className="py-3 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Hadir</th>
                <th className="py-3 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Lembur</th>
                <th className="py-3 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Cuti</th>
                <th className="py-3 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Cuti Melahirkan</th>
                <th className="py-3 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Izin</th>
                <th className="py-3 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Libur Mandiri</th>
                <th className="py-3 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Absen</th>
                <th className="py-3 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">Aksi</th>
              </tr>
            </thead>
            <tbody>
              {filteredData.map((emp: any) => (
                <tr key={emp.id} className="border-b border-gray-50 hover:bg-gray-50/50 transition-colors group">
                  <td className="py-4 px-6">
                    <div className="flex items-center gap-3">
                      <div className="w-8 h-8 rounded-full bg-[#006E62]/10 flex items-center justify-center text-[#006E62] font-bold text-xs">
                        {emp.name.charAt(0)}
                      </div>
                      <div>
                        <p className="text-xs font-bold text-gray-700">{emp.name}</p>
                        <p className="text-[10px] text-gray-400">{emp.nik}</p>
                      </div>
                    </div>
                  </td>
                  <td className="py-4 px-4 text-center">
                    <span className="text-xs font-bold text-emerald-600 bg-emerald-50 px-2 py-1 rounded">{emp.present}</span>
                  </td>
                  <td className="py-4 px-4 text-center">
                    <span className="text-xs font-bold text-amber-600 bg-amber-50 px-2 py-1 rounded">{emp.overtime}</span>
                  </td>
                  <td className="py-4 px-4 text-center text-xs font-medium text-gray-600">{emp.annual_leave}</td>
                  <td className="py-4 px-4 text-center text-xs font-medium text-gray-600">{emp.maternity_leave}</td>
                  <td className="py-4 px-4 text-center text-xs font-medium text-gray-600">{emp.permission}</td>
                  <td className="py-4 px-4 text-center text-xs font-medium text-gray-600">{emp.leave}</td>
                  <td className="py-4 px-4 text-center">
                    <span className="text-xs font-bold text-rose-600 bg-rose-50 px-2 py-1 rounded">{emp.absent}</span>
                  </td>
                  <td className="py-4 px-4 text-right">
                    <button 
                      onClick={() => setSelectedEmployee(emp.id)}
                      className="p-1.5 text-gray-400 hover:text-[#006E62] hover:bg-[#006E62]/5 rounded transition-all"
                    >
                      <ChevronRight size={18} />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Employee Detail Modal (Heatmap Placeholder) */}
      {selectedEmployee && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4 backdrop-blur-sm">
          <div className="bg-white rounded-2xl w-full max-w-4xl max-h-[90vh] flex flex-col overflow-hidden shadow-2xl animate-in fade-in zoom-in duration-200">
            <div className="p-6 border-b border-gray-100 flex items-center justify-between flex-shrink-0">
              <div className="flex items-center gap-4">
                <div className="w-12 h-12 rounded-full bg-[#006E62] flex items-center justify-center text-white font-bold text-lg">
                  {processedData.find(e => e.id === selectedEmployee)?.name.charAt(0)}
                </div>
                <div>
                  <h3 className="text-lg font-bold text-gray-800">{processedData.find(e => e.id === selectedEmployee)?.name}</h3>
                  <p className="text-xs text-gray-400">{processedData.find(e => e.id === selectedEmployee)?.nik}</p>
                </div>
              </div>
              <button 
                onClick={() => setSelectedEmployee(null)}
                className="p-2 hover:bg-gray-100 rounded-full transition-colors"
              >
                <XCircle size={24} className="text-gray-400" />
              </button>
            </div>
            <div className="p-8 overflow-y-auto flex-1">
              <div className="mb-8">
                <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-4">Pola Kehadiran (Heatmap)</h4>
                <div className="grid grid-cols-7 gap-2">
                  {eachDayOfInterval({
                    start: parseISO(dateRange.start),
                    end: parseISO(dateRange.end)
                  }).map((day, i) => {
                    const dateStr = format(day, 'yyyy-MM-dd');
                    const isPresent = reportData?.attendances.some((a: any) => 
                      a.account_id === selectedEmployee && a.check_in.startsWith(dateStr)
                    );
                    const isOnLeave = [
                      ...reportData?.leaves,
                      ...reportData?.annualLeaves,
                      ...reportData?.permissions,
                      ...reportData?.maternityLeaves
                    ].some((l: any) => 
                      l.account_id === selectedEmployee && 
                      isWithinInterval(day, { start: parseISO(l.start_date), end: parseISO(l.end_date) })
                    );

                    let bgColor = 'bg-gray-100 text-gray-400';
                    if (isPresent) bgColor = 'bg-[#006E62] text-white';
                    else if (isOnLeave) bgColor = 'bg-amber-100 text-amber-600';

                    return (
                      <div 
                        key={i} 
                        className={`aspect-square rounded-md flex items-center justify-center text-[10px] font-bold ${bgColor}`}
                        title={format(day, 'dd MMMM yyyy', { locale: id })}
                      >
                        {format(day, 'd')}
                      </div>
                    );
                  })}
                </div>
                <div className="mt-4 flex items-center gap-4 text-[10px] font-bold text-gray-400 uppercase tracking-tighter">
                  <div className="flex items-center gap-1">
                    <div className="w-3 h-3 bg-[#006E62] rounded-sm"></div>
                    <span>Hadir</span>
                  </div>
                  <div className="flex items-center gap-1">
                    <div className="w-3 h-3 bg-amber-100 rounded-sm"></div>
                    <span>Cuti/Izin</span>
                  </div>
                  <div className="flex items-center gap-1">
                    <div className="w-3 h-3 bg-gray-100 rounded-sm"></div>
                    <span>Absen</span>
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div className="p-4 bg-gray-50 rounded-xl border border-gray-100">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Total Hadir</p>
                  <p className="text-2xl font-bold text-gray-800">{processedData.find(e => e.id === selectedEmployee)?.present} Hari</p>
                </div>
                <div className="p-4 bg-gray-50 rounded-xl border border-gray-100">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Total Lembur</p>
                  <p className="text-2xl font-bold text-gray-800">{processedData.find(e => e.id === selectedEmployee)?.overtime} Kali</p>
                </div>
                <div className="p-4 bg-gray-50 rounded-xl border border-gray-100">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Total Cuti/Izin</p>
                  <p className="text-2xl font-bold text-gray-800">
                    {(processedData.find(e => e.id === selectedEmployee)?.leave || 0) + 
                     (processedData.find(e => e.id === selectedEmployee)?.annual_leave || 0) + 
                     (processedData.find(e => e.id === selectedEmployee)?.permission || 0)} Hari
                  </p>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default AttendanceReportMain;
