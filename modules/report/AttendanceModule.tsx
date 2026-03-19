import React, { useState, useEffect } from 'react';
import { reportService } from '../../services/reportService';
import { AttendanceSummary } from '../../types';
import { Search, Calendar, Download, Filter, User, ChevronRight, Fingerprint, Clock, AlertTriangle, Plane, Heart, ClipboardList, Coffee, CheckSquare } from 'lucide-react';
import StatCard from './ReportStatCardComponent';
import AttendanceHeatmap from './ReportHeatmapComponent';
import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip as RechartsTooltip, Legend } from 'recharts';
import { format, subDays, parseISO } from 'date-fns';
import { id } from 'date-fns/locale';
import * as XLSX from 'xlsx';

const AttendanceModule: React.FC = () => {
  const [startDate, setStartDate] = useState(format(subDays(new Date(), 30), 'yyyy-MM-dd'));
  const [endDate, setEndDate] = useState(format(new Date(), 'yyyy-MM-dd'));
  const [data, setData] = useState<AttendanceSummary[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedEmployee, setSelectedEmployee] = useState<AttendanceSummary | null>(null);

  useEffect(() => {
    fetchData();
  }, [startDate, endDate]);

  const fetchData = async () => {
    try {
      setIsLoading(true);
      const report = await reportService.getAttendanceReportSummary(startDate, endDate);
      setData(report);
    } catch (error) {
      console.error(error);
    } finally {
      setIsLoading(false);
    }
  };

  const filteredData = data.filter(item => 
    item.fullName.toLowerCase().includes(searchTerm.toLowerCase()) ||
    item.nik.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const totals = data.reduce((acc, curr) => ({
    present: acc.present + curr.present,
    late: acc.late + curr.late,
    absent: acc.absent + curr.absent,
    leave: acc.leave + curr.leave + curr.maternityLeave + curr.permission,
  }), { present: 0, late: 0, absent: 0, leave: 0 });

  const chartData = [
    { name: 'Hadir Tepat Waktu', value: totals.present - totals.late, color: '#10b981' },
    { name: 'Terlambat', value: totals.late, color: '#fbbf24' },
    { name: 'Mangkir', value: totals.absent, color: '#f43f5e' },
    { name: 'Cuti/Izin', value: totals.leave, color: '#60a5fa' },
  ];

  const exportToExcel = () => {
    const exportData = filteredData.map(item => ({
      'Nama Karyawan': item.fullName,
      'NIK': item.nik,
      'Total Hari': item.totalDays,
      'Hadir': item.present,
      'Terlambat': item.late,
      'Menit Terlambat': item.lateMinutes,
      'Pulang Cepat': item.earlyDeparture,
      'Mangkir': item.absent,
      'Cuti Tahunan': item.leave,
      'Cuti Melahirkan': item.maternityLeave,
      'Izin': item.permission,
      'Libur Reguler': item.holiday,
      'Libur Khusus': item.specialHoliday,
      'Tanpa Absen Pulang': item.noClockOut,
      'Rate Kehadiran (%)': item.attendanceRate.toFixed(2)
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Laporan Kehadiran');
    XLSX.writeFile(wb, `Laporan_Kehadiran_${startDate}_to_${endDate}.xlsx`);
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      {/* Filters & Actions */}
      <div className="bg-white p-6 rounded-2xl border border-gray-100 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex flex-col sm:flex-row gap-3">
          <div className="flex items-center gap-2 bg-gray-50 px-3 py-2 rounded-xl border border-gray-100">
            <Calendar size={16} className="text-gray-400" />
            <input 
              type="date" 
              value={startDate}
              onChange={(e) => setStartDate(e.target.value)}
              className="bg-transparent text-xs font-bold text-gray-600 outline-none"
            />
            <span className="text-gray-300">/</span>
            <input 
              type="date" 
              value={endDate}
              onChange={(e) => setEndDate(e.target.value)}
              className="bg-transparent text-xs font-bold text-gray-600 outline-none"
            />
          </div>
          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
            <input 
              type="text"
              placeholder="Cari nama atau NIK..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="pl-10 pr-4 py-2.5 bg-gray-50 border border-gray-100 rounded-xl text-xs focus:ring-2 focus:ring-[#006E62] outline-none w-full sm:w-64 transition-all"
            />
          </div>
        </div>
        <button 
          onClick={exportToExcel}
          className="flex items-center justify-center gap-2 bg-emerald-600 text-white px-6 py-3 rounded-xl font-bold text-xs uppercase tracking-wider hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-600/20 active:scale-95"
        >
          <Download size={18} />
          Export Excel
        </button>
      </div>

      {/* Summary Stats */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <div className="lg:col-span-2 grid grid-cols-2 sm:grid-cols-4 gap-4">
          <StatCard label="Hadir" value={totals.present} icon={Fingerprint} color="bg-emerald-500" />
          <StatCard label="Terlambat" value={totals.late} icon={Clock} color="bg-amber-400" />
          <StatCard label="Mangkir" value={totals.absent} icon={AlertTriangle} color="bg-rose-500" />
          <StatCard label="Cuti/Izin" value={totals.leave} icon={Plane} color="bg-blue-400" />
        </div>
        <div className="bg-white p-6 rounded-2xl border border-gray-100 shadow-sm flex flex-col items-center justify-center min-h-[200px]">
          <h3 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-4 w-full">Komposisi Kehadiran</h3>
          <div className="w-full h-40">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie
                  data={chartData}
                  cx="50%"
                  cy="50%"
                  innerRadius={40}
                  outerRadius={60}
                  paddingAngle={5}
                  dataKey="value"
                >
                  {chartData.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={entry.color} />
                  ))}
                </Pie>
                <RechartsTooltip />
              </PieChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      {/* Main Table */}
      <div className="bg-white rounded-2xl border border-gray-100 shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-gray-50/50 border-b border-gray-100">
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Karyawan</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Hadir</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Terlambat</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Mangkir</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Cuti/Izin</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Rate (%)</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">Detail</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {isLoading ? (
                <tr>
                  <td colSpan={7} className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center gap-2">
                      <div className="w-8 h-8 border-4 border-gray-200 border-t-[#006E62] rounded-full animate-spin"></div>
                      <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Menghitung Data...</span>
                    </div>
                  </td>
                </tr>
              ) : filteredData.length === 0 ? (
                <tr>
                  <td colSpan={7} className="px-6 py-12 text-center text-gray-400 italic text-xs">Tidak ada data untuk periode ini.</td>
                </tr>
              ) : (
                filteredData.map((item) => (
                  <tr 
                    key={item.accountId} 
                    className={`hover:bg-gray-50/50 transition-colors cursor-pointer group ${selectedEmployee?.accountId === item.accountId ? 'bg-emerald-50/30' : ''}`}
                    onClick={() => setSelectedEmployee(item)}
                  >
                    <td className="px-6 py-4">
                      <div className="flex items-center gap-3">
                        <div className="w-8 h-8 bg-gray-100 rounded-lg flex items-center justify-center text-gray-400">
                          <User size={16} />
                        </div>
                        <div>
                          <p className="text-xs font-bold text-gray-800">{item.fullName}</p>
                          <p className="text-[10px] text-gray-400 font-medium">{item.nik}</p>
                        </div>
                      </div>
                    </td>
                    <td className="px-6 py-4 text-center">
                      <span className="text-xs font-black text-emerald-600">{item.present}</span>
                    </td>
                    <td className="px-6 py-4 text-center">
                      <span className={`text-xs font-black ${item.late > 0 ? 'text-amber-500' : 'text-gray-300'}`}>{item.late}</span>
                    </td>
                    <td className="px-6 py-4 text-center">
                      <span className={`text-xs font-black ${item.absent > 0 ? 'text-rose-500' : 'text-gray-300'}`}>{item.absent}</span>
                    </td>
                    <td className="px-6 py-4 text-center">
                      <span className="text-xs font-black text-blue-500">{item.leave + item.maternityLeave + item.permission}</span>
                    </td>
                    <td className="px-6 py-4 text-center">
                      <div className="flex flex-col items-center gap-1">
                        <span className={`text-xs font-black ${item.attendanceRate >= 90 ? 'text-emerald-600' : item.attendanceRate >= 75 ? 'text-amber-500' : 'text-rose-500'}`}>
                          {item.attendanceRate.toFixed(1)}%
                        </span>
                        <div className="w-12 h-1 bg-gray-100 rounded-full overflow-hidden">
                          <div className={`h-full ${item.attendanceRate >= 90 ? 'bg-emerald-500' : item.attendanceRate >= 75 ? 'bg-amber-500' : 'bg-rose-500'}`} style={{ width: `${item.attendanceRate}%` }}></div>
                        </div>
                      </div>
                    </td>
                    <td className="px-6 py-4 text-right">
                      <ChevronRight size={16} className={`inline-block transition-transform ${selectedEmployee?.accountId === item.accountId ? 'rotate-90 text-[#006E62]' : 'text-gray-300 group-hover:translate-x-1'}`} />
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Employee Detail Section */}
      {selectedEmployee && (
        <div className="bg-white p-6 rounded-2xl border border-[#006E62]/20 shadow-xl shadow-[#006E62]/5 animate-in slide-in-from-bottom-4 duration-500">
          <div className="flex items-center justify-between mb-8 pb-4 border-b border-gray-100">
            <div className="flex items-center gap-4">
              <div className="w-12 h-12 bg-[#006E62] rounded-xl flex items-center justify-center text-white shadow-lg shadow-[#006E62]/20">
                <User size={24} />
              </div>
              <div>
                <h3 className="text-lg font-black text-gray-800">{selectedEmployee.fullName}</h3>
                <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Detail Kehadiran Periode {format(parseISO(startDate), 'dd MMM', { locale: id })} - {format(parseISO(endDate), 'dd MMM yyyy', { locale: id })}</p>
              </div>
            </div>
            <button 
              onClick={() => setSelectedEmployee(null)}
              className="text-gray-400 hover:text-rose-500 transition-colors"
            >
              Tutup Detail
            </button>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
            {/* Heatmap */}
            <AttendanceHeatmap details={selectedEmployee.dailyDetails} startDate={startDate} endDate={endDate} />

            {/* Detailed Stats */}
            <div className="space-y-6">
              <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Rincian Akumulasi</h4>
              <div className="grid grid-cols-2 sm:grid-cols-3 gap-4">
                {[
                  { label: 'Hadir', value: selectedEmployee.present, icon: Fingerprint, color: 'text-emerald-600' },
                  { label: 'Terlambat', value: selectedEmployee.late, icon: Clock, color: 'text-amber-500' },
                  { label: 'Mangkir', value: selectedEmployee.absent, icon: AlertTriangle, color: 'text-rose-500' },
                  { label: 'Cuti', value: selectedEmployee.leave, icon: Plane, color: 'text-blue-500' },
                  { label: 'Cuti Melahirkan', value: selectedEmployee.maternityLeave, icon: Heart, color: 'text-pink-500' },
                  { label: 'Izin', value: selectedEmployee.permission, icon: ClipboardList, color: 'text-indigo-500' },
                  { label: 'Libur', value: selectedEmployee.holiday + selectedEmployee.specialHoliday, icon: Coffee, color: 'text-gray-400' },
                  { label: 'Tanpa Out', value: selectedEmployee.noClockOut, icon: AlertTriangle, color: 'text-orange-500' },
                  { label: 'Dispensasi', value: selectedEmployee.dispensationCount, icon: CheckSquare, color: 'text-[#006E62]' },
                ].map((stat, i) => (
                  <div key={i} className="bg-gray-50/50 p-3 rounded-xl border border-gray-100 flex flex-col items-center text-center">
                    <stat.icon size={14} className={`${stat.color} mb-1`} />
                    <span className="text-[8px] font-bold text-gray-400 uppercase tracking-wider mb-1">{stat.label}</span>
                    <span className={`text-sm font-black ${stat.color}`}>{stat.value}</span>
                  </div>
                ))}
              </div>
              
              <div className="bg-[#006E62]/5 p-4 rounded-xl border border-[#006E62]/10">
                <div className="flex items-center justify-between mb-2">
                  <span className="text-[10px] font-bold text-[#006E62] uppercase tracking-widest">Total Keterlambatan</span>
                  <span className="text-sm font-black text-[#006E62]">{selectedEmployee.lateMinutes} Menit</span>
                </div>
                <div className="flex items-center justify-between">
                  <span className="text-[10px] font-bold text-[#006E62] uppercase tracking-widest">Total Pulang Cepat</span>
                  <span className="text-sm font-black text-[#006E62]">{selectedEmployee.earlyDepartureMinutes} Menit</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default AttendanceModule;
