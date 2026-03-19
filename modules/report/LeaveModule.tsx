import React, { useState, useEffect } from 'react';
import { reportService } from '../../services/reportService';
import { LeaveSummary } from '../../types';
import { Search, Download, Plane, Heart, ClipboardList, User, PieChart as PieIcon, BarChart3 } from 'lucide-react';
import StatCard from './ReportStatCardComponent';
import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip as RechartsTooltip, BarChart, Bar, XAxis, YAxis, CartesianGrid } from 'recharts';
import * as XLSX from 'xlsx';

const LeaveModule: React.FC = () => {
  const [data, setData] = useState<LeaveSummary[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [searchTerm, setSearchTerm] = useState('');

  useEffect(() => {
    fetchData();
  }, []);

  const fetchData = async () => {
    try {
      setIsLoading(true);
      const report = await reportService.getLeaveReport();
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
    used: acc.used + curr.usedQuota,
    maternity: acc.maternity + curr.maternityUsed,
    permission: acc.permission + curr.permissionCount,
  }), { used: 0, maternity: 0, permission: 0 });

  const chartData = [
    { name: 'Cuti Tahunan', value: totals.used, color: '#3b82f6' },
    { name: 'Cuti Melahirkan', value: totals.maternity, color: '#ec4899' },
    { name: 'Izin', value: totals.permission, color: '#6366f1' },
  ];

  const exportToExcel = () => {
    const exportData = filteredData.map(item => ({
      'Nama Karyawan': item.fullName,
      'NIK': item.nik,
      'Total Kuota Cuti': item.totalQuota,
      'Cuti Terpakai': item.usedQuota,
      'Sisa Cuti': item.remainingQuota,
      'Cuti Melahirkan (Hari)': item.maternityUsed,
      'Total Izin (Kali)': item.permissionCount
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Laporan Cuti & Izin');
    XLSX.writeFile(wb, `Laporan_Cuti_Izin_${new Date().toLocaleDateString()}.xlsx`);
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      {/* Filters & Actions */}
      <div className="bg-white p-6 rounded-2xl border border-gray-100 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
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
        <div className="lg:col-span-2 grid grid-cols-1 sm:grid-cols-3 gap-4">
          <StatCard label="Total Cuti Tahunan" value={`${totals.used} Hari`} icon={Plane} color="bg-blue-500" />
          <StatCard label="Total Cuti Melahirkan" value={`${totals.maternity} Hari`} icon={Heart} color="bg-pink-500" />
          <StatCard label="Total Izin" value={`${totals.permission} Kali`} icon={ClipboardList} color="bg-indigo-500" />
        </div>
        <div className="bg-white p-6 rounded-2xl border border-gray-100 shadow-sm flex flex-col items-center justify-center min-h-[200px]">
          <h3 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-4 w-full">Distribusi Ketidakhadiran</h3>
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
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Kuota Total</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Cuti Terpakai</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Sisa Cuti</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Cuti Melahirkan</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">Total Izin</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {isLoading ? (
                <tr>
                  <td colSpan={6} className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center gap-2">
                      <div className="w-8 h-8 border-4 border-gray-200 border-t-[#006E62] rounded-full animate-spin"></div>
                      <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Mengambil Data...</span>
                    </div>
                  </td>
                </tr>
              ) : filteredData.length === 0 ? (
                <tr>
                  <td colSpan={6} className="px-6 py-12 text-center text-gray-400 italic text-xs">Tidak ada data cuti/izin.</td>
                </tr>
              ) : (
                filteredData.map((item) => (
                  <tr key={item.accountId} className="hover:bg-gray-50/50 transition-colors">
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
                      <span className="text-xs font-bold text-gray-600">{item.totalQuota} Hari</span>
                    </td>
                    <td className="px-6 py-4 text-center">
                      <span className="text-xs font-black text-blue-600">{item.usedQuota} Hari</span>
                    </td>
                    <td className="px-6 py-4 text-center">
                      <span className={`text-xs font-black ${item.remainingQuota <= 2 ? 'text-rose-500' : 'text-emerald-600'}`}>{item.remainingQuota} Hari</span>
                    </td>
                    <td className="px-6 py-4 text-center">
                      <span className="text-xs font-bold text-pink-500">{item.maternityUsed} Hari</span>
                    </td>
                    <td className="px-6 py-4 text-right">
                      <span className="text-xs font-bold text-indigo-600">{item.permissionCount} Kali</span>
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

export default LeaveModule;
