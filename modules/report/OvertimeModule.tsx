import React, { useState, useEffect } from 'react';
import { reportService } from '../../services/reportService';
import { OvertimeSummary } from '../../types';
import { Search, Calendar, Download, Timer, User, Clock, DollarSign, BarChart3 } from 'lucide-react';
import StatCard from './ReportStatCardComponent';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, ResponsiveContainer, Cell } from 'recharts';
import { format, subDays } from 'date-fns';
import * as XLSX from 'xlsx';

interface OvertimeModuleProps {
  initialStartDate?: string;
  initialEndDate?: string;
  initialData?: OvertimeSummary[];
}

const OvertimeModule: React.FC<OvertimeModuleProps> = ({ 
  initialStartDate, 
  initialEndDate,
  initialData 
}) => {
  const [startDate, setStartDate] = useState(initialStartDate || format(subDays(new Date(), 30), 'yyyy-MM-dd'));
  const [endDate, setEndDate] = useState(initialEndDate || format(new Date(), 'yyyy-MM-dd'));
  const [data, setData] = useState<OvertimeSummary[]>(initialData || []);
  const [isLoading, setIsLoading] = useState(!initialData);
  const [searchTerm, setSearchTerm] = useState('');

  useEffect(() => {
    if (initialStartDate) setStartDate(initialStartDate);
    if (initialEndDate) setEndDate(initialEndDate);
  }, [initialStartDate, initialEndDate]);

  useEffect(() => {
    if (initialData) {
      setData(initialData);
      setIsLoading(false);
    } else {
      fetchData();
    }
  }, [startDate, endDate, initialData]);

  const fetchData = async () => {
    try {
      setIsLoading(true);
      const report = await reportService.getOvertimeReport(startDate, endDate);
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
    hours: acc.hours + curr.totalOvertimeHours,
    count: acc.count + curr.overtimeCount,
    cost: acc.cost + curr.estimatedCost,
  }), { hours: 0, count: 0, cost: 0 });

  const chartData = [...filteredData]
    .sort((a, b) => b.totalOvertimeHours - a.totalOvertimeHours)
    .slice(0, 10);

  const exportToExcel = () => {
    const exportData = filteredData.map(item => ({
      'Nama Karyawan': item.fullName,
      'NIK': item.nik,
      'Total Menit Lembur': item.totalOvertimeMinutes,
      'Total Jam Lembur': item.totalOvertimeHours,
      'Frekuensi Lembur': item.overtimeCount,
      'Estimasi Biaya Lembur': item.estimatedCost
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Laporan Lembur');
    XLSX.writeFile(wb, `Laporan_Lembur_${startDate}_to_${endDate}.xlsx`);
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      {/* Filters & Actions */}
      <div className="bg-white p-6 rounded-2xl border border-gray-100 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex flex-col sm:flex-row gap-3">
          {!initialStartDate && (
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
          )}
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
        <div className="lg:col-span-2 grid grid-cols-1 sm:grid-cols-3 gap-4">
          <StatCard label="Total Jam Lembur" value={`${totals.hours.toFixed(1)} Jam`} icon={Timer} color="bg-indigo-500" />
          <StatCard label="Frekuensi Lembur" value={`${totals.count} Kali`} icon={Clock} color="bg-blue-500" />
          <StatCard label="Estimasi Biaya" value={`Rp ${totals.cost.toLocaleString()}`} icon={DollarSign} color="bg-emerald-500" />
        </div>
        <div className="bg-white p-6 rounded-2xl border border-gray-100 shadow-sm flex flex-col items-center justify-center min-h-[200px]">
          <h3 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-4 w-full">Top 10 Lembur (Jam)</h3>
          <div className="w-full h-40">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={chartData} layout="vertical">
                <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f3f4f6" />
                <XAxis type="number" hide />
                <YAxis dataKey="fullName" type="category" hide />
                <RechartsTooltip 
                  cursor={{ fill: 'transparent' }}
                  content={({ active, payload }) => {
                    if (active && payload && payload.length) {
                      return (
                        <div className="bg-gray-800 text-white text-[10px] p-2 rounded shadow-xl">
                          <p className="font-bold">{payload[0].payload.fullName}</p>
                          <p>{payload[0].value} Jam</p>
                        </div>
                      );
                    }
                    return null;
                  }}
                />
                <Bar dataKey="totalOvertimeHours" radius={[0, 4, 4, 0]}>
                  {chartData.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={index === 0 ? '#4f46e5' : '#818cf8'} />
                  ))}
                </Bar>
              </BarChart>
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
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Menit Lembur</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Jam Lembur</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Frekuensi</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">Estimasi Biaya</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {isLoading ? (
                <tr>
                  <td colSpan={5} className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center gap-2">
                      <div className="w-8 h-8 border-4 border-gray-200 border-t-[#006E62] rounded-full animate-spin"></div>
                      <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Memproses Data...</span>
                    </div>
                  </td>
                </tr>
              ) : filteredData.length === 0 ? (
                <tr>
                  <td colSpan={5} className="px-6 py-12 text-center text-gray-400 italic text-xs">Tidak ada data lembur untuk periode ini.</td>
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
                      <span className="text-xs font-bold text-gray-600">{item.totalOvertimeMinutes.toLocaleString()}</span>
                    </td>
                    <td className="px-6 py-4 text-center">
                      <span className="text-xs font-black text-indigo-600">{item.totalOvertimeHours} Jam</span>
                    </td>
                    <td className="px-6 py-4 text-center">
                      <span className="text-xs font-bold text-gray-600">{item.overtimeCount}x</span>
                    </td>
                    <td className="px-6 py-4 text-right">
                      <span className="text-xs font-black text-emerald-600">Rp {item.estimatedCost.toLocaleString()}</span>
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

export default OvertimeModule;
