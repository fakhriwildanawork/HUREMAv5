import React from 'react';
import { ShieldAlert, UserMinus, Gavel, CheckCircle2, Clock, Download } from 'lucide-react';
import { format } from 'date-fns';
import { id } from 'date-fns/locale';
import * as XLSX from 'xlsx';

interface FinanceCompensationModuleProps {
  data: any[];
}

const FinanceCompensationModule: React.FC<FinanceCompensationModuleProps> = ({ data }) => {
  const totalSeverance = data.filter(item => item.type === 'Severance').reduce((sum, item) => sum + (item.amount || 0), 0);
  const totalPenalty = data.filter(item => item.type === 'Penalty').reduce((sum, item) => sum + (item.amount || 0), 0);
  const completedCount = data.filter(item => item.status === 'Completed').length;

  const formatCurrency = (amount: number) => {
    return new Intl.NumberFormat('id-ID', {
      style: 'currency',
      currency: 'IDR',
      minimumFractionDigits: 0
    }).format(amount);
  };

  const stats = [
    { label: 'Total Pesangon', value: totalSeverance, icon: UserMinus, color: 'text-emerald-600', bg: 'bg-emerald-50' },
    { label: 'Total Pinalti', value: totalPenalty, icon: Gavel, color: 'text-rose-600', bg: 'bg-rose-50' },
    { label: 'Total Kompensasi', value: totalSeverance + totalPenalty, icon: ShieldAlert, color: 'text-indigo-600', bg: 'bg-indigo-50' },
    { label: 'Status Selesai', value: completedCount, icon: CheckCircle2, color: 'text-blue-600', bg: 'bg-blue-50', isCount: true },
  ];

  const exportToExcel = () => {
    const exportData = data.map(item => ({
      'Tanggal': item.termination_date,
      'Nama Karyawan': item.account?.full_name,
      'NIK': item.account?.internal_nik,
      'Jenis': item.type === 'Severance' ? 'Pesangon' : 'Pinalti',
      'Alasan': item.reason,
      'Nominal': item.amount,
      'Status': item.status
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Laporan Kompensasi');
    XLSX.writeFile(wb, `Laporan_Kompensasi_${new Date().getTime()}.xlsx`);
  };

  return (
    <div className="space-y-6">
      {/* Stats Grid */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
        {stats.map((stat, idx) => (
          <div key={idx} className="bg-white p-6 rounded-2xl border border-gray-100 shadow-sm">
            <div className="flex items-center justify-between mb-4">
              <div className={`p-2 rounded-xl ${stat.bg} ${stat.color}`}>
                <stat.icon size={20} />
              </div>
            </div>
            <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">{stat.label}</p>
            <p className="text-xl font-black text-gray-800">
              {stat.isCount ? `${stat.value} Data` : formatCurrency(stat.value)}
            </p>
          </div>
        ))}
      </div>

      {/* Compensation Table */}
      <div className="bg-white rounded-2xl border border-gray-100 shadow-sm overflow-hidden">
        <div className="p-6 border-b border-gray-50 flex items-center justify-between">
          <h3 className="text-sm font-bold text-gray-700 uppercase tracking-wider">Rincian Pesangon & Pinalti</h3>
          <button 
            onClick={exportToExcel}
            className="flex items-center gap-2 bg-emerald-600 text-white px-4 py-2 rounded-xl font-bold text-[10px] uppercase tracking-wider hover:bg-emerald-700 transition-all shadow-md shadow-emerald-600/20 active:scale-95"
          >
            <Download size={14} />
            Export Excel
          </button>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-gray-50/50">
                <th className="py-4 px-6 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tanggal</th>
                <th className="py-4 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Karyawan</th>
                <th className="py-4 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Jenis</th>
                <th className="py-4 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Alasan</th>
                <th className="py-4 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">Nominal</th>
                <th className="py-4 px-6 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Status</th>
              </tr>
            </thead>
            <tbody>
              {data.length === 0 ? (
                <tr>
                  <td colSpan={6} className="py-12 text-center text-gray-400 text-xs font-bold uppercase tracking-widest">
                    Tidak ada data kompensasi untuk periode ini
                  </td>
                </tr>
              ) : (
                data.map((item, idx) => (
                  <tr key={idx} className="border-b border-gray-50 hover:bg-gray-50/50 transition-colors">
                    <td className="py-4 px-6 text-xs font-medium text-gray-600">
                      {format(new Date(item.termination_date), 'dd MMM yyyy', { locale: id })}
                    </td>
                    <td className="py-4 px-4">
                      <p className="text-xs font-bold text-gray-700">{item.account?.full_name}</p>
                      <p className="text-[10px] text-gray-400">{item.account?.internal_nik}</p>
                    </td>
                    <td className="py-4 px-4">
                      <span className={`text-[10px] font-bold px-2 py-1 rounded uppercase tracking-tighter ${
                        item.type === 'Severance' ? 'text-emerald-600 bg-emerald-50' : 'text-rose-600 bg-rose-50'
                      }`}>
                        {item.type === 'Severance' ? 'Pesangon' : 'Pinalti'}
                      </span>
                    </td>
                    <td className="py-4 px-4 text-xs text-gray-500 max-w-xs truncate">
                      {item.reason}
                    </td>
                    <td className="py-4 px-4 text-right text-xs font-black text-gray-800">
                      {formatCurrency(item.amount)}
                    </td>
                    <td className="py-4 px-6 text-center">
                      <span className={`text-[10px] font-black px-3 py-1 rounded-full uppercase tracking-widest ${
                        item.status === 'Completed' ? 'text-emerald-600 bg-emerald-50' : 'text-amber-600 bg-amber-50'
                      }`}>
                        {item.status === 'Completed' ? 'Selesai' : 'Pending'}
                      </span>
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

export default FinanceCompensationModule;
