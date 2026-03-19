import React from 'react';
import { Percent, ShieldCheck, Landmark, Scale, Download } from 'lucide-react';
import * as XLSX from 'xlsx';

interface FinanceTaxModuleProps {
  data: any[];
}

const FinanceTaxModule: React.FC<FinanceTaxModuleProps> = ({ data }) => {
  const totalPph21 = data.reduce((sum, item) => sum + (item.pph21 || 0), 0);
  const totalBpjsKes = data.reduce((sum, item) => sum + (item.bpjs_kesehatan || 0), 0);
  const totalBpjsKet = data.reduce((sum, item) => sum + (item.bpjs_ketenagakerjaan || 0), 0);
  const totalTaxAndInsurance = totalPph21 + totalBpjsKes + totalBpjsKet;

  const formatCurrency = (amount: number) => {
    return new Intl.NumberFormat('id-ID', {
      style: 'currency',
      currency: 'IDR',
      minimumFractionDigits: 0
    }).format(amount);
  };

  const stats = [
    { label: 'Total PPh 21', value: totalPph21, icon: Scale, color: 'text-rose-600', bg: 'bg-rose-50' },
    { label: 'BPJS Kesehatan', value: totalBpjsKes, icon: ShieldCheck, color: 'text-emerald-600', bg: 'bg-emerald-50' },
    { label: 'BPJS Ketenagakerjaan', value: totalBpjsKet, icon: Landmark, color: 'text-blue-600', bg: 'bg-blue-50' },
    { label: 'Total Kewajiban', value: totalTaxAndInsurance, icon: Percent, color: 'text-indigo-600', bg: 'bg-indigo-50' },
  ];

  const exportToExcel = () => {
    const exportData = data.map(item => ({
      'Nama Karyawan': item.account?.full_name,
      'NIK': item.account?.internal_nik,
      'PPh 21': item.pph21 || 0,
      'BPJS Kesehatan': item.bpjs_kesehatan || 0,
      'BPJS Ketenagakerjaan': item.bpjs_ketenagakerjaan || 0,
      'Total Potongan': (item.pph21 || 0) + (item.bpjs_kesehatan || 0) + (item.bpjs_ketenagakerjaan || 0)
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Laporan Pajak & BPJS');
    XLSX.writeFile(wb, `Laporan_Pajak_BPJS_${new Date().getTime()}.xlsx`);
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
            <p className="text-xl font-black text-gray-800">{formatCurrency(stat.value)}</p>
          </div>
        ))}
      </div>

      {/* Tax Table */}
      <div className="bg-white rounded-2xl border border-gray-100 shadow-sm overflow-hidden">
        <div className="p-6 border-b border-gray-50 flex items-center justify-between">
          <h3 className="text-sm font-bold text-gray-700 uppercase tracking-wider">Rincian Pajak & BPJS Karyawan</h3>
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
                <th className="py-4 px-6 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Karyawan</th>
                <th className="py-4 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">PPh 21</th>
                <th className="py-4 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">BPJS Kesehatan</th>
                <th className="py-4 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">BPJS Ketenagakerjaan</th>
                <th className="py-4 px-6 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">Total Potongan</th>
              </tr>
            </thead>
            <tbody>
              {data.length === 0 ? (
                <tr>
                  <td colSpan={5} className="py-12 text-center text-gray-400 text-xs font-bold uppercase tracking-widest">
                    Tidak ada data pajak untuk periode ini
                  </td>
                </tr>
              ) : (
                data.map((item, idx) => (
                  <tr key={idx} className="border-b border-gray-50 hover:bg-gray-50/50 transition-colors">
                    <td className="py-4 px-6">
                      <div className="flex items-center gap-3">
                        <div className="w-8 h-8 rounded-full bg-rose-100 flex items-center justify-center text-rose-600 font-bold text-xs">
                          {item.account?.full_name?.charAt(0)}
                        </div>
                        <div>
                          <p className="text-xs font-bold text-gray-700">{item.account?.full_name}</p>
                          <p className="text-[10px] text-gray-400">{item.account?.internal_nik}</p>
                        </div>
                      </div>
                    </td>
                    <td className="py-4 px-4 text-right text-xs font-medium text-rose-600">
                      {formatCurrency(item.pph21 || 0)}
                    </td>
                    <td className="py-4 px-4 text-right text-xs font-medium text-emerald-600">
                      {formatCurrency(item.bpjs_kesehatan || 0)}
                    </td>
                    <td className="py-4 px-4 text-right text-xs font-medium text-blue-600">
                      {formatCurrency(item.bpjs_ketenagakerjaan || 0)}
                    </td>
                    <td className="py-4 px-6 text-right">
                      <span className="text-xs font-black text-gray-800 bg-gray-100 px-3 py-1 rounded-lg">
                        {formatCurrency((item.pph21 || 0) + (item.bpjs_kesehatan || 0) + (item.bpjs_ketenagakerjaan || 0))}
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

export default FinanceTaxModule;
