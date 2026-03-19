import React from 'react';
import { Wallet, TrendingUp, Users, ArrowUpRight, ArrowDownRight, Download } from 'lucide-react';
import * as XLSX from 'xlsx';

interface FinancePayrollModuleProps {
  data: any[];
}

const FinancePayrollModule: React.FC<FinancePayrollModuleProps> = ({ data }) => {
  const totalBasicSalary = data.reduce((sum, item) => sum + (item.basic_salary || 0), 0);
  const totalAllowances = data.reduce((sum, item) => 
    sum + (item.position_allowance || 0) + (item.placement_allowance || 0) + (item.other_allowance || 0), 0);
  const totalOvertime = data.reduce((sum, item) => sum + (item.overtime_pay || 0), 0);
  const totalReimbursement = data.reduce((sum, item) => sum + (item.reimbursement_pay || 0), 0);
  const totalDeductions = data.reduce((sum, item) => sum + (item.total_deduction || 0), 0);
  const totalNetPay = data.reduce((sum, item) => sum + (item.take_home_pay || 0) + (item.reimbursement_pay || 0), 0);

  const stats = [
    { label: 'Total Dibayarkan', value: totalNetPay, icon: Wallet, color: 'text-emerald-600', bg: 'bg-emerald-50' },
    { label: 'Total Gaji Pokok', value: totalBasicSalary, icon: TrendingUp, color: 'text-blue-600', bg: 'bg-blue-50' },
    { label: 'Total Reimbursement', value: totalReimbursement, icon: ArrowUpRight, color: 'text-indigo-600', bg: 'bg-indigo-50' },
    { label: 'Total Potongan', value: totalDeductions, icon: ArrowDownRight, color: 'text-rose-600', bg: 'bg-rose-50' },
  ];

  const formatCurrency = (amount: number) => {
    return new Intl.NumberFormat('id-ID', {
      style: 'currency',
      currency: 'IDR',
      minimumFractionDigits: 0
    }).format(amount);
  };

  const exportToExcel = () => {
    const exportData = data.map(item => ({
      'Nama Karyawan': item.account?.full_name,
      'NIK': item.account?.internal_nik,
      'Gaji Pokok': item.basic_salary,
      'Tunjangan Jabatan': item.position_allowance || 0,
      'Tunjangan Penempatan': item.placement_allowance || 0,
      'Tunjangan Lainnya': item.other_allowance || 0,
      'Lembur': item.overtime_pay || 0,
      'Reimbursement': item.reimbursement_pay || 0,
      'Potongan': item.total_deduction || 0,
      'Total Dibayarkan': (item.take_home_pay || 0) + (item.reimbursement_pay || 0)
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Laporan Payroll');
    XLSX.writeFile(wb, `Laporan_Payroll_${new Date().getTime()}.xlsx`);
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

      {/* Payroll Table */}
      <div className="bg-white rounded-2xl border border-gray-100 shadow-sm overflow-hidden">
        <div className="p-6 border-b border-gray-50 flex items-center justify-between">
          <h3 className="text-sm font-bold text-gray-700 uppercase tracking-wider">Rincian Penggajian Karyawan</h3>
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
                <th className="py-4 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">Gaji Pokok</th>
                <th className="py-4 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">Tunjangan</th>
                <th className="py-4 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">Lembur</th>
                <th className="py-4 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">Reimbursement</th>
                <th className="py-4 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">Potongan</th>
                <th className="py-4 px-6 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">Total Dibayarkan</th>
              </tr>
            </thead>
            <tbody>
              {data.length === 0 ? (
                <tr>
                  <td colSpan={6} className="py-12 text-center text-gray-400 text-xs font-bold uppercase tracking-widest">
                    Tidak ada data payroll untuk periode ini
                  </td>
                </tr>
              ) : (
                data.map((item, idx) => (
                  <tr key={idx} className="border-b border-gray-50 hover:bg-gray-50/50 transition-colors">
                    <td className="py-4 px-6">
                      <div className="flex items-center gap-3">
                        <div className="w-8 h-8 rounded-full bg-emerald-100 flex items-center justify-center text-emerald-600 font-bold text-xs">
                          {item.account?.full_name?.charAt(0)}
                        </div>
                        <div>
                          <p className="text-xs font-bold text-gray-700">{item.account?.full_name}</p>
                          <p className="text-[10px] text-gray-400">{item.account?.internal_nik}</p>
                        </div>
                      </div>
                    </td>
                    <td className="py-4 px-4 text-right text-xs font-medium text-gray-600">
                      {formatCurrency(item.basic_salary)}
                    </td>
                    <td className="py-4 px-4 text-right text-xs font-medium text-emerald-600">
                      {formatCurrency((item.position_allowance || 0) + (item.placement_allowance || 0) + (item.other_allowance || 0))}
                    </td>
                    <td className="py-4 px-4 text-right text-xs font-medium text-blue-600">
                      {formatCurrency(item.overtime_pay || 0)}
                    </td>
                    <td className="py-4 px-4 text-right text-xs font-medium text-indigo-600">
                      {formatCurrency(item.reimbursement_pay || 0)}
                    </td>
                    <td className="py-4 px-4 text-right text-xs font-medium text-rose-600">
                      {formatCurrency(item.total_deduction || 0)}
                    </td>
                    <td className="py-4 px-6 text-right">
                      <span className="text-xs font-black text-gray-800 bg-gray-100 px-3 py-1 rounded-lg">
                        {formatCurrency((item.take_home_pay || 0) + (item.reimbursement_pay || 0))}
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

export default FinancePayrollModule;
