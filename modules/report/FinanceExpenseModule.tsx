import React from 'react';
import { Receipt, CheckCircle2, Clock, XCircle, FileText, Download } from 'lucide-react';
import { format } from 'date-fns';
import { id } from 'date-fns/locale';
import * as XLSX from 'xlsx';

interface FinanceExpenseModuleProps {
  data: any[];
}


const FinanceExpenseModule: React.FC<FinanceExpenseModuleProps> = ({ data }) => {
  const totalRequested = data.reduce((sum, item) => sum + (item.amount_requested || 0), 0);
  const totalApproved = data.reduce((sum, item) => sum + (item.amount_approved || 0), 0);
  const approvedCount = data.filter(item => item.status === 'Approved').length;
  const partiallyApprovedCount = data.filter(item => item.status === 'Partially Approved').length;

  const formatCurrency = (amount: number) => {
    return new Intl.NumberFormat('id-ID', {
      style: 'currency',
      currency: 'IDR',
      minimumFractionDigits: 0
    }).format(amount);
  };

  const getStatusColor = (status: string) => {
    switch (status) {
      case 'Approved': return 'text-emerald-600 bg-emerald-50';
      case 'Partially Approved': return 'text-blue-600 bg-blue-50';
      case 'Rejected': return 'text-rose-600 bg-rose-50';
      default: return 'text-amber-600 bg-amber-50';
    }
  };

  const exportToExcel = () => {
    const exportData = data.map(item => ({
      'Tanggal': item.transaction_date,
      'Nama Karyawan': item.account?.full_name,
      'NIK': item.account?.internal_nik,
      'Kategori': item.category,
      'Diajukan': item.amount_requested,
      'Disetujui': item.amount_approved || 0,
      'Status': item.status,
      'Deskripsi': item.description
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Laporan Reimbursement');
    XLSX.writeFile(wb, `Laporan_Reimbursement_${new Date().getTime()}.xlsx`);
  };

  return (
    <div className="space-y-6">
      {/* Summary Cards */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
        <div className="bg-white p-6 rounded-2xl border border-gray-100 shadow-sm">
          <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Total Diajukan</p>
          <p className="text-xl font-black text-gray-800">{formatCurrency(totalRequested)}</p>
          <div className="mt-2 flex items-center gap-1 text-[10px] font-bold text-blue-600">
            <Clock size={12} />
            <span>{data.length} Transaksi</span>
          </div>
        </div>
        <div className="bg-white p-6 rounded-2xl border border-gray-100 shadow-sm">
          <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Total Disetujui</p>
          <p className="text-xl font-black text-emerald-600">{formatCurrency(totalApproved)}</p>
          <div className="mt-2 flex items-center gap-1 text-[10px] font-bold text-emerald-600">
            <CheckCircle2 size={12} />
            <span>{approvedCount} Full Approved</span>
          </div>
        </div>
        <div className="bg-white p-6 rounded-2xl border border-gray-100 shadow-sm">
          <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Partial Approved</p>
          <p className="text-xl font-black text-blue-600">{partiallyApprovedCount}</p>
          <p className="text-[10px] text-gray-400 mt-1 uppercase font-bold tracking-tighter">Transaksi Disetujui Sebagian</p>
        </div>
        <div className="bg-white p-6 rounded-2xl border border-gray-100 shadow-sm">
          <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Efisiensi Klaim</p>
          <p className="text-xl font-black text-indigo-600">
            {totalRequested > 0 ? Math.round((totalApproved / totalRequested) * 100) : 0}%
          </p>
          <p className="text-[10px] text-gray-400 mt-1 uppercase font-bold tracking-tighter">Rasio Disetujui vs Diajukan</p>
        </div>
      </div>

      {/* Expense Table */}
      <div className="bg-white rounded-2xl border border-gray-100 shadow-sm overflow-hidden">
        <div className="p-6 border-b border-gray-50 flex items-center justify-between">
          <h3 className="text-sm font-bold text-gray-700 uppercase tracking-wider">Daftar Klaim & Reimbursement</h3>
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
                <th className="py-4 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Kategori</th>
                <th className="py-4 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">Diajukan</th>
                <th className="py-4 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">Disetujui</th>
                <th className="py-4 px-6 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Status</th>
              </tr>
            </thead>
            <tbody>
              {data.length === 0 ? (
                <tr>
                  <td colSpan={6} className="py-12 text-center text-gray-400 text-xs font-bold uppercase tracking-widest">
                    Tidak ada data reimbursement untuk periode ini
                  </td>
                </tr>
              ) : (
                data.map((item, idx) => (
                  <tr key={idx} className="border-b border-gray-50 hover:bg-gray-50/50 transition-colors">
                    <td className="py-4 px-6 text-xs font-medium text-gray-600">
                      {format(new Date(item.transaction_date), 'dd MMM yyyy', { locale: id })}
                    </td>
                    <td className="py-4 px-4">
                      <p className="text-xs font-bold text-gray-700">{item.account?.full_name}</p>
                      <p className="text-[10px] text-gray-400">{item.account?.internal_nik}</p>
                    </td>
                    <td className="py-4 px-4">
                      <span className="text-[10px] font-bold text-indigo-600 bg-indigo-50 px-2 py-1 rounded uppercase tracking-tighter">
                        {item.category}
                      </span>
                    </td>
                    <td className="py-4 px-4 text-right text-xs font-medium text-gray-600">
                      {formatCurrency(item.amount_requested)}
                    </td>
                    <td className="py-4 px-4 text-right text-xs font-bold text-emerald-600">
                      {item.amount_approved ? formatCurrency(item.amount_approved) : '-'}
                    </td>
                    <td className="py-4 px-6 text-center">
                      <span className={`text-[10px] font-black px-3 py-1 rounded-full uppercase tracking-widest ${getStatusColor(item.status)}`}>
                        {item.status}
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

export default FinanceExpenseModule;
