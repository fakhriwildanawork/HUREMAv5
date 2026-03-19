import React, { useState, useEffect } from 'react';
import { Wallet, Receipt, Percent, ShieldAlert, Calendar, Download, Filter, Search, ChevronRight, Timer } from 'lucide-react';
import { format, startOfMonth, endOfMonth } from 'date-fns';
import { id } from 'date-fns/locale';
import { reportService } from '../../services/reportService';
import FinancePayrollModule from './FinancePayrollModule';
import FinanceExpenseModule from './FinanceExpenseModule';
import FinanceTaxModule from './FinanceTaxModule';
import FinanceCompensationModule from './FinanceCompensationModule';
import OvertimeModule from './OvertimeModule';

const FinanceReportMain: React.FC = () => {
  const [activeTab, setActiveTab] = useState<'payroll' | 'overtime' | 'expense' | 'tax' | 'compensation'>('payroll');
  const [dateRange, setDateRange] = useState({
    start: format(startOfMonth(new Date()), 'yyyy-MM-dd'),
    end: format(endOfMonth(new Date()), 'yyyy-MM-dd')
  });
  const [reportData, setReportData] = useState<any>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    fetchData();
  }, [dateRange]);

  const fetchData = async () => {
    setLoading(true);
    try {
      const data = await reportService.getFinanceReport(dateRange.start, dateRange.end);
      setReportData(data);
    } catch (error) {
      console.error('Error fetching finance report:', error);
    } finally {
      setLoading(false);
    }
  };

  const tabs = [
    { id: 'payroll', label: 'Rekapitulasi Gaji', icon: Wallet, color: 'text-emerald-600', bg: 'bg-emerald-50' },
    { id: 'overtime', label: 'Lembur', icon: Timer, color: 'text-indigo-600', bg: 'bg-indigo-50' },
    { id: 'expense', label: 'Reimbursement & Klaim', icon: Receipt, color: 'text-blue-600', bg: 'bg-blue-50' },
    { id: 'tax', label: 'Pajak PPh 21 & BPJS', icon: Percent, color: 'text-rose-600', bg: 'bg-rose-50' },
    { id: 'compensation', label: 'Kompensasi', icon: ShieldAlert, color: 'text-amber-600', bg: 'bg-amber-50' },
  ];

  return (
    <div className="space-y-6">
      {/* Header Section */}
      <div className="mb-8">
        <div className="flex items-center gap-3 mb-2">
          <div className="w-10 h-10 bg-[#006E62] rounded-xl flex items-center justify-center text-white shadow-lg shadow-[#006E62]/20">
            <Wallet size={24} />
          </div>
          <h1 className="text-2xl font-black text-gray-800 tracking-tight">Laporan Finance</h1>
        </div>
        <p className="text-sm text-gray-500 font-medium">Analisis data penggajian, lembur, reimbursement, pajak, dan kompensasi karyawan.</p>
      </div>

      {/* Filter Section */}
      <div className="bg-white p-6 rounded-2xl border border-gray-100 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-center gap-4">
          <div className="relative">
            <Calendar className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
            <input 
              type="date" 
              value={dateRange.start}
              onChange={(e) => setDateRange(prev => ({ ...prev, start: e.target.value }))}
              className="pl-10 pr-4 py-2 text-sm border border-gray-200 rounded-xl outline-none focus:ring-2 focus:ring-[#006E62]/10"
            />
          </div>
          <span className="text-gray-400 font-bold">s/d</span>
          <div className="relative">
            <Calendar className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
            <input 
              type="date" 
              value={dateRange.end}
              onChange={(e) => setDateRange(prev => ({ ...prev, end: e.target.value }))}
              className="pl-10 pr-4 py-2 text-sm border border-gray-200 rounded-xl outline-none focus:ring-2 focus:ring-[#006E62]/10"
            />
          </div>
        </div>
        <button 
          onClick={() => {/* Export Logic */}}
          className="flex items-center gap-2 px-6 py-2 bg-[#006E62] text-white rounded-xl font-bold text-xs uppercase tracking-wider hover:bg-[#005a50] transition-all shadow-lg shadow-[#006E62]/20"
        >
          <Download size={16} />
          Export Laporan
        </button>
      </div>

      {/* Sub-Tab Navigation */}
      <div className="flex flex-wrap gap-2">
        {tabs.map((tab) => (
          <button
            key={tab.id}
            onClick={() => setActiveTab(tab.id as any)}
            className={`flex items-center gap-2 px-4 py-2.5 rounded-xl font-bold text-[10px] uppercase tracking-widest transition-all ${
              activeTab === tab.id 
                ? `${tab.bg} ${tab.color} border border-current` 
                : 'bg-white text-gray-400 hover:bg-gray-50 border border-gray-100'
            }`}
          >
            <tab.icon size={14} />
            {tab.label}
          </button>
        ))}
      </div>

      {/* Module Content */}
      <div className="animate-in fade-in slide-in-from-bottom-4 duration-500">
        {loading ? (
          <div className="flex flex-col items-center justify-center py-20 bg-white rounded-2xl border border-gray-100">
            <div className="w-10 h-10 border-4 border-[#006E62]/20 border-t-[#006E62] rounded-full animate-spin mb-4"></div>
            <p className="text-sm text-gray-400 font-bold uppercase tracking-widest">Memuat Data Keuangan...</p>
          </div>
        ) : (
          <>
            {activeTab === 'payroll' && <FinancePayrollModule data={reportData?.payrollItems || []} />}
            {activeTab === 'overtime' && (
              <OvertimeModule 
                initialStartDate={dateRange.start} 
                initialEndDate={dateRange.end} 
                initialData={reportData?.overtimes || []} 
              />
            )}
            {activeTab === 'expense' && <FinanceExpenseModule data={reportData?.reimbursements || []} />}
            {activeTab === 'tax' && <FinanceTaxModule data={reportData?.payrollItems || []} />}
            {activeTab === 'compensation' && <FinanceCompensationModule data={reportData?.compensations || []} />}
          </>
        )}
      </div>
    </div>
  );
};

export default FinanceReportMain;
