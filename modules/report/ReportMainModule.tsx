import React, { useState, useEffect } from 'react';
import { Fingerprint, Timer, Plane, BarChart3, PieChart, TrendingUp, FileText, Wallet } from 'lucide-react';
import AttendanceModule from './AttendanceModule';
import OvertimeModule from './OvertimeModule';
import LeaveModule from './LeaveModule';
import FinanceReportMain from './FinanceReportMain';

interface ReportMainModuleProps {
  initialTab?: 'attendance' | 'overtime' | 'leave' | 'finance';
}

const ReportMainModule: React.FC<ReportMainModuleProps> = ({ initialTab = 'attendance' }) => {
  const [activeSubTab, setActiveSubTab] = useState<'attendance' | 'overtime' | 'leave' | 'finance'>(initialTab);

  useEffect(() => {
    setActiveSubTab(initialTab);
  }, [initialTab]);

  const tabs = [
    { id: 'attendance', label: 'Kehadiran Reguler', icon: Fingerprint, color: 'text-emerald-600', bg: 'bg-emerald-50' },
    { id: 'overtime', label: 'Lembur & Ekstra', icon: Timer, color: 'text-indigo-600', bg: 'bg-indigo-50' },
    { id: 'leave', label: 'Cuti & Izin', icon: Plane, color: 'text-blue-600', bg: 'bg-blue-50' },
    { id: 'finance', label: 'Finance', icon: Wallet, color: 'text-rose-600', bg: 'bg-rose-50' },
  ];

  return (
    <div className="min-h-screen bg-gray-50/30 p-4 md:p-8">
      {/* Header Section */}
      <div className="mb-8">
        <div className="flex items-center gap-3 mb-2">
          <div className="w-10 h-10 bg-[#006E62] rounded-xl flex items-center justify-center text-white shadow-lg shadow-[#006E62]/20">
            <BarChart3 size={24} />
          </div>
          <h1 className="text-2xl font-black text-gray-800 tracking-tight">Analitik & Laporan SDM</h1>
        </div>
        <p className="text-sm text-gray-500 font-medium">Pantau performa, kedisiplinan, dan tren kehadiran seluruh karyawan dalam satu dashboard terpadu.</p>
      </div>

      {/* Tab Navigation */}
      <div className="flex flex-wrap gap-2 mb-8">
        {tabs.map((tab) => (
          <button
            key={tab.id}
            onClick={() => setActiveSubTab(tab.id as any)}
            className={`flex items-center gap-3 px-6 py-3 rounded-2xl font-bold text-xs uppercase tracking-wider transition-all duration-300 ${
              activeSubTab === tab.id 
                ? `${tab.bg} ${tab.color} shadow-sm border border-${tab.color.split('-')[1]}-100` 
                : 'bg-white text-gray-400 hover:bg-gray-50 border border-gray-100'
            }`}
          >
            <tab.icon size={18} />
            {tab.label}
          </button>
        ))}
      </div>

      {/* Content Area */}
      <div className="animate-in fade-in slide-in-from-bottom-4 duration-500">
        {activeSubTab === 'attendance' && <AttendanceModule />}
        {activeSubTab === 'overtime' && <OvertimeModule />}
        {activeSubTab === 'leave' && <LeaveModule />}
        {activeSubTab === 'finance' && <FinanceReportMain />}
      </div>
    </div>
  );
};

export default ReportMainModule;
