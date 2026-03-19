import React, { useState, useEffect } from 'react';
import { Search, Eye, Download, DollarSign, Calendar, FileText, AlertCircle, X, MapPin, Trash2 } from 'lucide-react';
import { financeService } from '../../services/financeService';
import { authService } from '../../services/authService';
import { Payroll, PayrollItem, PayrollSettings } from '../../types';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';

const MyPayslip: React.FC = () => {
  const user = authService.getCurrentUser();
  const [items, setItems] = useState<PayrollItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [viewingItem, setViewingItem] = useState<PayrollItem | null>(null);
  const [settings, setSettings] = useState<PayrollSettings | null>(null);
  const payslipRef = React.useRef<HTMLDivElement>(null);

  useEffect(() => {
    const fetchData = async () => {
      if (!user?.id) return;
      setLoading(true);
      try {
        const [iData, sData] = await Promise.all([
          financeService.getEmployeePayslips(user.id),
          financeService.getPayrollSettings()
        ]);
        setItems(iData);
        setSettings(sData);
      } catch (error) {
        console.error('Error fetching payslips:', error);
      } finally {
        setLoading(false);
      }
    };
    fetchData();
  }, [user?.id]);

  const downloadPDF = async (item: PayrollItem) => {
    if (!payslipRef.current) return;
    const canvas = await html2canvas(payslipRef.current, { scale: 2, useCORS: true });
    const imgData = canvas.toDataURL('image/png');
    const pdf = new jsPDF('p', 'mm', 'a4');
    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfHeight = (canvas.height * pdfWidth) / canvas.width;
    pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, pdfHeight);
    pdf.save(`Payslip_${new Date(0, (item.payroll?.month || 1) - 1).toLocaleString('id-ID', { month: 'long' })}${item.payroll?.year}.pdf`);
  };

  if (loading) {
    return (
      <div className="flex flex-col items-center justify-center min-h-[400px] gap-4">
        <div className="w-12 h-12 border-4 border-[#006E62] border-t-transparent rounded-full animate-spin"></div>
        <p className="text-sm font-bold text-gray-400 uppercase tracking-widest">Memuat Slip Gaji...</p>
      </div>
    );
  }

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-2xl font-bold text-gray-800 tracking-tight">Slip Gaji Saya</h2>
          <p className="text-sm text-gray-500">Riwayat pendapatan dan slip gaji bulanan Anda.</p>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        {items.length === 0 ? (
          <div className="col-span-full bg-white p-12 rounded-2xl border border-dashed border-gray-200 text-center">
            <DollarSign size={48} className="mx-auto text-gray-300 mb-4" strokeWidth={1} />
            <p className="text-gray-500 font-medium">Belum ada slip gaji yang diterbitkan untuk Anda.</p>
          </div>
        ) : (
          items.map((item) => (
            <div key={item.id} className="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 hover:shadow-md transition-all group">
              <div className="flex items-start justify-between mb-4">
                <div className="p-3 bg-emerald-50 text-[#006E62] rounded-xl">
                  <Calendar size={24} />
                </div>
                <div className="text-right">
                  <div className="text-xs font-bold text-gray-400 uppercase tracking-widest">Periode</div>
                  <div className="text-lg font-black text-gray-800">
                    {new Date(0, (item.payroll?.month || 1) - 1).toLocaleString('id-ID', { month: 'long' })} {item.payroll?.year}
                  </div>
                </div>
              </div>

              <div className="space-y-3 mb-6">
                <div className="flex justify-between text-sm">
                  <span className="text-gray-500">Take Home Pay</span>
                  <span className="font-black text-[#006E62]">Rp {item.take_home_pay.toLocaleString('id-ID')}</span>
                </div>
                <div className="w-full bg-gray-100 h-1.5 rounded-full overflow-hidden">
                  <div className="bg-[#006E62] h-full" style={{ width: '100%' }}></div>
                </div>
              </div>

              <div className="flex gap-2">
                <button
                  onClick={() => setViewingItem(item)}
                  className="flex-1 flex items-center justify-center gap-2 py-2.5 bg-[#006E62] text-white rounded-lg hover:bg-[#005a50] transition-all text-[10px] font-bold uppercase tracking-wider"
                >
                  <Eye size={14} />
                  Lihat Detail
                </button>
                <button
                  onClick={() => downloadPDF(item)}
                  className="p-2.5 border border-gray-200 text-gray-500 rounded-lg hover:bg-gray-50 transition-all"
                >
                  <Download size={18} />
                </button>
              </div>
            </div>
          ))
        )}
      </div>

      {/* Modal View Detail (Reuse the same design as PayslipDetail) */}
      {viewingItem && (
        <div className="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-4xl max-h-[90vh] flex flex-col">
            <div className="p-4 border-b border-gray-100 flex items-center justify-between bg-gray-50 rounded-t-2xl">
              <div className="flex items-center gap-2">
                <FileText className="text-[#006E62]" size={20} />
                <h3 className="font-bold text-gray-800 uppercase tracking-widest text-xs">Slip Gaji Anda</h3>
              </div>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => downloadPDF(viewingItem)}
                  className="flex items-center gap-2 px-4 py-1.5 bg-[#006E62] text-white rounded-lg hover:bg-[#005a50] transition-all text-[10px] font-bold uppercase tracking-wider"
                >
                  <Download size={14} />
                  Download PDF
                </button>
                <button
                  onClick={() => setViewingItem(null)}
                  className="p-1.5 hover:bg-gray-200 rounded-full transition-colors text-gray-500"
                >
                  <X size={20} />
                </button>
              </div>
            </div>
            <div className="flex-1 overflow-y-auto p-8 bg-gray-100">
               {/* Same Payslip Content as in PayslipDetail.tsx */}
               <div ref={payslipRef} className="bg-white w-full max-w-[210mm] mx-auto p-12 shadow-lg min-h-[297mm] flex flex-col">
                  {/* ... Header, Employee Info, Components, Summary, Footer ... */}
                  {/* (I'll copy the content from PayslipDetail for consistency) */}
                  <div className="flex justify-between items-start border-b-4 border-[#006E62] pb-6 mb-8">
                    <div className="flex items-center gap-6">
                      {settings?.company_logo_url && <img src={settings.company_logo_url} alt="Logo" className="w-24 h-24 object-contain" referrerPolicy="no-referrer" />}
                      <div className="space-y-1">
                        <h1 className="text-3xl font-black text-[#006E62] tracking-tighter uppercase">{settings?.company_name || 'HUREMA HRIS'}</h1>
                        <div className="text-xs text-gray-500 max-w-md leading-relaxed">
                          {settings?.company_address && <div className="flex items-center gap-2"><MapPin size={10} /> {settings.company_address}</div>}
                          <div className="flex items-center gap-4 mt-1">
                            {settings?.company_phone && <span>Telp: {settings.company_phone}</span>}
                            {settings?.company_email && <span>Email: {settings.company_email}</span>}
                          </div>
                        </div>
                      </div>
                    </div>
                    <div className="text-right">
                      <h2 className="text-2xl font-black text-gray-300 uppercase tracking-widest">Slip Gaji</h2>
                      <div className="text-sm font-bold text-gray-800 mt-2">
                        {new Date(0, (viewingItem.payroll?.month || 1) - 1).toLocaleString('id-ID', { month: 'long' })} {viewingItem.payroll?.year}
                      </div>
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-12 mb-12 bg-gray-50 p-6 rounded-2xl border border-gray-100">
                    <div className="space-y-4">
                      <div className="space-y-1">
                        <div className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Nama Pegawai</div>
                        <div className="text-lg font-black text-gray-800">{viewingItem.account?.full_name}</div>
                      </div>
                      <div className="space-y-1">
                        <div className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Nomor Induk Karyawan (NIK)</div>
                        <div className="text-sm font-bold text-gray-700">{viewingItem.account?.internal_nik}</div>
                      </div>
                    </div>
                    <div className="space-y-4">
                      <div className="space-y-1">
                        <div className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Jabatan / Divisi</div>
                        <div className="text-sm font-bold text-gray-800">{viewingItem.account?.position} / {viewingItem.account?.department}</div>
                      </div>
                      <div className="space-y-1">
                        <div className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Lokasi Penugasan</div>
                        <div className="text-sm font-bold text-gray-700">{viewingItem.account?.location}</div>
                      </div>
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-12 flex-1">
                    <div className="space-y-6">
                      <div className="flex items-center gap-2 border-b-2 border-emerald-100 pb-2">
                        <div className="p-1.5 bg-emerald-100 text-emerald-600 rounded-lg"><DollarSign size={14} /></div>
                        <h3 className="text-xs font-black text-emerald-700 uppercase tracking-widest">Pendapatan</h3>
                      </div>
                      <div className="space-y-4">
                        <div className="flex justify-between items-start">
                          <div className="text-xs font-bold text-gray-700">Gaji Pokok ({viewingItem.salary_type})</div>
                          <div className="text-xs font-bold text-gray-800">Rp {viewingItem.basic_salary.toLocaleString('id-ID')}</div>
                        </div>
                        <div className="flex justify-between items-start">
                          <div className="text-xs font-bold text-gray-700">Tunjangan Jabatan</div>
                          <div className="text-xs font-bold text-gray-800">Rp {viewingItem.position_allowance.toLocaleString('id-ID')}</div>
                        </div>
                        <div className="flex justify-between items-start">
                          <div className="text-xs font-bold text-gray-700">Tunjangan Penempatan</div>
                          <div className="text-xs font-bold text-gray-800">Rp {viewingItem.placement_allowance.toLocaleString('id-ID')}</div>
                        </div>
                        <div className="flex justify-between items-start">
                          <div className="text-xs font-bold text-gray-700">Tunjangan Lainnya</div>
                          <div className="text-xs font-bold text-gray-800">Rp {viewingItem.other_allowance.toLocaleString('id-ID')}</div>
                        </div>
                        <div className="flex justify-between items-start">
                          <div className="text-xs font-bold text-gray-700">Tambahan Gaji Lain</div>
                          <div className="text-xs font-bold text-gray-800">Rp {viewingItem.other_additions.toLocaleString('id-ID')}</div>
                        </div>
                      </div>
                    </div>
                    <div className="space-y-6">
                      <div className="flex items-center gap-2 border-b-2 border-rose-100 pb-2">
                        <div className="p-1.5 bg-rose-100 text-rose-600 rounded-lg"><Trash2 size={14} /></div>
                        <h3 className="text-xs font-black text-rose-700 uppercase tracking-widest">Potongan</h3>
                      </div>
                      <div className="space-y-4">
                        <div className="flex justify-between items-start">
                          <div className="text-xs font-bold text-gray-700">Potongan Keterlambatan</div>
                          <div className="text-xs font-bold text-rose-600">Rp {viewingItem.late_deduction.toLocaleString('id-ID')}</div>
                        </div>
                        <div className="flex justify-between items-start">
                          <div className="text-xs font-bold text-gray-700">Potongan Pulang Cepat</div>
                          <div className="text-xs font-bold text-rose-600">Rp {viewingItem.early_leave_deduction.toLocaleString('id-ID')}</div>
                        </div>
                        <div className="flex justify-between items-start">
                          <div className="text-xs font-bold text-gray-700">Potongan Absensi</div>
                          <div className="text-xs font-bold text-rose-600">Rp {viewingItem.absent_deduction.toLocaleString('id-ID')}</div>
                        </div>
                        <div className="flex justify-between items-start">
                          <div className="text-xs font-bold text-gray-700">Potongan Lainnya</div>
                          <div className="text-xs font-bold text-rose-600">Rp {viewingItem.other_deductions.toLocaleString('id-ID')}</div>
                        </div>
                        <div className="flex justify-between items-start">
                          <div className="text-xs font-bold text-gray-700">BPJS Kesehatan</div>
                          <div className="text-xs font-bold text-rose-600">Rp {viewingItem.bpjs_kesehatan.toLocaleString('id-ID')}</div>
                        </div>
                        <div className="flex justify-between items-start">
                          <div className="text-xs font-bold text-gray-700">BPJS Ketenagakerjaan</div>
                          <div className="text-xs font-bold text-rose-600">Rp {viewingItem.bpjs_ketenagakerjaan.toLocaleString('id-ID')}</div>
                        </div>
                        <div className="flex justify-between items-start">
                          <div className="text-xs font-bold text-gray-700">PPh 21</div>
                          <div className="text-xs font-bold text-rose-600">Rp {viewingItem.pph21.toLocaleString('id-ID')}</div>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div className="mt-12 pt-8 border-t-2 border-gray-100">
                    <div className="grid grid-cols-2 gap-12">
                      <div className="space-y-4">
                        <div className="flex justify-between items-center px-4 py-2 bg-emerald-50 rounded-lg">
                          <span className="text-[10px] font-black text-emerald-700 uppercase tracking-widest">Total Pendapatan</span>
                          <span className="text-sm font-black text-emerald-700">Rp {viewingItem.total_income.toLocaleString('id-ID')}</span>
                        </div>
                        <div className="flex justify-between items-center px-4 py-2 bg-rose-50 rounded-lg">
                          <span className="text-[10px] font-black text-rose-700 uppercase tracking-widest">Total Potongan</span>
                          <span className="text-sm font-black text-rose-700">Rp {viewingItem.total_deduction.toLocaleString('id-ID')}</span>
                        </div>
                      </div>
                      <div className="flex flex-col items-end justify-center bg-[#006E62] text-white p-6 rounded-2xl shadow-xl">
                        <div className="text-[10px] font-bold uppercase tracking-[0.2em] opacity-80 mb-1">Take Home Pay</div>
                        <div className="text-3xl font-black tracking-tighter">Rp {viewingItem.take_home_pay.toLocaleString('id-ID')}</div>
                      </div>
                    </div>
                  </div>

                  <div className="mt-auto pt-16 grid grid-cols-3 gap-12 text-center">
                    <div className="space-y-20">
                      <div className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Diterima Oleh,</div>
                      <div className="border-t border-gray-200 pt-2 text-xs font-bold text-gray-800">{viewingItem.account?.full_name}</div>
                    </div>
                    <div></div>
                    <div className="space-y-20">
                      <div className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Disetujui Oleh,</div>
                      <div className="border-t border-gray-200 pt-2 text-xs font-bold text-gray-800">HR Manager</div>
                    </div>
                  </div>
               </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default MyPayslip;
