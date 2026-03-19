import React, { useState, useEffect, useRef } from 'react';
import { ArrowLeft, Search, Eye, Download, Trash2, Send, CheckCircle, AlertCircle, Printer, X, DollarSign, User, MapPin, Calendar, FileText } from 'lucide-react';
import { financeService } from '../../services/financeService';
import { Payroll, PayrollItem, PayrollSettings } from '../../types';
import { jsPDF } from 'jspdf';
import html2canvas from 'html2canvas';
import Swal from 'sweetalert2';

interface PayslipDetailProps {
  payroll: Payroll;
  onBack: () => void;
}

const PayslipDetail: React.FC<PayslipDetailProps> = ({ payroll, onBack }) => {
  const [items, setItems] = useState<PayrollItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedItems, setSelectedItems] = useState<string[]>([]);
  const [viewingItem, setViewingItem] = useState<PayrollItem | null>(null);
  const [settings, setSettings] = useState<PayrollSettings | null>(null);
  const payslipRef = useRef<HTMLDivElement>(null);

  const fetchItems = async () => {
    setLoading(true);
    try {
      const [iData, sData] = await Promise.all([
        financeService.getPayrollItems(payroll.id),
        financeService.getPayrollSettings()
      ]);
      setItems(iData);
      setSettings(sData);
    } catch (error) {
      console.error('Error fetching items:', error);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchItems();
  }, [payroll.id]);

  const handleStatusUpdate = async (status: Payroll['status']) => {
    try {
      await financeService.updatePayrollStatus(payroll.id, status);
      Swal.fire('Berhasil!', `Status payroll berhasil diubah menjadi ${status}.`, 'success');
      onBack();
    } catch (error) {
      Swal.fire('Gagal!', 'Terjadi kesalahan saat memperbarui status.', 'error');
    }
  };

  const handleDeleteItems = async (ids: string[]) => {
    const result = await Swal.fire({
      title: 'Hapus Payslip?',
      text: `${ids.length} data payslip akan dihapus permanen.`,
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
      confirmButtonText: 'Ya, Hapus'
    });

    if (result.isConfirmed) {
      try {
        await financeService.deletePayrollItems(ids);
        Swal.fire('Berhasil!', 'Data payslip berhasil dihapus.', 'success');
        fetchItems();
        setSelectedItems([]);
      } catch (error) {
        Swal.fire('Gagal!', 'Terjadi kesalahan saat menghapus data.', 'error');
      }
    }
  };

  const downloadPDF = async (item: PayrollItem) => {
    if (!payslipRef.current) return;
    
    try {
      Swal.fire({
        title: 'Generating PDF...',
        text: 'Mohon tunggu sebentar',
        allowOutsideClick: false,
        didOpen: () => {
          Swal.showLoading();
        }
      });

      const canvas = await html2canvas(payslipRef.current, {
        scale: 2,
        useCORS: true,
        logging: false,
        backgroundColor: '#ffffff',
        width: 720,
        onclone: (clonedDoc) => {
          // 1. Remove ALL style and link tags to completely avoid oklch parser errors
          const styleStuff = clonedDoc.querySelectorAll('style, link[rel="stylesheet"]');
          styleStuff.forEach(el => el.remove());

          // 2. Add a basic font-face for Inter to maintain typography
          const fontStyle = clonedDoc.createElement('style');
          fontStyle.textContent = `
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;700;900&display=swap');
            * { font-family: 'Inter', sans-serif !important; box-sizing: border-box; }
            body { margin: 0; padding: 0; background: white; }
          `;
          clonedDoc.head.appendChild(fontStyle);

          // 3. Isolate the capture element and force its dimensions
          const captureEl = clonedDoc.getElementById('payslip-capture-area');
          if (captureEl) {
            clonedDoc.body.innerHTML = '';
            clonedDoc.body.appendChild(captureEl);
            
            captureEl.style.position = 'absolute';
            captureEl.style.top = '0';
            captureEl.style.left = '0';
            captureEl.style.margin = '0';
            captureEl.style.width = '720px';
            captureEl.style.minHeight = '1018px'; // A4 aspect ratio for 720px width
            captureEl.style.padding = '60px'; // Balanced padding
          }

          // 4. Final check for any inline oklch
          const all = clonedDoc.querySelectorAll('*');
          all.forEach(el => {
            const htmlEl = el as HTMLElement;
            if (htmlEl.style && htmlEl.style.cssText.includes('oklch')) {
              htmlEl.style.cssText = htmlEl.style.cssText.replace(/oklch\([^)]+\)/g, '#000000');
            }
          });
        }
      });
      
      const imgData = canvas.toDataURL('image/png');
      const pdf = new jsPDF('p', 'mm', 'a4');
      const imgProps = pdf.getImageProperties(imgData);
      
      const pdfWidth = pdf.internal.pageSize.getWidth();
      const pdfHeight = pdf.internal.pageSize.getHeight();
      
      // Calculate dimensions to fit the image on the A4 page while maintaining aspect ratio
      const ratio = imgProps.width / imgProps.height;
      let finalWidth = pdfWidth;
      let finalHeight = pdfWidth / ratio;
      
      if (finalHeight > pdfHeight) {
        finalHeight = pdfHeight;
        finalWidth = pdfHeight * ratio;
      }
      
      // Center the image on the page
      const xOffset = (pdfWidth - finalWidth) / 2;
      const yOffset = (pdfHeight - finalHeight) / 2;
      
      pdf.addImage(imgData, 'PNG', xOffset, yOffset, finalWidth, finalHeight);
      pdf.save(`Payslip_${item.account?.full_name}_${new Date(0, payroll.month - 1).toLocaleString('id-ID', { month: 'long' })}${payroll.year}.pdf`);
      
      Swal.close();
    } catch (error) {
      console.error('PDF Error:', error);
      Swal.fire('Gagal!', 'Terjadi kesalahan saat membuat PDF.', 'error');
    }
  };

  const filteredItems = items.filter(i => 
    i.account?.full_name.toLowerCase().includes(searchQuery.toLowerCase()) ||
    i.account?.internal_nik.toLowerCase().includes(searchQuery.toLowerCase())
  );

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-center gap-4">
          <button
            onClick={onBack}
            className="p-2 hover:bg-gray-100 rounded-full transition-colors text-gray-500"
          >
            <ArrowLeft size={20} />
          </button>
          <div>
            <h2 className="text-2xl font-bold text-gray-800">
              Payslip {new Date(0, payroll.month - 1).toLocaleString('id-ID', { month: 'long' })} {payroll.year}
            </h2>
            <p className="text-sm text-gray-500">Daftar slip gaji karyawan untuk periode ini.</p>
          </div>
        </div>
        <div className="flex items-center gap-2">
          {payroll.status === 'Draft' && (
            <button
              onClick={() => handleStatusUpdate('Pending')}
              className="flex items-center gap-2 px-6 py-2.5 bg-amber-600 text-white rounded-lg hover:bg-amber-700 transition-all text-sm font-bold uppercase tracking-wider shadow-md"
            >
              <Send size={18} />
              Ajukan Verifikasi
            </button>
          )}
          {payroll.status === 'Pending' && (
            <button
              onClick={() => handleStatusUpdate('Approved')}
              className="flex items-center gap-2 px-6 py-2.5 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 transition-all text-sm font-bold uppercase tracking-wider shadow-md"
            >
              <CheckCircle size={18} />
              Setujui Payroll
            </button>
          )}
          {payroll.status === 'Approved' && (
            <button
              onClick={() => handleStatusUpdate('Paid')}
              className="flex items-center gap-2 px-6 py-2.5 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-all text-sm font-bold uppercase tracking-wider shadow-md"
            >
              <Send size={18} />
              Sent Payslip (Terbitkan)
            </button>
          )}
        </div>
      </div>

      <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="p-4 border-b border-gray-50 flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="relative flex-1 max-w-md">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
            <input
              type="text"
              placeholder="Cari karyawan..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full pl-10 pr-4 py-2 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all"
            />
          </div>
          {selectedItems.length > 0 && (
            <button
              onClick={() => handleDeleteItems(selectedItems)}
              className="flex items-center gap-2 px-4 py-2 bg-rose-50 text-rose-600 rounded-lg hover:bg-rose-100 transition-all text-xs font-bold uppercase tracking-wider"
            >
              <Trash2 size={16} />
              Hapus Terpilih ({selectedItems.length})
            </button>
          )}
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead>
              <tr className="bg-gray-50 text-gray-500 text-[10px] uppercase tracking-widest font-bold">
                <th className="px-6 py-4 w-10">
                  <input 
                    type="checkbox" 
                    checked={selectedItems.length === filteredItems.length && filteredItems.length > 0}
                    onChange={(e) => {
                      if (e.target.checked) setSelectedItems(filteredItems.map(i => i.id));
                      else setSelectedItems([]);
                    }}
                    className="rounded border-gray-300 text-[#006E62] focus:ring-[#006E62]"
                  />
                </th>
                <th className="px-6 py-4">Karyawan</th>
                <th className="px-6 py-4">Jabatan / Lokasi</th>
                <th className="px-6 py-4">Total Pendapatan</th>
                <th className="px-6 py-4">Total Potongan</th>
                <th className="px-6 py-4">Take Home Pay</th>
                <th className="px-6 py-4 text-right">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {loading ? (
                <tr>
                  <td colSpan={7} className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center gap-2">
                      <div className="w-8 h-8 border-4 border-[#006E62] border-t-transparent rounded-full animate-spin"></div>
                      <span className="text-xs font-bold text-gray-400 uppercase tracking-widest">Memuat Data...</span>
                    </div>
                  </td>
                </tr>
              ) : filteredItems.length === 0 ? (
                <tr>
                  <td colSpan={7} className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center gap-2 text-gray-400">
                      <DollarSign size={48} strokeWidth={1} />
                      <p className="text-sm font-medium">Belum ada data payslip.</p>
                    </div>
                  </td>
                </tr>
              ) : (
                filteredItems.map((item) => (
                  <tr key={item.id} className="hover:bg-gray-50 transition-colors group">
                    <td className="px-6 py-4">
                      <input 
                        type="checkbox" 
                        checked={selectedItems.includes(item.id)}
                        onChange={() => {
                          setSelectedItems(prev => 
                            prev.includes(item.id) ? prev.filter(i => i !== item.id) : [...prev, item.id]
                          );
                        }}
                        className="rounded border-gray-300 text-[#006E62] focus:ring-[#006E62]"
                      />
                    </td>
                    <td className="px-6 py-4">
                      <div className="font-bold text-gray-800">{item.account?.full_name}</div>
                      <div className="text-[10px] text-gray-400 uppercase tracking-widest font-medium">{item.account?.internal_nik}</div>
                    </td>
                    <td className="px-6 py-4">
                      <div className="text-xs text-gray-600">{item.account?.position}</div>
                      <div className="text-[10px] text-gray-400">{item.account?.location?.name || (item.account as any)?.location || '-'}</div>
                    </td>
                    <td className="px-6 py-4 font-mono text-xs text-emerald-600 font-bold">
                      Rp {item.total_income.toLocaleString('id-ID')}
                    </td>
                    <td className="px-6 py-4 font-mono text-xs text-rose-600 font-bold">
                      Rp {item.total_deduction.toLocaleString('id-ID')}
                    </td>
                    <td className="px-6 py-4 font-mono text-sm text-[#006E62] font-bold">
                      Rp {item.take_home_pay.toLocaleString('id-ID')}
                    </td>
                    <td className="px-6 py-4 text-right">
                      <div className="flex items-center justify-end gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                        <button
                          onClick={() => setViewingItem(item)}
                          className="p-2 text-gray-400 hover:text-[#006E62] hover:bg-emerald-50 rounded-lg transition-colors"
                          title="Lihat Slip Gaji"
                        >
                          <Eye size={18} />
                        </button>
                        <button
                          onClick={() => downloadPDF(item)}
                          className="p-2 text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                          title="Download PDF"
                        >
                          <Download size={18} />
                        </button>
                        <button
                          onClick={() => handleDeleteItems([item.id])}
                          className="p-2 text-red-600 hover:bg-red-50 rounded-lg transition-colors"
                          title="Hapus"
                        >
                          <Trash2 size={18} />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Payslip Modal */}
      {viewingItem && (
        <div className="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-4xl max-h-[90vh] flex flex-col">
            <div className="p-4 border-b border-gray-100 flex items-center justify-between bg-gray-50 rounded-t-2xl">
              <div className="flex items-center gap-2">
                <FileText className="text-[#006E62]" size={20} />
                <h3 className="font-bold text-gray-800 uppercase tracking-widest text-xs">Preview Slip Gaji</h3>
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

            <div className="flex-1 overflow-y-auto p-8 bg-gray-100 custom-scrollbar">
              <div 
                ref={payslipRef}
                id="payslip-capture-area"
                className="mx-auto flex flex-col"
                style={{ 
                  fontFamily: "'Inter', sans-serif",
                  backgroundColor: '#ffffff',
                  width: '720px',
                  minHeight: '1018px',
                  boxSizing: 'border-box',
                  position: 'relative',
                  padding: '60px'
                }}
              >
                {/* Header / Kop */}
                <div className="flex justify-between items-start border-b-4 pb-6 mb-8" style={{ borderBottomColor: '#006E62', display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', borderBottomWidth: '4px', paddingBottom: '24px', marginBottom: '32px' }}>
                  <div className="flex items-center gap-6" style={{ display: 'flex', alignItems: 'center', gap: '24px' }}>
                    {settings?.company_logo_url && (
                      <img 
                        src={settings.company_logo_url} 
                        alt="Logo" 
                        className="w-24 h-24 object-contain"
                        style={{ width: '96px', height: '96px', objectFit: 'contain' }}
                        referrerPolicy="no-referrer"
                      />
                    )}
                    <div className="space-y-1" style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                      <h1 className="text-3xl font-black tracking-tighter uppercase" style={{ color: '#006E62', fontSize: '30px', fontWeight: '900', margin: 0 }}>{settings?.company_name || 'HUREMA HRIS'}</h1>
                      <div className="text-xs max-w-md leading-relaxed" style={{ color: '#6b7280', fontSize: '12px', maxWidth: '448px' }}>
                        {settings?.company_address && <div className="flex items-center gap-2" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}><MapPin size={10} /> {settings.company_address}</div>}
                        <div className="flex items-center gap-4 mt-1" style={{ display: 'flex', alignItems: 'center', gap: '16px', marginTop: '4px' }}>
                          {settings?.company_phone && <span>Telp: {settings.company_phone}</span>}
                          {settings?.company_email && <span>Email: {settings.company_email}</span>}
                        </div>
                      </div>
                    </div>
                  </div>
                  <div className="text-right" style={{ textAlign: 'right' }}>
                    <h2 className="text-2xl font-black uppercase tracking-widest" style={{ color: '#d1d5db', fontSize: '24px', fontWeight: '900', margin: 0 }}>Slip Gaji</h2>
                    <div className="text-sm font-bold mt-2" style={{ color: '#1f2937', fontSize: '14px', fontWeight: '700' }}>
                      {new Date(0, payroll.month - 1).toLocaleString('id-ID', { month: 'long' })} {payroll.year}
                    </div>
                  </div>
                </div>

                {/* Employee Info */}
                <div className="mb-12 p-6 rounded-2xl border" style={{ backgroundColor: '#f9fafb', borderColor: '#f3f4f6', display: 'flex', justifyContent: 'space-between', gap: '48px', padding: '24px', borderRadius: '16px', borderStyle: 'solid', borderWidth: '1px', marginBottom: '48px', width: '100%', boxSizing: 'border-box' }}>
                  <div className="space-y-4" style={{ display: 'flex', flexDirection: 'column', gap: '16px', flex: 1 }}>
                    <div className="space-y-1" style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                      <div className="text-[10px] font-bold uppercase tracking-widest" style={{ color: '#9ca3af', fontSize: '10px', fontWeight: '700' }}>Nama Pegawai</div>
                      <div className="text-lg font-black" style={{ color: '#1f2937', fontSize: '18px', fontWeight: '900' }}>{viewingItem.account?.full_name}</div>
                    </div>
                    <div className="space-y-1" style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                      <div className="text-[10px] font-bold uppercase tracking-widest" style={{ color: '#9ca3af', fontSize: '10px', fontWeight: '700' }}>Nomor Induk Karyawan (NIK)</div>
                      <div className="text-sm font-bold" style={{ color: '#374151', fontSize: '14px', fontWeight: '700' }}>{viewingItem.account?.internal_nik}</div>
                    </div>
                  </div>
                  <div className="space-y-4" style={{ display: 'flex', flexDirection: 'column', gap: '16px', flex: 1 }}>
                    <div className="space-y-1" style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                      <div className="text-[10px] font-bold uppercase tracking-widest" style={{ color: '#9ca3af', fontSize: '10px', fontWeight: '700' }}>Golongan / Jabatan</div>
                      <div className="text-sm font-bold" style={{ color: '#1f2937', fontSize: '14px', fontWeight: '700' }}>{viewingItem.account?.grade || '-'} / {viewingItem.account?.position}</div>
                    </div>
                    <div className="space-y-1" style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                      <div className="text-[10px] font-bold uppercase tracking-widest" style={{ color: '#9ca3af', fontSize: '10px', fontWeight: '700' }}>Lokasi Penugasan</div>
                      <div className="text-sm font-bold" style={{ color: '#374151', fontSize: '14px', fontWeight: '700' }}>{viewingItem.account?.location?.name || (viewingItem.account as any)?.location || '-'}</div>
                    </div>
                  </div>
                </div>

                {/* Salary Components */}
                <div className="flex-1" style={{ display: 'flex', justifyContent: 'space-between', gap: '48px', flex: 1, width: '100%', boxSizing: 'border-box' }}>
                  {/* Earnings */}
                  <div className="space-y-6" style={{ display: 'flex', flexDirection: 'column', gap: '24px', flex: 1 }}>
                    <div className="flex items-center gap-2 border-b-2 pb-2" style={{ borderColor: '#d1fae5', display: 'flex', alignItems: 'center', gap: '8px', borderBottomStyle: 'solid', borderBottomWidth: '2px', paddingBottom: '8px' }}>
                      <div className="p-1.5 rounded-lg" style={{ backgroundColor: '#d1fae5', color: '#059669', padding: '6px' }}><DollarSign size={14} /></div>
                      <h3 className="text-xs font-black uppercase tracking-widest" style={{ color: '#047857', fontSize: '12px', fontWeight: '900' }}>Pendapatan (Earnings)</h3>
                    </div>
                    <div className="space-y-4" style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                      <div className="flex justify-between items-start" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                        <div className="space-y-0.5" style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                          <div className="text-xs font-bold" style={{ color: '#374151', fontSize: '12px', fontWeight: '700' }}>Gaji Pokok ({viewingItem.salary_type})</div>
                          {viewingItem.basic_salary_notes && <div className="text-[9px] italic" style={{ color: '#9ca3af', fontSize: '9px' }}>{viewingItem.basic_salary_notes}</div>}
                        </div>
                        <div className="text-xs font-bold" style={{ color: '#1f2937', fontSize: '12px', fontWeight: '700' }}>Rp {viewingItem.basic_salary.toLocaleString('id-ID')}</div>
                      </div>
                      <div className="flex justify-between items-start" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                        <div className="space-y-0.5" style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                          <div className="text-xs font-bold" style={{ color: '#374151', fontSize: '12px', fontWeight: '700' }}>Tunjangan Jabatan</div>
                          {viewingItem.position_allowance_notes && <div className="text-[9px] italic" style={{ color: '#9ca3af', fontSize: '9px' }}>{viewingItem.position_allowance_notes}</div>}
                        </div>
                        <div className="text-xs font-bold" style={{ color: '#1f2937', fontSize: '12px', fontWeight: '700' }}>Rp {viewingItem.position_allowance.toLocaleString('id-ID')}</div>
                      </div>
                      <div className="flex justify-between items-start" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                        <div className="space-y-0.5" style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                          <div className="text-xs font-bold" style={{ color: '#374151', fontSize: '12px', fontWeight: '700' }}>Tunjangan Penempatan</div>
                          {viewingItem.placement_allowance_notes && <div className="text-[9px] italic" style={{ color: '#9ca3af', fontSize: '9px' }}>{viewingItem.placement_allowance_notes}</div>}
                        </div>
                        <div className="text-xs font-bold" style={{ color: '#1f2937', fontSize: '12px', fontWeight: '700' }}>Rp {viewingItem.placement_allowance.toLocaleString('id-ID')}</div>
                      </div>
                      <div className="flex justify-between items-start" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                        <div className="space-y-0.5" style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                          <div className="text-xs font-bold" style={{ color: '#374151', fontSize: '12px', fontWeight: '700' }}>Tunjangan Lainnya</div>
                          {viewingItem.other_allowance_notes && <div className="text-[9px] italic" style={{ color: '#9ca3af', fontSize: '9px' }}>{viewingItem.other_allowance_notes}</div>}
                        </div>
                        <div className="text-xs font-bold" style={{ color: '#1f2937', fontSize: '12px', fontWeight: '700' }}>Rp {viewingItem.other_allowance.toLocaleString('id-ID')}</div>
                      </div>
                      <div className="flex justify-between items-start" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                        <div className="space-y-0.5" style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                          <div className="text-xs font-bold" style={{ color: '#374151', fontSize: '12px', fontWeight: '700' }}>Tambahan Gaji Lain</div>
                          {viewingItem.other_additions_notes && <div className="text-[9px] italic" style={{ color: '#9ca3af', fontSize: '9px' }}>{viewingItem.other_additions_notes}</div>}
                        </div>
                        <div className="text-xs font-bold" style={{ color: '#1f2937', fontSize: '12px', fontWeight: '700' }}>Rp {viewingItem.other_additions.toLocaleString('id-ID')}</div>
                      </div>
                    </div>
                  </div>

                  {/* Deductions */}
                  <div className="space-y-6" style={{ display: 'flex', flexDirection: 'column', gap: '24px', flex: 1 }}>
                    <div className="flex items-center gap-2 border-b-2 pb-2" style={{ borderColor: '#ffe4e6', display: 'flex', alignItems: 'center', gap: '8px', borderBottomStyle: 'solid', borderBottomWidth: '2px', paddingBottom: '8px' }}>
                      <div className="p-1.5 rounded-lg" style={{ backgroundColor: '#ffe4e6', color: '#e11d48', padding: '6px' }}><Trash2 size={14} /></div>
                      <h3 className="text-xs font-black uppercase tracking-widest" style={{ color: '#be123c', fontSize: '12px', fontWeight: '900' }}>Potongan (Deductions)</h3>
                    </div>
                    <div className="space-y-4" style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                      <div className="flex justify-between items-start" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                        <div className="space-y-0.5" style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                          <div className="text-xs font-bold" style={{ color: '#374151', fontSize: '12px', fontWeight: '700' }}>Potongan Keterlambatan</div>
                          {viewingItem.late_deduction_notes && <div className="text-[9px] italic" style={{ color: '#9ca3af', fontSize: '9px' }}>{viewingItem.late_deduction_notes}</div>}
                        </div>
                        <div className="text-xs font-bold" style={{ color: '#e11d48', fontSize: '12px', fontWeight: '700' }}>Rp {viewingItem.late_deduction.toLocaleString('id-ID')}</div>
                      </div>
                      <div className="flex justify-between items-start" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                        <div className="space-y-0.5" style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                          <div className="text-xs font-bold" style={{ color: '#374151', fontSize: '12px', fontWeight: '700' }}>Potongan Pulang Cepat</div>
                          {viewingItem.early_leave_deduction_notes && <div className="text-[9px] italic" style={{ color: '#9ca3af', fontSize: '9px' }}>{viewingItem.early_leave_deduction_notes}</div>}
                        </div>
                        <div className="text-xs font-bold" style={{ color: '#e11d48', fontSize: '12px', fontWeight: '700' }}>Rp {viewingItem.early_leave_deduction.toLocaleString('id-ID')}</div>
                      </div>
                      <div className="flex justify-between items-start" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                        <div className="space-y-0.5" style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                          <div className="text-xs font-bold" style={{ color: '#374151', fontSize: '12px', fontWeight: '700' }}>Potongan Absensi</div>
                          {viewingItem.absent_deduction_notes && <div className="text-[9px] italic" style={{ color: '#9ca3af', fontSize: '9px' }}>{viewingItem.absent_deduction_notes}</div>}
                        </div>
                        <div className="text-xs font-bold" style={{ color: '#e11d48', fontSize: '12px', fontWeight: '700' }}>Rp {viewingItem.absent_deduction.toLocaleString('id-ID')}</div>
                      </div>
                      <div className="flex justify-between items-start" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                        <div className="space-y-0.5" style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                          <div className="text-xs font-bold" style={{ color: '#374151', fontSize: '12px', fontWeight: '700' }}>Potongan Lainnya</div>
                          {viewingItem.other_deductions_notes && <div className="text-[9px] italic" style={{ color: '#9ca3af', fontSize: '9px' }}>{viewingItem.other_deductions_notes}</div>}
                        </div>
                        <div className="text-xs font-bold" style={{ color: '#e11d48', fontSize: '12px', fontWeight: '700' }}>Rp {viewingItem.other_deductions.toLocaleString('id-ID')}</div>
                      </div>
                      <div className="flex justify-between items-start" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                        <div className="text-xs font-bold" style={{ color: '#374151', fontSize: '12px', fontWeight: '700' }}>BPJS Kesehatan</div>
                        <div className="text-xs font-bold" style={{ color: '#e11d48', fontSize: '12px', fontWeight: '700' }}>Rp {viewingItem.bpjs_kesehatan.toLocaleString('id-ID')}</div>
                      </div>
                      <div className="flex justify-between items-start" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                        <div className="text-xs font-bold" style={{ color: '#374151', fontSize: '12px', fontWeight: '700' }}>BPJS Ketenagakerjaan</div>
                        <div className="text-xs font-bold" style={{ color: '#e11d48', fontSize: '12px', fontWeight: '700' }}>Rp {viewingItem.bpjs_ketenagakerjaan.toLocaleString('id-ID')}</div>
                      </div>
                      <div className="flex justify-between items-start" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                        <div className="text-xs font-bold" style={{ color: '#374151', fontSize: '12px', fontWeight: '700' }}>PPh 21</div>
                        <div className="text-xs font-bold" style={{ color: '#e11d48', fontSize: '12px', fontWeight: '700' }}>Rp {viewingItem.pph21.toLocaleString('id-ID')}</div>
                      </div>
                    </div>
                  </div>
                </div>

                {/* Summary */}
                <div className="mt-12 pt-8 border-t-2" style={{ borderColor: '#f3f4f6', borderTopStyle: 'solid', borderTopWidth: '2px', paddingTop: '32px', marginTop: '48px', width: '100%', boxSizing: 'border-box' }}>
                  <div className="gap-12" style={{ display: 'flex', justifyContent: 'space-between', gap: '48px' }}>
                    <div className="space-y-4" style={{ display: 'flex', flexDirection: 'column', gap: '16px', flex: 1 }}>
                      <div className="flex justify-between items-center px-4 py-2 rounded-lg" style={{ backgroundColor: '#ecfdf5', display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 16px', borderRadius: '8px' }}>
                        <span className="text-[10px] font-black uppercase tracking-widest" style={{ color: '#047857', fontSize: '10px', fontWeight: '900' }}>Total Pendapatan</span>
                        <span className="text-sm font-black" style={{ color: '#047857', fontSize: '14px', fontWeight: '900' }}>Rp {viewingItem.total_income.toLocaleString('id-ID')}</span>
                      </div>
                      <div className="flex justify-between items-center px-4 py-2 rounded-lg" style={{ backgroundColor: '#fff1f2', display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 16px', borderRadius: '8px' }}>
                        <span className="text-[10px] font-black uppercase tracking-widest" style={{ color: '#be123c', fontSize: '10px', fontWeight: '900' }}>Total Potongan</span>
                        <span className="text-sm font-black" style={{ color: '#be123c', fontSize: '14px', fontWeight: '900' }}>Rp {viewingItem.total_deduction.toLocaleString('id-ID')}</span>
                      </div>
                    </div>
                    <div className="flex flex-col items-end justify-center text-white p-6 rounded-2xl shadow-xl" style={{ backgroundColor: '#006E62', display: 'flex', flexDirection: 'column', alignItems: 'flex-end', justifyContent: 'center', padding: '24px', borderRadius: '16px', color: '#ffffff', flex: 1 }}>
                      <div className="text-[10px] font-bold uppercase tracking-[0.2em] opacity-80 mb-1" style={{ fontSize: '10px', fontWeight: '700', opacity: 0.8, marginBottom: '4px' }}>Take Home Pay</div>
                      <div className="text-3xl font-black tracking-tighter" style={{ fontSize: '30px', fontWeight: '900' }}>Rp {viewingItem.take_home_pay.toLocaleString('id-ID')}</div>
                    </div>
                  </div>
                </div>

                {/* Footer / Signatures */}
                <div className="mt-auto pt-16 grid grid-cols-3 gap-12 text-center" style={{ marginTop: 'auto', paddingTop: '64px', display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '48px', textAlign: 'center' }}>
                  <div className="space-y-20" style={{ display: 'flex', flexDirection: 'column', gap: '80px' }}>
                    <div className="text-[10px] font-bold uppercase tracking-widest" style={{ color: '#9ca3af', fontSize: '10px' }}>Diterima Oleh,</div>
                    <div className="border-t pt-2 text-xs font-bold" style={{ borderColor: '#e5e7eb', color: '#1f2937', borderTopWidth: '1px', paddingTop: '8px', fontSize: '12px' }}>{viewingItem.account?.full_name}</div>
                  </div>
                  <div></div>
                  <div className="space-y-20" style={{ display: 'flex', flexDirection: 'column', gap: '80px' }}>
                    <div className="text-[10px] font-bold uppercase tracking-widest" style={{ color: '#9ca3af', fontSize: '10px' }}>Disetujui Oleh,</div>
                    <div className="border-t pt-2 text-xs font-bold" style={{ borderColor: '#e5e7eb', color: '#1f2937', borderTopWidth: '1px', paddingTop: '8px', fontSize: '12px' }}>{payroll.verifier?.full_name || 'HR Manager'}</div>
                  </div>
                </div>

                <div className="mt-12 text-[8px] text-center uppercase tracking-[0.3em]" style={{ color: '#d1d5db', marginTop: '48px', fontSize: '8px', textAlign: 'center' }}>
                  Dokumen ini dihasilkan secara otomatis oleh sistem HUREMA HRIS • {new Date().toLocaleString('id-ID')}
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default PayslipDetail;
