
import React, { useState } from 'react';
import { X, Save, Upload, Loader2, Receipt, Calendar, DollarSign, FileText, CheckCircle2 } from 'lucide-react';
import { financeService } from '../../services/financeService';
import { googleDriveService } from '../../services/googleDriveService';
import { Compensation } from '../../types';
import Swal from 'sweetalert2';

interface CompensationDetailProps {
  compensation: Compensation;
  onClose: () => void;
  onSuccess: () => void;
}

const CompensationDetail: React.FC<CompensationDetailProps> = ({ compensation, onClose, onSuccess }) => {
  const [formData, setFormData] = useState({
    transaction_date: new Date().toISOString().split('T')[0],
    processed_amount: compensation.amount,
    notes: '',
    proof_file_id: ''
  });

  const [uploading, setUploading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      setUploading(true);
      const fileId = await googleDriveService.uploadFile(file);
      setFormData(prev => ({ ...prev, proof_file_id: fileId }));
    } catch (e) {
      Swal.fire('Gagal', 'Gagal mengunggah bukti pembayaran.', 'error');
    } finally {
      setUploading(false);
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    
    const result = await Swal.fire({
      title: 'Konfirmasi Pembayaran?',
      text: "Pastikan data transaksi sudah benar. Status akan diubah menjadi Selesai.",
      icon: 'question',
      showCancelButton: true,
      confirmButtonColor: '#006E62',
      confirmButtonText: 'Ya, Selesaikan'
    });

    if (result.isConfirmed) {
      try {
        setIsSaving(true);
        await financeService.updateCompensation(compensation.id, {
          ...formData,
          status: 'Completed'
        });
        Swal.fire({ title: 'Berhasil!', text: 'Data kompensasi telah diproses.', icon: 'success', timer: 1500, showConfirmButton: false });
        onSuccess();
      } catch (error) {
        console.error('Error updating compensation:', error);
        Swal.fire('Gagal', 'Terjadi kesalahan saat menyimpan data.', 'error');
      } finally {
        setIsSaving(false);
      }
    }
  };

  return (
    <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black bg-opacity-40 p-4">
      <div className="bg-white rounded-xl shadow-2xl w-full max-w-lg overflow-hidden animate-in zoom-in duration-200">
        <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
          <div>
            <h3 className="text-base font-bold text-gray-800">Detail Kompensasi</h3>
            <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Verifikasi & Proses Pembayaran</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors"><X size={20} /></button>
        </div>

        <div className="p-6 space-y-6 max-h-[80vh] overflow-y-auto scrollbar-none">
          {/* Info Section */}
          <div className="grid grid-cols-2 gap-4 bg-gray-50 p-4 rounded-lg border border-gray-100">
            <div>
              <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Nama Pegawai</p>
              <p className="text-sm font-bold text-gray-800">{compensation.account?.full_name}</p>
              <p className="text-xs text-gray-500">{compensation.account?.internal_nik}</p>
            </div>
            <div>
              <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Tipe & Tanggal Exit</p>
              <p className="text-sm font-bold text-gray-800">{compensation.termination_type}</p>
              <p className="text-xs text-gray-500">{new Date(compensation.termination_date).toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' })}</p>
            </div>
            <div className="col-span-2 pt-2 border-t border-gray-200 mt-2">
              <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Jenis & Estimasi Nominal</p>
              <div className="flex items-center justify-between">
                <span className={`px-2 py-0.5 rounded text-[10px] font-bold uppercase ${compensation.type === 'Severance' ? 'bg-emerald-100 text-emerald-700' : 'bg-orange-100 text-orange-700'}`}>
                  {compensation.type === 'Severance' ? 'Pesangon' : 'Pinalti'}
                </span>
                <span className="text-lg font-mono font-bold text-gray-900">Rp {compensation.amount.toLocaleString('id-ID')}</span>
              </div>
            </div>
          </div>

          {compensation.status === 'Completed' ? (
            <div className="space-y-4">
              <div className="bg-emerald-50 p-4 rounded-lg border border-emerald-100 flex items-center gap-3">
                <CheckCircle2 className="text-emerald-600" size={24} />
                <div>
                  <p className="text-sm font-bold text-emerald-800">Transaksi Selesai</p>
                  <p className="text-[10px] text-emerald-600 font-medium uppercase tracking-widest">Telah Diverifikasi oleh Finance</p>
                </div>
              </div>
              
              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1">
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tanggal Transaksi</label>
                  <div className="px-3 py-2 bg-gray-50 border border-gray-200 rounded text-xs text-gray-600">
                    {new Date(compensation.transaction_date!).toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' })}
                  </div>
                </div>
                <div className="space-y-1">
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Nominal Dibayarkan</label>
                  <div className="px-3 py-2 bg-gray-50 border border-gray-200 rounded text-xs font-bold text-gray-800">
                    Rp {compensation.processed_amount?.toLocaleString('id-ID')}
                  </div>
                </div>
              </div>

              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Catatan Finance</label>
                <div className="px-3 py-2 bg-gray-50 border border-gray-200 rounded text-xs text-gray-600 italic">
                  {compensation.notes || 'Tidak ada catatan.'}
                </div>
              </div>

              {compensation.proof_file_id && (
                <div className="space-y-1">
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Bukti Pembayaran</label>
                  <a 
                    href={googleDriveService.getFileUrl(compensation.proof_file_id)}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="flex items-center gap-2 p-3 bg-emerald-50 text-emerald-700 rounded-lg border border-emerald-100 hover:bg-emerald-100 transition-all"
                  >
                    <FileText size={16} />
                    <span className="text-xs font-bold uppercase tracking-widest">Lihat Bukti Transfer</span>
                  </a>
                </div>
              )}
            </div>
          ) : (
            <form onSubmit={handleSubmit} className="space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1">
                  <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Tanggal Transaksi</label>
                  <div className="relative">
                    <Calendar className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={14} />
                    <input 
                      type="date"
                      required
                      value={formData.transaction_date}
                      onChange={(e) => setFormData({...formData, transaction_date: e.target.value})}
                      className="w-full pl-9 pr-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none"
                    />
                  </div>
                </div>
                <div className="space-y-1">
                  <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Nominal Aktual (IDR)</label>
                  <div className="relative">
                    <DollarSign className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={14} />
                    <input 
                      type="number"
                      required
                      value={formData.processed_amount}
                      onChange={(e) => setFormData({...formData, processed_amount: Number(e.target.value)})}
                      className="w-full pl-9 pr-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none font-bold"
                    />
                  </div>
                </div>
              </div>

              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Catatan / Keterangan</label>
                <textarea 
                  rows={3}
                  value={formData.notes}
                  onChange={(e) => setFormData({...formData, notes: e.target.value})}
                  placeholder="Tambahkan catatan jika nominal berbeda atau info transfer..."
                  className="w-full px-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none resize-none"
                />
              </div>

              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Bukti Pembayaran (Optional)</label>
                <div className={`p-3 bg-gray-50 border border-dashed rounded-md ${formData.proof_file_id ? 'border-[#006E62] bg-emerald-50/20' : 'border-gray-200'}`}>
                  <label className="flex items-center gap-3 cursor-pointer">
                    <Upload size={16} className={formData.proof_file_id ? 'text-[#006E62]' : 'text-gray-300'} />
                    <div className="flex-1">
                      <p className="text-[10px] font-bold text-gray-600 uppercase">{uploading ? 'Mengunggah...' : formData.proof_file_id ? 'BUKTI TERUNGGAH' : 'UNGGAH BUKTI TRANSFER'}</p>
                    </div>
                    <input type="file" className="hidden" accept="image/*,application/pdf" onChange={handleFileUpload} disabled={uploading} />
                  </label>
                </div>
              </div>

              <div className="flex justify-end gap-3 pt-4">
                <button type="button" onClick={onClose} className="px-4 py-2 text-xs font-bold text-gray-500 uppercase">Batal</button>
                <button 
                  disabled={isSaving || uploading}
                  className="flex items-center gap-2 bg-[#006E62] text-white px-8 py-2 rounded shadow-md hover:bg-[#005a50] transition-all text-xs font-bold uppercase disabled:opacity-50"
                >
                  {isSaving ? <Loader2 size={14} className="animate-spin" /> : <Save size={14} />} Selesaikan Pembayaran
                </button>
              </div>
            </form>
          )}
        </div>
      </div>
    </div>
  );
};

export default CompensationDetail;
