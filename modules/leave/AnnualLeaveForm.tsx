import React, { useState } from 'react';
import { X, Save, Calendar, FileText, Upload, Info, XCircle } from 'lucide-react';
import { AnnualLeaveRequestInput } from '../../types';
import { googleDriveService } from '../../services/googleDriveService';

interface AnnualLeaveFormProps {
  accountId: string;
  leaveQuota: number;
  usedDays: number;
  onClose: () => void;
  onSubmit: (data: AnnualLeaveRequestInput) => void;
}

const AnnualLeaveForm: React.FC<AnnualLeaveFormProps> = ({ accountId, leaveQuota, usedDays, onClose, onSubmit }) => {
  const [formData, setFormData] = useState<AnnualLeaveRequestInput>({
    account_id: accountId,
    start_date: new Date().toISOString().split('T')[0],
    end_date: new Date().toISOString().split('T')[0],
    description: '',
    file_id: null
  });

  const [uploading, setUploading] = useState(false);

  const calculateDuration = (start: string, end: string) => {
    const s = new Date(start);
    const e = new Date(end);
    if (isNaN(s.getTime()) || isNaN(e.getTime())) return 0;
    const diffTime = Math.abs(e.getTime() - s.getTime());
    return Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
  };

  const duration = calculateDuration(formData.start_date, formData.end_date);
  const remainingQuota = leaveQuota - usedDays;
  const isOverQuota = duration > remainingQuota;

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    const { name, value } = e.target;
    setFormData(prev => ({ ...prev, [name]: value }));
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      setUploading(true);
      const fileId = await googleDriveService.uploadFile(file);
      setFormData(prev => ({ ...prev, file_id: fileId }));
    } catch (error) {
      alert('Gagal mengunggah file.');
    } finally {
      setUploading(false);
    }
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (new Date(formData.start_date) > new Date(formData.end_date)) {
      alert('Tanggal awal tidak boleh lebih besar dari tanggal akhir.');
      return;
    }
    onSubmit(formData);
  };

  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 p-4 backdrop-blur-sm">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md overflow-hidden animate-in zoom-in duration-200">
        <div className="px-6 py-4 border-b border-gray-50 flex items-center justify-between bg-gray-50/50">
          <div className="flex items-center gap-2">
            <Calendar size={18} className="text-[#006E62]" />
            <h3 className="text-sm font-bold text-gray-800 uppercase tracking-tight">Form Pengajuan Cuti</h3>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={20} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-6 space-y-5">
          <div className="p-4 bg-[#006E62]/5 border border-[#006E62]/10 rounded-2xl flex items-center justify-between">
            <div>
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-0.5">Sisa Jatah Cuti</p>
              <p className="text-lg font-bold text-[#006E62]">{remainingQuota} <span className="text-xs font-medium text-gray-400">Hari</span></p>
            </div>
            <div className="text-right">
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-0.5">Durasi Pengajuan</p>
              <p className={`text-lg font-bold ${isOverQuota ? 'text-rose-600' : 'text-gray-800'}`}>{duration} <span className="text-xs font-medium text-gray-400">Hari</span></p>
            </div>
          </div>

          {isOverQuota && (
            <div className="p-3 bg-rose-50 border border-rose-100 rounded-xl flex gap-3 items-center animate-in slide-in-from-top-2 duration-300">
              <XCircle size={18} className="text-rose-500 shrink-0" />
              <p className="text-[10px] text-rose-700 font-bold uppercase leading-tight">
                Durasi pengajuan melebihi sisa jatah cuti Anda!
              </p>
            </div>
          )}

          <div className="grid grid-cols-2 gap-4">
            <div className="space-y-1.5">
              <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Tanggal Awal</label>
              <input 
                required
                type="date" 
                name="start_date" 
                value={formData.start_date} 
                onChange={handleChange}
                className="w-full px-3 py-2.5 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] text-xs font-bold"
              />
            </div>
            <div className="space-y-1.5">
              <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Tanggal Akhir</label>
              <input 
                required
                type="date" 
                name="end_date" 
                value={formData.end_date} 
                onChange={handleChange}
                className="w-full px-3 py-2.5 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] text-xs font-bold"
              />
            </div>
          </div>

          <div className="space-y-1.5">
            <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Alasan Cuti</label>
            <textarea 
              required
              name="description" 
              value={formData.description} 
              onChange={handleChange}
              rows={3}
              placeholder="Jelaskan alasan pengajuan cuti Anda..."
              className="w-full px-3 py-2.5 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] text-xs font-medium resize-none"
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Dokumen Pendukung (Opsional)</label>
            <div className="flex items-center gap-3">
              <label className="flex-1 flex items-center gap-2 px-4 py-2.5 bg-gray-50 border border-dashed border-gray-200 rounded-xl cursor-pointer hover:bg-gray-100 transition-all group overflow-hidden">
                <Upload size={16} className="text-gray-400 group-hover:text-[#006E62]" />
                <span className="text-[10px] font-bold text-gray-500 uppercase truncate">
                  {formData.file_id ? 'File Terunggah' : 'Pilih Gambar/PDF'}
                </span>
                <input type="file" className="hidden" accept="image/*,application/pdf" onChange={handleFileUpload} />
              </label>
              {uploading && <div className="w-5 h-5 border-2 border-[#006E62] border-t-transparent rounded-full animate-spin"></div>}
            </div>
          </div>

          <div className="p-3 bg-amber-50 border border-amber-100 rounded-xl flex gap-3 items-start">
            <Info size={16} className="text-amber-500 shrink-0 mt-0.5" />
            <p className="text-[9px] text-amber-700 leading-relaxed font-medium">
              Pengajuan ini akan masuk ke antrean verifikasi Admin. Admin dapat menyetujui, menolak, atau menawarkan negosiasi tanggal jika diperlukan.
            </p>
          </div>

          <div className="flex gap-3 pt-2">
            <button 
              type="button" 
              onClick={onClose}
              className="flex-1 py-3 border border-gray-100 text-gray-400 rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-gray-50 transition-all"
            >
              Batal
            </button>
            <button 
              type="submit"
              disabled={uploading || isOverQuota || duration <= 0}
              className="flex-2 py-3 bg-[#006E62] text-white rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-[#005a50] shadow-lg shadow-[#006E62]/20 transition-all flex items-center justify-center gap-2 disabled:opacity-50"
            >
              <Save size={16} />
              Kirim Pengajuan
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};

export default AnnualLeaveForm;
