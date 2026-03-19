import React, { useState, useEffect } from 'react';
import { X, Save, Upload, Calendar, FileText, ClipboardList, Clock, Loader2 } from 'lucide-react';
import Swal from 'sweetalert2';
import { submissionService } from '../../services/submissionService';
import { authService } from '../../services/authService';
import { googleDriveService } from '../../services/googleDriveService';

interface SubmissionFormProps {
  onClose: () => void;
  onSuccess: () => void;
}

const SubmissionForm: React.FC<SubmissionFormProps> = ({ onClose, onSuccess }) => {
  const [formData, setFormData] = useState<any>({
    type: 'Cuti',
    description: '',
    file_id: '',
    submission_data: {
      start_date: '',
      end_date: '',
      duration_days: 0,
      leave_type: 'Tahunan'
    }
  });

  const [uploading, setUploading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const currentUser = authService.getCurrentUser();

  const handleTypeChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const type = e.target.value;
    let initialData = {};
    
    if (type === 'Cuti') {
      initialData = { start_date: '', end_date: '', duration_days: 0, leave_type: 'Tahunan' };
    } else if (type === 'Lembur') {
      initialData = { date: '', start_time: '', end_time: '' };
    } else if (type === 'Koreksi Absen') {
      initialData = { attendance_date: '', corrected_time: '', correction_type: 'IN' };
    }

    setFormData({ ...formData, type, submission_data: initialData });
  };

  const handleDataChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) => {
    const { name, value } = e.target;
    setFormData({
      ...formData,
      submission_data: { ...formData.submission_data, [name]: value }
    });
  };

  useEffect(() => {
    // Auto-calculate duration for Cuti
    if (formData.type === 'Cuti' && formData.submission_data.start_date && formData.submission_data.end_date) {
      const start = new Date(formData.submission_data.start_date);
      const end = new Date(formData.submission_data.end_date);
      const diff = Math.ceil((end.getTime() - start.getTime()) / (1000 * 3600 * 24)) + 1;
      if (diff > 0) {
        setFormData((prev: any) => ({
          ...prev,
          submission_data: { ...prev.submission_data, duration_days: diff }
        }));
      }
    }
  }, [formData.submission_data.start_date, formData.submission_data.end_date]);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      setUploading(true);
      const fileId = await googleDriveService.uploadFile(file);
      setFormData({ ...formData, file_id: fileId });
    } catch (error) {
      Swal.fire('Gagal', 'Gagal mengunggah lampiran.', 'error');
    } finally {
      setUploading(false);
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!currentUser) return;

    try {
      setIsSaving(true);
      await submissionService.create({
        account_id: currentUser.id,
        type: formData.type,
        description: formData.description,
        file_id: formData.file_id,
        submission_data: formData.submission_data
      });
      Swal.fire({ title: 'Berhasil!', text: 'Pengajuan Anda telah dikirim.', icon: 'success', timer: 1500, showConfirmButton: false });
      onSuccess();
    } catch (error) {
      Swal.fire('Gagal', 'Gagal mengirim pengajuan.', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-40 p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg overflow-hidden flex flex-col max-h-[90vh] animate-in zoom-in duration-200">
        <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between bg-[#006E62]/5">
          <div>
            <h3 className="text-base font-bold text-[#006E62]">Buat Pengajuan Baru</h3>
            <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Formulir Workflow Digital</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={20} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="flex-1 overflow-y-auto p-6 space-y-5">
          <div className="space-y-1">
            <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Jenis Pengajuan</label>
            <select 
              value={formData.type}
              onChange={handleTypeChange}
              className="w-full px-3 py-2 text-xs border border-gray-200 rounded-lg focus:ring-1 focus:ring-[#006E62] outline-none bg-gray-50 font-bold"
            >
              <option value="Cuti">Pengajuan Cuti</option>
              <option value="Lembur">Permohonan Lembur</option>
              <option value="Izin">Izin Tidak Masuk</option>
              <option value="Koreksi Absen">Koreksi Data Absensi</option>
            </select>
          </div>

          {/* Dynamic Inputs based on type */}
          {formData.type === 'Cuti' && (
            <div className="grid grid-cols-2 gap-4 p-4 bg-emerald-50/50 rounded-xl border border-emerald-100/50">
              <div className="space-y-1">
                <label className="text-[9px] font-bold text-[#006E62] uppercase">Mulai Tanggal</label>
                <input type="date" required name="start_date" value={formData.submission_data.start_date} onChange={handleDataChange} className="w-full px-2 py-1.5 text-xs border rounded-md" />
              </div>
              <div className="space-y-1">
                <label className="text-[9px] font-bold text-[#006E62] uppercase">Sampai Tanggal</label>
                <input type="date" required name="end_date" value={formData.submission_data.end_date} onChange={handleDataChange} className="w-full px-2 py-1.5 text-xs border rounded-md" />
              </div>
              <div className="col-span-2 text-center pt-2">
                 <p className="text-[10px] font-bold text-[#006E62] uppercase">Total: {formData.submission_data.duration_days} Hari Kerja</p>
              </div>
            </div>
          )}

          {formData.type === 'Lembur' && (
            <div className="grid grid-cols-3 gap-3 p-4 bg-orange-50/50 rounded-xl border border-orange-100/50">
              <div className="space-y-1">
                <label className="text-[9px] font-bold text-orange-600 uppercase">Tanggal</label>
                <input type="date" required name="date" value={formData.submission_data.date} onChange={handleDataChange} className="w-full px-2 py-1.5 text-xs border rounded-md" />
              </div>
              <div className="space-y-1">
                <label className="text-[9px] font-bold text-orange-600 uppercase">Jam Mulai</label>
                <input type="time" required name="start_time" value={formData.submission_data.start_time} onChange={handleDataChange} className="w-full px-2 py-1.5 text-xs border rounded-md" />
              </div>
              <div className="space-y-1">
                <label className="text-[9px] font-bold text-orange-600 uppercase">Jam Akhir</label>
                <input type="time" required name="end_time" value={formData.submission_data.end_time} onChange={handleDataChange} className="w-full px-2 py-1.5 text-xs border rounded-md" />
              </div>
            </div>
          )}

          <div className="space-y-1">
            <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Alasan / Keterangan</label>
            <textarea 
              required
              rows={3}
              name="description"
              value={formData.description}
              onChange={(e) => setFormData({...formData, description: e.target.value})}
              placeholder="Jelaskan secara detail alasan pengajuan Anda..."
              className="w-full px-3 py-2 text-xs border border-gray-200 rounded-lg focus:ring-1 focus:ring-[#006E62] outline-none resize-none"
            />
          </div>

          <div className="space-y-1">
            <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Lampiran Bukti (Opsional)</label>
            <div className={`p-4 border border-dashed rounded-xl transition-all ${formData.file_id ? 'border-[#006E62] bg-emerald-50/20' : 'border-gray-200 bg-gray-50'}`}>
              <label className="flex flex-col items-center justify-center cursor-pointer min-h-[60px]">
                {uploading ? <Loader2 className="animate-spin text-[#006E62]" /> : <Upload className={formData.file_id ? 'text-[#006E62]' : 'text-gray-300'} size={24} />}
                <span className="text-[10px] font-bold uppercase mt-2 text-gray-400">
                  {formData.file_id ? 'FILE TERUNGGAH' : 'UNGGAH PDF / GAMBAR'}
                </span>
                <input type="file" className="hidden" onChange={handleFileUpload} disabled={uploading} />
              </label>
            </div>
          </div>
        </form>

        <div className="px-6 py-4 border-t border-gray-100 flex justify-end gap-3 bg-gray-50">
          <button type="button" onClick={onClose} className="px-4 py-2 text-xs font-bold text-gray-500 uppercase">Batal</button>
          <button 
            onClick={handleSubmit}
            disabled={isSaving || uploading}
            className="flex items-center gap-2 bg-[#006E62] text-white px-8 py-2 rounded-lg shadow-md hover:bg-[#005a50] transition-all text-xs font-bold uppercase disabled:opacity-50"
          >
            {isSaving ? <Loader2 size={14} className="animate-spin" /> : <Save size={14} />}
            <span>Kirim Pengajuan</span>
          </button>
        </div>
      </div>
    </div>
  );
};

export default SubmissionForm;