import React, { useState } from 'react';
import { X, Save, Upload, Link as LinkIcon, Info, Target, Plus, Trash2 } from 'lucide-react';
import { googleDriveService } from '../../../services/googleDriveService';

interface KPIReportFormProps {
  onClose: () => void;
  onSubmit: (data: { description: string, file_ids: string[], links: string[], self_assessment: number }) => void;
  kpiTitle: string;
}

const KPIReportForm: React.FC<KPIReportFormProps> = ({ onClose, onSubmit, kpiTitle }) => {
  const [formData, setFormData] = useState({
    description: '',
    file_ids: [] as string[],
    links: [] as string[],
    self_assessment: 100
  });

  const [uploading, setUploading] = useState(false);
  const [newLink, setNewLink] = useState('');

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    try {
      setUploading(true);
      const uploadedIds = await Promise.all(
        Array.from(files).map(file => googleDriveService.uploadFile(file as File))
      );
      setFormData(prev => ({ ...prev, file_ids: [...prev.file_ids, ...uploadedIds] }));
    } catch (error) {
      alert('Gagal mengunggah beberapa file.');
    } finally {
      setUploading(false);
    }
  };

  const addLink = () => {
    if (newLink && !formData.links.includes(newLink)) {
      setFormData(prev => ({ ...prev, links: [...prev.links, newLink] }));
      setNewLink('');
    }
  };

  const removeLink = (link: string) => {
    setFormData(prev => ({ ...prev, links: prev.links.filter(l => l !== link) }));
  };

  const removeFile = (id: string) => {
    setFormData(prev => ({ ...prev, file_ids: prev.file_ids.filter(f => f !== id) }));
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    onSubmit(formData);
  };

  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 p-4 backdrop-blur-sm">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg overflow-hidden animate-in zoom-in duration-200">
        <div className="px-6 py-4 border-b border-gray-50 flex items-center justify-between bg-gray-50/50">
          <div className="flex items-center gap-2">
            <Target size={18} className="text-[#006E62]" />
            <h3 className="text-sm font-bold text-gray-800 uppercase tracking-tight">Laporan Capaian KPI</h3>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={20} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-6 space-y-5 max-h-[80vh] overflow-y-auto scrollbar-thin">
          <div className="p-3 bg-emerald-50 border border-emerald-100 rounded-xl">
            <p className="text-[9px] font-bold text-emerald-600 uppercase tracking-widest mb-1">KPI yang Dilaporkan</p>
            <p className="text-xs font-bold text-gray-800">{kpiTitle}</p>
          </div>

          <div className="space-y-1.5">
            <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Deskripsi Hasil Kerja</label>
            <textarea 
              required
              name="description" 
              value={formData.description} 
              onChange={(e) => setFormData(prev => ({ ...prev, description: e.target.value }))}
              rows={4}
              placeholder="Jelaskan apa saja yang telah Anda capai untuk KPI ini..."
              className="w-full px-3 py-2.5 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] text-xs font-medium resize-none"
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Lampiran File (Multiple)</label>
            <div className="space-y-2">
              <label className="flex items-center gap-2 px-4 py-2.5 bg-gray-50 border border-dashed border-gray-200 rounded-xl cursor-pointer hover:bg-gray-100 transition-all group overflow-hidden">
                <Upload size={16} className="text-gray-400 group-hover:text-[#006E62]" />
                <span className="text-[10px] font-bold text-gray-500 uppercase">Pilih File (Gambar/PDF)</span>
                <input type="file" multiple className="hidden" accept="image/*,application/pdf" onChange={handleFileUpload} />
              </label>
              {formData.file_ids.length > 0 && (
                <div className="flex flex-wrap gap-2">
                  {formData.file_ids.map(id => (
                    <div key={id} className="flex items-center gap-2 px-2 py-1 bg-emerald-50 text-emerald-600 rounded-lg text-[10px] font-bold">
                      <span>File Terunggah</span>
                      <button type="button" onClick={() => removeFile(id)} className="text-rose-500 hover:text-rose-700">
                        <Trash2 size={12} />
                      </button>
                    </div>
                  ))}
                </div>
              )}
              {uploading && <div className="w-4 h-4 border-2 border-[#006E62] border-t-transparent rounded-full animate-spin"></div>}
            </div>
          </div>

          <div className="space-y-1.5">
            <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Lampiran Link (Multiple)</label>
            <div className="flex gap-2">
              <input 
                type="url" 
                value={newLink}
                onChange={(e) => setNewLink(e.target.value)}
                placeholder="https://..."
                className="flex-1 px-3 py-2 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] text-xs"
              />
              <button 
                type="button" 
                onClick={addLink}
                className="px-3 py-2 bg-[#006E62] text-white rounded-xl hover:bg-[#005a50] transition-all"
              >
                <Plus size={16} />
              </button>
            </div>
            {formData.links.length > 0 && (
              <div className="space-y-1 mt-2">
                {formData.links.map(link => (
                  <div key={link} className="flex items-center justify-between px-3 py-1.5 bg-gray-50 border border-gray-100 rounded-lg">
                    <div className="flex items-center gap-2 overflow-hidden">
                      <LinkIcon size={12} className="text-gray-400 shrink-0" />
                      <span className="text-[10px] text-gray-600 truncate">{link}</span>
                    </div>
                    <button type="button" onClick={() => removeLink(link)} className="text-rose-500 hover:text-rose-700">
                      <Trash2 size={12} />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>

          <div className="space-y-1.5">
            <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Klaim Capaian Mandiri (0-100%)</label>
            <div className="flex items-center gap-4">
              <input 
                type="range" 
                min="0" 
                max="100" 
                step="5"
                value={formData.self_assessment}
                onChange={(e) => setFormData(prev => ({ ...prev, self_assessment: parseInt(e.target.value) }))}
                className="flex-1 accent-[#006E62]"
              />
              <span className="text-sm font-bold text-[#006E62] w-12 text-right">{formData.self_assessment}%</span>
            </div>
          </div>

          <div className="p-3 bg-amber-50 border border-amber-100 rounded-xl flex gap-3 items-start">
            <Info size={16} className="text-amber-500 shrink-0 mt-0.5" />
            <p className="text-[9px] text-amber-700 leading-relaxed font-medium">
              Laporan Anda akan diverifikasi oleh Admin. Pastikan bukti file dan link sudah lengkap untuk mendukung klaim capaian Anda.
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
              disabled={uploading}
              className="flex-2 py-3 bg-[#006E62] text-white rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-[#005a50] shadow-lg shadow-[#006E62]/20 transition-all flex items-center justify-center gap-2 disabled:opacity-50"
            >
              <Save size={16} />
              Kirim Laporan
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};

export default KPIReportForm;
