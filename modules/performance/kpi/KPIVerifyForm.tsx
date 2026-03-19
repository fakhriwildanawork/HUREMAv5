import React, { useState } from 'react';
import { X, Save, Target, Info, FileText, Link as LinkIcon, CheckCircle2 } from 'lucide-react';
import { KPI } from '../../../types';
import { googleDriveService } from '../../../services/googleDriveService';

interface KPIVerifyFormProps {
  onClose: () => void;
  onSubmit: (data: { score: number, notes: string }) => void;
  kpi: KPI;
}

const KPIVerifyForm: React.FC<KPIVerifyFormProps> = ({ onClose, onSubmit, kpi }) => {
  const [formData, setFormData] = useState({
    score: kpi.report_data?.self_assessment || 100,
    notes: ''
  });

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    onSubmit(formData);
  };

  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 p-4 backdrop-blur-sm">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg overflow-hidden animate-in zoom-in duration-200">
        <div className="px-6 py-4 border-b border-gray-50 flex items-center justify-between bg-gray-50/50">
          <div className="flex items-center gap-2">
            <CheckCircle2 size={18} className="text-[#006E62]" />
            <h3 className="text-sm font-bold text-gray-800 uppercase tracking-tight">Verifikasi Capaian KPI</h3>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={20} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-6 space-y-5 max-h-[80vh] overflow-y-auto scrollbar-thin">
          <div className="p-4 bg-gray-50 border border-gray-100 rounded-2xl space-y-3">
            <div>
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">KPI & Pegawai</p>
              <p className="text-xs font-bold text-gray-800">{kpi.title}</p>
              <p className="text-[10px] text-[#006E62] font-bold uppercase">{kpi.account?.full_name}</p>
            </div>
            <div className="pt-2 border-t border-gray-100">
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Laporan Pegawai</p>
              <p className="text-xs text-gray-600 italic leading-relaxed">"{kpi.report_data?.description}"</p>
            </div>
          </div>

          <div className="grid grid-cols-2 gap-4">
            <div className="space-y-1.5">
              <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Lampiran File</label>
              <div className="space-y-1">
                {kpi.report_data?.file_ids && kpi.report_data.file_ids.length > 0 ? (
                  kpi.report_data.file_ids.map((id, idx) => (
                    <a 
                      key={id} 
                      href={googleDriveService.getFileUrl(id)} 
                      target="_blank" 
                      rel="noopener noreferrer"
                      className="flex items-center gap-2 px-3 py-1.5 bg-white border border-gray-100 rounded-lg hover:bg-emerald-50 transition-colors group"
                    >
                      <FileText size={12} className="text-gray-400 group-hover:text-[#006E62]" />
                      <span className="text-[10px] font-bold text-gray-500 uppercase">File {idx + 1}</span>
                    </a>
                  ))
                ) : (
                  <p className="text-[10px] text-gray-400 italic px-1">Tidak ada file</p>
                )}
              </div>
            </div>
            <div className="space-y-1.5">
              <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Lampiran Link</label>
              <div className="space-y-1">
                {kpi.report_data?.links && kpi.report_data.links.length > 0 ? (
                  kpi.report_data.links.map((link, idx) => (
                    <a 
                      key={idx} 
                      href={link} 
                      target="_blank" 
                      rel="noopener noreferrer"
                      className="flex items-center gap-2 px-3 py-1.5 bg-white border border-gray-100 rounded-lg hover:bg-emerald-50 transition-colors group"
                    >
                      <LinkIcon size={12} className="text-gray-400 group-hover:text-[#006E62]" />
                      <span className="text-[10px] font-bold text-gray-500 uppercase truncate">Link {idx + 1}</span>
                    </a>
                  ))
                ) : (
                  <p className="text-[10px] text-gray-400 italic px-1">Tidak ada link</p>
                )}
              </div>
            </div>
          </div>

          <div className="space-y-1.5">
            <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Skor Verifikasi (0-100)</label>
            <div className="flex items-center gap-4">
              <input 
                type="range" 
                min="0" 
                max="100" 
                step="5"
                value={formData.score}
                onChange={(e) => setFormData(prev => ({ ...prev, score: parseInt(e.target.value) }))}
                className="flex-1 accent-[#006E62]"
              />
              <span className="text-sm font-bold text-[#006E62] w-12 text-right">{formData.score}%</span>
            </div>
            <p className="text-[9px] text-gray-400 italic px-1">Klaim Pegawai: {kpi.report_data?.self_assessment}%</p>
          </div>

          <div className="space-y-1.5">
            <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Catatan Verifikasi</label>
            <textarea 
              required
              name="notes" 
              value={formData.notes} 
              onChange={(e) => setFormData(prev => ({ ...prev, notes: e.target.value }))}
              rows={3}
              placeholder="Berikan feedback atau alasan pemberian skor..."
              className="w-full px-3 py-2.5 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] text-xs font-medium resize-none"
            />
          </div>

          <div className="p-3 bg-blue-50 border border-blue-100 rounded-xl flex gap-3 items-start">
            <Info size={16} className="text-blue-500 shrink-0 mt-0.5" />
            <p className="text-[9px] text-blue-700 leading-relaxed font-medium">
              Skor verifikasi ini akan menjadi nilai final KPI dan akan mempengaruhi skor performa kumulatif pegawai.
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
              className="flex-2 py-3 bg-[#006E62] text-white rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-[#005a50] shadow-lg shadow-[#006E62]/20 transition-all flex items-center justify-center gap-2"
            >
              <Save size={16} />
              Simpan Verifikasi
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};

export default KPIVerifyForm;
