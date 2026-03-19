import React, { useState } from 'react';
import { X, Send, FileText, Link as LinkIcon, Plus, Trash2, Calendar, Target, Info } from 'lucide-react';
import { KeyActivityReportInput } from '../../../types';

interface KeyActivityReportFormProps {
  activityTitle: string;
  dueDate: string;
  onClose: () => void;
  onSubmit: (data: KeyActivityReportInput) => void;
}

const KeyActivityReportForm: React.FC<KeyActivityReportFormProps> = ({ activityTitle, dueDate, onClose, onSubmit }) => {
  const [formData, setFormData] = useState<KeyActivityReportInput>({
    activity_id: '', // Will be set by parent
    account_id: '', // Will be set by parent
    due_date: dueDate,
    description: '',
    file_ids: [],
    links: []
  });

  const [newLink, setNewLink] = useState('');

  const addLink = () => {
    if (newLink && !formData.links.includes(newLink)) {
      setFormData(prev => ({ ...prev, links: [...prev.links, newLink] }));
      setNewLink('');
    }
  };

  const removeLink = (link: string) => {
    setFormData(prev => ({ ...prev, links: prev.links.filter(l => l !== link) }));
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
            <FileText size={18} className="text-[#006E62]" />
            <h3 className="text-sm font-bold text-gray-800 uppercase tracking-tight">Laporan Key Activity</h3>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={20} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-6 space-y-5">
          <div className="p-4 bg-gray-50 rounded-2xl border border-gray-100 space-y-2">
            <div className="flex items-center gap-2">
              <Target size={14} className="text-[#006E62]" />
              <p className="text-xs font-bold text-gray-800">{activityTitle}</p>
            </div>
            <div className="flex items-center gap-2">
              <Calendar size={14} className="text-rose-500" />
              <p className="text-[10px] font-bold text-rose-500 uppercase tracking-widest">
                Untuk Tanggal: {new Date(dueDate + 'T00:00:00').toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' })}
              </p>
            </div>
          </div>

          <div className="space-y-1.5">
            <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Deskripsi Capaian & Hasil</label>
            <textarea 
              required
              value={formData.description} 
              onChange={(e) => setFormData(prev => ({ ...prev, description: e.target.value }))}
              rows={4}
              placeholder="Jelaskan apa yang telah Anda lakukan untuk aktivitas ini..."
              className="w-full px-3 py-2.5 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] text-xs font-medium resize-none"
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Lampiran Link Bukti (Opsional)</label>
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

          <div className="p-3 bg-amber-50 border border-amber-100 rounded-xl flex gap-3 items-start">
            <Info size={16} className="text-amber-500 shrink-0 mt-0.5" />
            <p className="text-[9px] text-amber-700 leading-relaxed font-medium">
              Pastikan laporan Anda sudah lengkap. Laporan yang sudah dikirim akan menunggu verifikasi dari Admin untuk mendapatkan skor performa.
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
              <Send size={16} />
              Kirim Laporan
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};

export default KeyActivityReportForm;
