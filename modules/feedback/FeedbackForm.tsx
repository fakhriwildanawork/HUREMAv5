import React, { useState, useRef } from 'react';
import { X, Send, Paperclip, Link as LinkIcon, Shield, ShieldOff, Trash2, Plus, Loader2, Upload } from 'lucide-react';
import { googleDriveService } from '../../services/googleDriveService';

interface FeedbackFormProps {
  onClose: () => void;
  onSubmit: (data: any) => void;
}

const FeedbackForm: React.FC<FeedbackFormProps> = ({ onClose, onSubmit }) => {
  const [category, setCategory] = useState('Gaji');
  const [priority, setPriority] = useState('Medium');
  const [isAnonymous, setIsAnonymous] = useState(false);
  const [description, setDescription] = useState('');
  const [links, setLinks] = useState<string[]>(['']);
  const [attachments, setAttachments] = useState<string[]>([]);
  const [isUploading, setIsUploading] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleAddLink = () => setLinks([...links, '']);
  const handleRemoveLink = (index: number) => setLinks(links.filter((_, i) => i !== index));
  const handleLinkChange = (index: number, value: string) => {
    const newLinks = [...links];
    newLinks[index] = value;
    setLinks(newLinks);
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    try {
      setIsUploading(true);
      for (let i = 0; i < files.length; i++) {
        const fileId = await googleDriveService.uploadFile(files[i]);
        setAttachments(prev => [...prev, fileId]);
      }
    } catch (error) {
      alert('Gagal mengunggah file.');
    } finally {
      setIsUploading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const handleRemoveAttachment = (index: number) => setAttachments(attachments.filter((_, i) => i !== index));

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!description.trim()) return;

    onSubmit({
      category,
      priority,
      is_anonymous: isAnonymous,
      description,
      links: links.filter(l => l.trim() !== ''),
      attachments: attachments.filter(a => a.trim() !== '')
    });
  };

  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
      <div className="bg-white w-full max-w-2xl rounded-3xl overflow-hidden shadow-2xl flex flex-col animate-in fade-in zoom-in duration-300 max-h-[90vh]">
        <div className="p-6 border-b border-gray-100 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-emerald-50 flex items-center justify-center text-[#006E62]">
              <Send size={20} />
            </div>
            <div>
              <h3 className="text-lg font-bold text-gray-800 tracking-tight">Kirim Feedback</h3>
              <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Sampaikan aspirasi Anda secara profesional</p>
            </div>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-gray-100 rounded-xl transition-colors text-gray-400">
            <X size={24} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-6 overflow-y-auto space-y-6 scrollbar-none">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="space-y-1.5">
              <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Kategori Feedback</label>
              <select 
                value={category}
                onChange={(e) => setCategory(e.target.value)}
                className="w-full px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] transition-all text-sm font-bold text-gray-700"
              >
                <option value="Gaji">Gaji & Kompensasi</option>
                <option value="Fasilitas">Fasilitas Kantor</option>
                <option value="Hubungan Kerja">Hubungan Kerja</option>
                <option value="Lainnya">Lainnya</option>
              </select>
            </div>

            <div className="space-y-1.5">
              <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Tingkat Urgensi</label>
              <select 
                value={priority}
                onChange={(e) => setPriority(e.target.value)}
                className="w-full px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] transition-all text-sm font-bold text-gray-700"
              >
                <option value="Low">Low - Saran/Masukan</option>
                <option value="Medium">Medium - Perlu Perhatian</option>
                <option value="High">High - Mendesak/Kritis</option>
              </select>
            </div>
          </div>

          <div className="p-4 bg-amber-50 rounded-2xl border border-amber-100 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className={`w-10 h-10 rounded-xl flex items-center justify-center transition-colors ${isAnonymous ? 'bg-gray-800 text-white' : 'bg-white text-gray-400 border border-gray-100'}`}>
                {isAnonymous ? <Shield size={20} /> : <ShieldOff size={20} />}
              </div>
              <div>
                <p className="text-xs font-bold text-gray-800">Kirim sebagai Anonim?</p>
                <p className="text-[10px] text-gray-500 font-medium">Identitas Anda tidak akan terlihat oleh Admin</p>
              </div>
            </div>
            <button 
              type="button"
              onClick={() => setIsAnonymous(!isAnonymous)}
              className={`w-12 h-6 rounded-full transition-all relative ${isAnonymous ? 'bg-[#006E62]' : 'bg-gray-200'}`}
            >
              <div className={`absolute top-1 w-4 h-4 rounded-full bg-white transition-all ${isAnonymous ? 'left-7' : 'left-1'}`} />
            </button>
          </div>

          <div className="space-y-1.5">
            <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Deskripsi Feedback</label>
            <textarea 
              required
              rows={4}
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              placeholder="Tuliskan detail masukan atau keluhan Anda di sini..."
              className="w-full px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] transition-all text-sm font-medium text-gray-700 resize-none"
            />
          </div>

          <div className="space-y-4">
            <div className="flex items-center justify-between px-1">
              <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Lampiran File</label>
              <button 
                type="button" 
                onClick={() => fileInputRef.current?.click()} 
                disabled={isUploading}
                className="text-[10px] font-bold text-[#006E62] uppercase flex items-center gap-1 hover:underline disabled:opacity-50"
              >
                <Plus size={12} /> Tambah File
              </button>
              <input 
                ref={fileInputRef}
                type="file"
                multiple
                onChange={handleFileUpload}
                className="hidden"
              />
            </div>

            {isUploading && (
              <div className="flex items-center gap-2 px-3 py-2 bg-emerald-50 text-[#006E62] rounded-xl animate-pulse">
                <Loader2 size={14} className="animate-spin" />
                <span className="text-[10px] font-bold uppercase tracking-widest">Mengunggah File...</span>
              </div>
            )}

            <div className="space-y-2">
              {attachments.map((fileId, idx) => (
                <div key={idx} className="flex items-center justify-between px-3 py-2 bg-gray-50 border border-gray-100 rounded-xl group">
                  <div className="flex items-center gap-2 overflow-hidden">
                    <Paperclip size={12} className="text-gray-400 shrink-0" />
                    <span className="text-[10px] text-gray-600 truncate">File ID: {fileId}</span>
                  </div>
                  <button 
                    type="button" 
                    onClick={() => handleRemoveAttachment(idx)} 
                    className="p-1.5 text-gray-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-all"
                  >
                    <Trash2 size={14} />
                  </button>
                </div>
              ))}
            </div>
          </div>

          <div className="space-y-4">
            <div className="flex items-center justify-between px-1">
              <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Link Referensi (URL)</label>
              <button type="button" onClick={handleAddLink} className="text-[10px] font-bold text-[#006E62] uppercase flex items-center gap-1 hover:underline">
                <Plus size={12} /> Tambah Link
              </button>
            </div>
            <div className="space-y-2">
              {links.map((link, idx) => (
                <div key={idx} className="flex gap-2">
                  <div className="relative flex-1">
                    <LinkIcon className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-300" size={14} />
                    <input 
                      type="url"
                      value={link}
                      onChange={(e) => handleLinkChange(idx, e.target.value)}
                      placeholder="https://example.com/..."
                      className="w-full pl-9 pr-4 py-2 bg-gray-50 border border-gray-100 rounded-lg focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 text-xs font-medium"
                    />
                  </div>
                  {links.length > 1 && (
                    <button type="button" onClick={() => handleRemoveLink(idx)} className="p-2 text-rose-500 hover:bg-rose-50 rounded-lg transition-colors">
                      <Trash2 size={16} />
                    </button>
                  )}
                </div>
              ))}
            </div>
          </div>
        </form>

        <div className="p-6 border-t border-gray-100 bg-gray-50/50 flex gap-3 shrink-0">
          <button 
            type="button"
            onClick={onClose}
            className="flex-1 py-3 border border-gray-200 text-gray-600 rounded-xl font-bold text-xs uppercase tracking-widest hover:bg-white transition-all"
          >
            Batal
          </button>
          <button 
            onClick={handleSubmit}
            disabled={isUploading}
            className="flex-2 py-3 bg-[#006E62] text-white rounded-xl font-bold text-xs uppercase tracking-widest hover:bg-[#005a50] transition-all shadow-lg shadow-[#006E62]/20 flex items-center justify-center gap-2 disabled:opacity-50"
          >
            <Send size={16} />
            Kirim Sekarang
          </button>
        </div>
      </div>
    </div>
  );
};

export default FeedbackForm;
