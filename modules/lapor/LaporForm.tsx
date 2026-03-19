import React, { useState, useRef } from 'react';
import { X, Send, Paperclip, Link as LinkIcon, AlertTriangle, Trash2, Plus, Loader2, Users, Search } from 'lucide-react';
import { googleDriveService } from '../../services/googleDriveService';
import { Account } from '../../types';

interface LaporFormProps {
  onClose: () => void;
  onSubmit: (data: any) => void;
  accounts: Account[];
}

const LaporForm: React.FC<LaporFormProps> = ({ onClose, onSubmit, accounts }) => {
  const [category, setCategory] = useState('Pencurian');
  const [description, setDescription] = useState('');
  const [reportedAccountIds, setReportedAccountIds] = useState<string[]>([]);
  const [links, setLinks] = useState<string[]>(['']);
  const [attachments, setAttachments] = useState<string[]>([]);
  const [isUploading, setIsUploading] = useState(false);
  const [accountSearch, setAccountSearch] = useState('');
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

  const toggleAccountSelection = (accountId: string) => {
    setReportedAccountIds(prev => 
      prev.includes(accountId) 
        ? prev.filter(id => id !== accountId) 
        : [...prev, accountId]
    );
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!description.trim()) return;
    if (reportedAccountIds.length === 0) {
      alert('Pilih setidaknya satu pegawai yang dilaporkan.');
      return;
    }

    onSubmit({
      category,
      description,
      reported_account_ids: reportedAccountIds,
      links: links.filter(l => l.trim() !== ''),
      attachments: attachments.filter(a => a.trim() !== '')
    });
  };

  const filteredAccounts = accounts.filter(a => 
    a.full_name.toLowerCase().includes(accountSearch.toLowerCase()) ||
    a.internal_nik.toLowerCase().includes(accountSearch.toLowerCase())
  );

  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
      <div className="bg-white w-full max-w-3xl rounded-3xl overflow-hidden shadow-2xl flex flex-col animate-in fade-in zoom-in duration-300 max-h-[95vh]">
        <div className="p-6 border-b border-gray-100 flex items-center justify-between shrink-0 bg-rose-50/30">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-rose-100 flex items-center justify-center text-rose-600">
              <AlertTriangle size={20} />
            </div>
            <div>
              <h3 className="text-lg font-bold text-gray-800 tracking-tight">Buat Laporan Pelanggaran</h3>
              <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Sampaikan laporan Anda dengan jujur dan bertanggung jawab</p>
            </div>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-gray-100 rounded-xl transition-colors text-gray-400">
            <X size={24} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-6 overflow-y-auto space-y-6 scrollbar-none">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            <div className="md:col-span-2 space-y-6">
              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Kategori Pelanggaran</label>
                <select 
                  value={category}
                  onChange={(e) => setCategory(e.target.value)}
                  className="w-full px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-rose-600/20 focus:border-rose-600 transition-all text-sm font-bold text-gray-700"
                >
                  <option value="Pencurian">Pencurian / Fraud</option>
                  <option value="Perusakan">Perusakan Fasilitas</option>
                  <option value="Bullying">Bullying / Pelecehan</option>
                  <option value="Pelanggaran SOP">Pelanggaran SOP Berat</option>
                  <option value="Lainnya">Lainnya</option>
                </select>
              </div>

              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Deskripsi Kejadian</label>
                <textarea 
                  required
                  rows={6}
                  value={description}
                  onChange={(e) => setDescription(e.target.value)}
                  placeholder="Ceritakan kronologi kejadian secara detail (Waktu, Tempat, dan Kejadian)..."
                  className="w-full px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-rose-600/20 focus:border-rose-600 transition-all text-sm font-medium text-gray-700 resize-none"
                />
              </div>

              <div className="space-y-4">
                <div className="flex items-center justify-between px-1">
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Lampiran Bukti (File)</label>
                  <button 
                    type="button" 
                    onClick={() => fileInputRef.current?.click()} 
                    disabled={isUploading}
                    className="text-[10px] font-bold text-rose-600 uppercase flex items-center gap-1 hover:underline disabled:opacity-50"
                  >
                    <Plus size={12} /> Tambah Bukti
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
                  <div className="flex items-center gap-2 px-3 py-2 bg-rose-50 text-rose-600 rounded-xl animate-pulse">
                    <Loader2 size={14} className="animate-spin" />
                    <span className="text-[10px] font-bold uppercase tracking-widest">Mengunggah Bukti...</span>
                  </div>
                )}

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  {attachments.map((fileId, idx) => (
                    <div key={idx} className="flex items-center justify-between px-3 py-2 bg-gray-50 border border-gray-100 rounded-xl group">
                      <div className="flex items-center gap-2 overflow-hidden">
                        <Paperclip size={12} className="text-gray-400 shrink-0" />
                        <span className="text-[10px] text-gray-600 truncate">Bukti_{idx + 1}</span>
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
            </div>

            <div className="space-y-4">
              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Pilih Terlapor ({reportedAccountIds.length})</label>
                <div className="relative mb-2">
                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={14} />
                  <input 
                    type="text"
                    value={accountSearch}
                    onChange={(e) => setAccountSearch(e.target.value)}
                    placeholder="Cari nama/NIK..."
                    className="w-full pl-9 pr-4 py-2 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-rose-600/20 text-[11px] font-bold"
                  />
                </div>
                <div className="bg-gray-50 border border-gray-100 rounded-2xl overflow-hidden max-h-[400px] overflow-y-auto scrollbar-thin">
                  {filteredAccounts.map(account => (
                    <button
                      key={account.id}
                      type="button"
                      onClick={() => toggleAccountSelection(account.id)}
                      className={`w-full flex items-center gap-3 p-3 transition-all border-b border-gray-50 last:border-0 ${
                        reportedAccountIds.includes(account.id) 
                          ? 'bg-rose-50 text-rose-600' 
                          : 'hover:bg-white text-gray-600'
                      }`}
                    >
                      <div className={`w-8 h-8 rounded-full flex items-center justify-center text-[10px] font-bold shrink-0 ${
                        reportedAccountIds.includes(account.id) ? 'bg-rose-600 text-white' : 'bg-gray-200 text-gray-500'
                      }`}>
                        {account.full_name.charAt(0)}
                      </div>
                      <div className="text-left overflow-hidden">
                        <p className="text-[11px] font-bold truncate">{account.full_name}</p>
                        <p className="text-[9px] opacity-60 font-medium">{account.internal_nik}</p>
                      </div>
                    </button>
                  ))}
                  {filteredAccounts.length === 0 && (
                    <div className="p-8 text-center text-gray-400">
                      <p className="text-[10px] font-bold uppercase">Tidak ditemukan</p>
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>

          <div className="space-y-4">
            <div className="flex items-center justify-between px-1">
              <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Link Referensi (URL)</label>
              <button type="button" onClick={handleAddLink} className="text-[10px] font-bold text-rose-600 uppercase flex items-center gap-1 hover:underline">
                <Plus size={12} /> Tambah Link
              </button>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
              {links.map((link, idx) => (
                <div key={idx} className="flex gap-2">
                  <div className="relative flex-1">
                    <LinkIcon className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-300" size={14} />
                    <input 
                      type="url"
                      value={link}
                      onChange={(e) => handleLinkChange(idx, e.target.value)}
                      placeholder="https://example.com/..."
                      className="w-full pl-9 pr-4 py-2 bg-gray-50 border border-gray-100 rounded-lg focus:outline-none focus:ring-2 focus:ring-rose-600/20 text-xs font-medium"
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
            className="flex-[2] py-3 bg-rose-600 text-white rounded-xl font-bold text-xs uppercase tracking-widest hover:bg-rose-700 transition-all shadow-lg shadow-rose-600/20 flex items-center justify-center gap-2 disabled:opacity-50"
          >
            <Send size={16} />
            Kirim Laporan
          </button>
        </div>
      </div>
    </div>
  );
};

export default LaporForm;
