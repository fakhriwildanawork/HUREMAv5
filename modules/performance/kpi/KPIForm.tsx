import React, { useState, useEffect } from 'react';
import { X, Save, Target, Calendar, User, Info, Plus, Trash2, Link as LinkIcon } from 'lucide-react';
import { KPIInput, Account } from '../../../types';
import { accountService } from '../../../services/accountService';

interface KPIFormProps {
  onClose: () => void;
  onSubmit: (data: KPIInput) => void;
  initialData?: any;
}

const KPIForm: React.FC<KPIFormProps> = ({ onClose, onSubmit, initialData }) => {
  const [formData, setFormData] = useState<KPIInput>({
    account_id: initialData?.account_id || '',
    title: initialData?.title || '',
    description: initialData?.description || '',
    weight: initialData?.weight || 1,
    start_date: initialData?.start_date || new Date().toLocaleDateString('en-CA'),
    deadline: initialData?.deadline || new Date().toLocaleDateString('en-CA'),
    supporting_links: initialData?.supporting_links || []
  });

  const [accounts, setAccounts] = useState<Account[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [newLink, setNewLink] = useState('');

  useEffect(() => {
    accountService.getAll().then(data => {
      setAccounts(data);
      setIsLoading(false);
    });
  }, []);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => {
    const { name, value } = e.target;
    setFormData(prev => ({ ...prev, [name]: name === 'weight' ? parseInt(value) : value }));
  };

  const addLink = () => {
    if (newLink && !formData.supporting_links.includes(newLink)) {
      setFormData(prev => ({ ...prev, supporting_links: [...prev.supporting_links, newLink] }));
      setNewLink('');
    }
  };

  const removeLink = (link: string) => {
    setFormData(prev => ({ ...prev, supporting_links: prev.supporting_links.filter(l => l !== link) }));
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (new Date(formData.start_date + 'T00:00:00') > new Date(formData.deadline + 'T00:00:00')) {
      alert('Tanggal mulai tidak boleh lebih besar dari deadline.');
      return;
    }
    onSubmit(formData);
  };

  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 p-4 backdrop-blur-sm">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg overflow-hidden animate-in zoom-in duration-200">
        <div className="px-6 py-4 border-b border-gray-50 flex items-center justify-between bg-gray-50/50">
          <div className="flex items-center gap-2">
            <Target size={18} className="text-[#006E62]" />
            <h3 className="text-sm font-bold text-gray-800 uppercase tracking-tight">
              {initialData ? 'Ubah KPI' : 'Buat KPI Baru'}
            </h3>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={20} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-6 space-y-5 max-h-[80vh] overflow-y-auto scrollbar-thin">
          <div className="space-y-1.5">
            <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Pilih Pegawai</label>
            <div className="relative">
              <select
                required
                name="account_id"
                value={formData.account_id}
                onChange={handleChange}
                className="w-full px-3 py-2.5 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] text-xs font-bold appearance-none"
              >
                <option value="">-- Pilih Pegawai --</option>
                {accounts.map(acc => (
                  <option key={acc.id} value={acc.id}>{acc.full_name} ({acc.internal_nik})</option>
                ))}
              </select>
              <User size={14} className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 pointer-events-none" />
            </div>
          </div>

          <div className="space-y-1.5">
            <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Judul KPI</label>
            <input 
              required
              type="text" 
              name="title" 
              value={formData.title} 
              onChange={handleChange}
              placeholder="Contoh: Pencapaian Target Penjualan Q1"
              className="w-full px-3 py-2.5 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] text-xs font-bold"
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Deskripsi & Target</label>
            <textarea 
              required
              name="description" 
              value={formData.description} 
              onChange={handleChange}
              rows={3}
              placeholder="Jelaskan detail KPI dan target yang harus dicapai..."
              className="w-full px-3 py-2.5 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] text-xs font-medium resize-none"
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Link Pendukung/Referensi (Multiple)</label>
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
            {formData.supporting_links.length > 0 && (
              <div className="space-y-1 mt-2">
                {formData.supporting_links.map(link => (
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

          <div className="grid grid-cols-2 gap-4">
            <div className="space-y-1.5">
              <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Tanggal Mulai</label>
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
              <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Deadline</label>
              <input 
                required
                type="date" 
                name="deadline" 
                value={formData.deadline} 
                onChange={handleChange}
                className="w-full px-3 py-2.5 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] text-xs font-bold"
              />
            </div>
          </div>

          <div className="space-y-1.5">
            <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Bobot KPI (1-10)</label>
            <input 
              required
              type="number" 
              name="weight" 
              min="1"
              max="10"
              value={formData.weight} 
              onChange={handleChange}
              className="w-full px-3 py-2.5 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] text-xs font-bold"
            />
          </div>

          <div className="p-3 bg-blue-50 border border-blue-100 rounded-xl flex gap-3 items-start">
            <Info size={16} className="text-blue-500 shrink-0 mt-0.5" />
            <p className="text-[9px] text-blue-700 leading-relaxed font-medium">
              KPI yang dibuat akan muncul di dashboard pegawai. Bobot digunakan untuk menghitung skor performa kumulatif pegawai secara otomatis.
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
              Simpan KPI
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};

export default KPIForm;
