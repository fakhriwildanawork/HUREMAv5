import React, { useState, useEffect } from 'react';
import { X, Save, Target, Calendar, User, Info, Plus, Trash2, Link as LinkIcon, Clock, CheckSquare } from 'lucide-react';
import { KeyActivityInput, Account } from '../../../types';
import { accountService } from '../../../services/accountService';

interface KeyActivityFormProps {
  onClose: () => void;
  onSubmit: (data: KeyActivityInput) => void;
  initialData?: any;
}

const KeyActivityForm: React.FC<KeyActivityFormProps> = ({ onClose, onSubmit, initialData }) => {
  const [formData, setFormData] = useState<KeyActivityInput>({
    title: initialData?.title || '',
    description: initialData?.description || '',
    weight: initialData?.weight || 1,
    start_date: initialData?.start_date || new Date().toLocaleDateString('en-CA'),
    end_date: initialData?.end_date || new Date().toLocaleDateString('en-CA'),
    recurrence_type: initialData?.recurrence_type || 'Daily',
    recurrence_rule: initialData?.recurrence_rule || { days_of_week: [], dates_of_month: [] },
    supporting_links: initialData?.supporting_links || [],
    assigned_account_ids: initialData?.assignments?.map((a: any) => a.account_id) || []
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

  const handleRecurrenceChange = (type: string) => {
    setFormData(prev => ({ 
      ...prev, 
      recurrence_type: type as any,
      recurrence_rule: { days_of_week: [], dates_of_month: [] }
    }));
  };

  const toggleDay = (day: number) => {
    setFormData(prev => {
      const days = prev.recurrence_rule?.days_of_week || [];
      const newDays = days.includes(day) ? days.filter(d => d !== day) : [...days, day];
      return { ...prev, recurrence_rule: { ...prev.recurrence_rule, days_of_week: newDays } };
    });
  };

  const toggleDate = (date: number) => {
    setFormData(prev => {
      const dates = prev.recurrence_rule?.dates_of_month || [];
      const newDates = dates.includes(date) ? dates.filter(d => d !== date) : [...dates, date];
      return { ...prev, recurrence_rule: { ...prev.recurrence_rule, dates_of_month: newDates } };
    });
  };

  const toggleAccount = (accountId: string) => {
    setFormData(prev => {
      const ids = prev.assigned_account_ids;
      const newIds = ids.includes(accountId) ? ids.filter(id => id !== accountId) : [...ids, accountId];
      return { ...prev, assigned_account_ids: newIds };
    });
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
    if (formData.assigned_account_ids.length === 0) {
      alert('Pilih setidaknya satu pegawai.');
      return;
    }
    if (new Date(formData.start_date) > new Date(formData.end_date)) {
      alert('Tanggal mulai tidak boleh lebih besar dari tanggal berakhir.');
      return;
    }
    onSubmit(formData);
  };

  const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];

  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 p-4 backdrop-blur-sm">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in duration-200 flex flex-col max-h-[90vh]">
        <div className="px-6 py-4 border-b border-gray-50 flex items-center justify-between bg-gray-50/50 shrink-0">
          <div className="flex items-center gap-2">
            <CheckSquare size={18} className="text-[#006E62]" />
            <h3 className="text-sm font-bold text-gray-800 uppercase tracking-tight">
              {initialData ? 'Ubah Key Activity' : 'Buat Key Activity Baru'}
            </h3>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={20} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-6 space-y-6 overflow-y-auto scrollbar-thin">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="space-y-5">
              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Judul Aktivitas</label>
                <input 
                  required
                  type="text" 
                  name="title" 
                  value={formData.title} 
                  onChange={handleChange}
                  placeholder="Contoh: Laporan Penjualan Harian"
                  className="w-full px-3 py-2.5 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] text-xs font-bold"
                />
              </div>

              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Deskripsi & Instruksi</label>
                <textarea 
                  required
                  name="description" 
                  value={formData.description} 
                  onChange={handleChange}
                  rows={3}
                  placeholder="Jelaskan detail aktivitas yang harus dilakukan..."
                  className="w-full px-3 py-2.5 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] text-xs font-medium resize-none"
                />
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
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Tanggal Berakhir</label>
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
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Bobot Aktivitas (1-10)</label>
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
            </div>

            <div className="space-y-5">
              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Tipe Repetisi</label>
                <select
                  value={formData.recurrence_type}
                  onChange={(e) => handleRecurrenceChange(e.target.value)}
                  className="w-full px-3 py-2.5 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] text-xs font-bold"
                >
                  <option value="Once">Sekali Saja</option>
                  <option value="Daily">Setiap Hari</option>
                  <option value="Weekly">Mingguan (Pilih Hari)</option>
                  <option value="Monthly">Bulanan (Pilih Tanggal)</option>
                  <option value="EndOfMonth">Setiap Akhir Bulan</option>
                </select>
              </div>

              {formData.recurrence_type === 'Weekly' && (
                <div className="space-y-2">
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Pilih Hari</label>
                  <div className="flex flex-wrap gap-2">
                    {days.map((day, idx) => (
                      <button
                        key={day}
                        type="button"
                        onClick={() => toggleDay(idx)}
                        className={`px-3 py-1.5 rounded-lg text-[10px] font-bold uppercase transition-all border ${
                          formData.recurrence_rule?.days_of_week?.includes(idx)
                            ? 'bg-[#006E62] text-white border-[#006E62]'
                            : 'bg-gray-50 text-gray-400 border-gray-100 hover:bg-gray-100'
                        }`}
                      >
                        {day}
                      </button>
                    ))}
                  </div>
                </div>
              )}

              {formData.recurrence_type === 'Monthly' && (
                <div className="space-y-2">
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Pilih Tanggal</label>
                  <div className="grid grid-cols-7 gap-1">
                    {Array.from({ length: 31 }, (_, i) => i + 1).map(date => (
                      <button
                        key={date}
                        type="button"
                        onClick={() => toggleDate(date)}
                        className={`w-8 h-8 rounded-lg text-[10px] font-bold transition-all border flex items-center justify-center ${
                          formData.recurrence_rule?.dates_of_month?.includes(date)
                            ? 'bg-[#006E62] text-white border-[#006E62]'
                            : 'bg-gray-50 text-gray-400 border-gray-100 hover:bg-gray-100'
                        }`}
                      >
                        {date}
                      </button>
                    ))}
                  </div>
                </div>
              )}

              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Pilih Pegawai (Multi-Assignment)</label>
                <div className="max-h-40 overflow-y-auto border border-gray-100 rounded-xl p-2 space-y-1 scrollbar-thin">
                  {accounts.map(acc => (
                    <label key={acc.id} className="flex items-center gap-2 p-2 hover:bg-gray-50 rounded-lg cursor-pointer transition-colors">
                      <input 
                        type="checkbox" 
                        checked={formData.assigned_account_ids.includes(acc.id)}
                        onChange={() => toggleAccount(acc.id)}
                        className="rounded text-[#006E62] focus:ring-[#006E62]"
                      />
                      <div className="flex flex-col">
                        <span className="text-[10px] font-bold text-gray-700">{acc.full_name}</span>
                        <span className="text-[8px] text-gray-400 font-medium">{acc.internal_nik}</span>
                      </div>
                    </label>
                  ))}
                </div>
              </div>
            </div>
          </div>

          <div className="space-y-1.5">
            <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Link Pendukung/Referensi</label>
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
              <div className="flex flex-wrap gap-2 mt-2">
                {formData.supporting_links.map(link => (
                  <div key={link} className="flex items-center gap-2 px-3 py-1.5 bg-gray-50 border border-gray-100 rounded-lg">
                    <LinkIcon size={12} className="text-gray-400 shrink-0" />
                    <span className="text-[10px] text-gray-600 truncate max-w-[150px]">{link}</span>
                    <button type="button" onClick={() => removeLink(link)} className="text-rose-500 hover:text-rose-700">
                      <Trash2 size={12} />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>

          <div className="p-3 bg-blue-50 border border-blue-100 rounded-xl flex gap-3 items-start">
            <Info size={16} className="text-blue-500 shrink-0 mt-0.5" />
            <p className="text-[9px] text-blue-700 leading-relaxed font-medium">
              Key Activity yang bersifat repetisi akan muncul secara otomatis di dashboard pegawai sesuai jadwal. Jika pegawai melewatkan laporan, aktivitas tersebut akan masuk ke daftar <strong>Backlog</strong>.
            </p>
          </div>
        </form>

        <div className="px-6 py-4 border-t border-gray-50 bg-gray-50/30 flex gap-3 shrink-0">
          <button 
            type="button" 
            onClick={onClose}
            className="flex-1 py-3 border border-gray-100 text-gray-400 rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-gray-50 transition-all"
          >
            Batal
          </button>
          <button 
            onClick={handleSubmit}
            className="flex-2 py-3 bg-[#006E62] text-white rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-[#005a50] shadow-lg shadow-[#006E62]/20 transition-all flex items-center justify-center gap-2"
          >
            <Save size={16} />
            Simpan Aktivitas
          </button>
        </div>
      </div>
    </div>
  );
};

export default KeyActivityForm;
