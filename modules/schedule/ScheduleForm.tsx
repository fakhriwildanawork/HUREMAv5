
import React, { useState, useEffect, useRef } from 'react';
import { X, Save, Clock, Calendar, Users, MapPin, Check, ChevronDown, Search } from 'lucide-react';
import { ScheduleInput, Location, Account } from '../../types';
import { locationService } from '../../services/locationService';
import { accountService } from '../../services/accountService';

interface ScheduleFormProps {
  onClose: () => void;
  onSubmit: (data: any) => void;
  initialData?: any;
}

const ScheduleForm: React.FC<ScheduleFormProps> = ({ onClose, onSubmit, initialData }) => {
  const [formData, setFormData] = useState<any>({
    name: initialData?.name || '',
    type: initialData?.type || 1,
    tolerance_minutes: initialData?.tolerance_minutes || 0,
    tolerance_checkin_minutes: initialData?.tolerance_checkin_minutes || 0,
    start_date: initialData?.start_date || '',
    end_date: initialData?.end_date || '',
    excluded_account_ids: initialData?.excluded_account_ids || [],
    rules: initialData?.rules || [],
    location_ids: initialData?.location_ids || []
  });

  const [locations, setLocations] = useState<Location[]>([]);
  const [accounts, setAccounts] = useState<Account[]>([]);
  const [activeTab, setActiveTab] = useState<'info' | 'toleransi' | 'rules' | 'locations' | 'exclusions'>('info');
  const [locationSearch, setLocationSearch] = useState('');
  const [exclusionSearch, setExclusionSearch] = useState('');

  const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];

  useEffect(() => {
    locationService.getAll().then(setLocations);
    accountService.getAll().then(accs => setAccounts(accs as Account[]));
    
    if (!initialData) {
      // Init default rules for type 1
      const initialRules = days.map((_, idx) => ({
        day_of_week: idx,
        check_in_time: '08:00',
        check_out_time: '17:00',
        is_holiday: idx === 0 || idx === 6 // Sun & Sat holiday by default
      }));
      setFormData(prev => ({ ...prev, rules: initialRules }));
    }
  }, []);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) => {
    const { name, value } = e.target;
    const parsedValue = (name === 'type' || name === 'tolerance_minutes' || name === 'tolerance_checkin_minutes') ? parseInt(value) || 0 : value;
    
    setFormData(prev => {
      const newData = { ...prev, [name]: parsedValue };
      
      // LOGIKA AUTO-RESET: Jika ganti ke Tipe 3, nolkan toleransi & rules
      if (name === 'type' && parsedValue === 3) {
        newData.tolerance_minutes = 0;
        newData.tolerance_checkin_minutes = 0;
        newData.rules = [];
        // Jika sedang di tab Toleransi atau Rules, pindah ke Info
        if (activeTab === 'toleransi' || activeTab === 'rules') setActiveTab('info');
      } else if (name === 'type' && parsedValue !== 3 && prev.type === 3) {
        // Balikin rules default jika pindah dari Tipe 3 ke kerja
        newData.rules = days.map((_, idx) => ({
          day_of_week: idx,
          check_in_time: '08:00',
          check_out_time: '17:00',
          is_holiday: parsedValue === 2 ? false : (idx === 0 || idx === 6)
        }));
      } else if (name === 'type' && parsedValue === 2) {
        // Jika pindah ke Tipe 2, pastikan semua is_holiday false
        newData.rules = prev.rules.map((rule: any) => ({ ...rule, is_holiday: false }));
      }

      return newData;
    });
  };

  const handleRuleChange = (idx: number, field: string, value: any) => {
    const newRules = [...formData.rules];
    newRules[idx] = { ...newRules[idx], [field]: value };
    setFormData(prev => ({ ...prev, rules: newRules }));
  };

  const toggleLocation = (id: string) => {
    setFormData(prev => ({
      ...prev,
      location_ids: prev.location_ids.includes(id) 
        ? prev.location_ids.filter(lid => lid !== id)
        : [...prev.location_ids, id]
    }));
  };

  const toggleExclusion = (id: string) => {
    setFormData(prev => ({
      ...prev,
      excluded_account_ids: prev.excluded_account_ids.includes(id)
        ? prev.excluded_account_ids.filter(aid => aid !== id)
        : [...prev.excluded_account_ids, id]
    }));
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    
    // Validasi Tanggal untuk Tipe Periodik
    if ((formData.type === 3 || formData.type === 4) && (!formData.start_date || !formData.end_date)) {
      alert('Tanggal Mulai dan Selesai wajib diisi untuk tipe jadwal ini.');
      return;
    }

    onSubmit(formData);
  };

  const filteredAccounts = accounts.filter(acc => 
    formData.location_ids.length === 0 || (acc.location_id && formData.location_ids.includes(acc.location_id))
  );

  // LOGIKA TABS DINAMIS
  const tabs = [
    { id: 'info', label: 'Informasi' },
  ];

  if (formData.type !== 3) {
    tabs.push({ id: 'toleransi', label: 'Toleransi' });
    tabs.push({ id: 'rules', label: 'Aturan Jam' });
  }

  tabs.push({ id: 'locations', label: 'Lokasi' });

  if (formData.type === 3 || formData.type === 4) {
    tabs.push({ id: 'exclusions', label: 'Pengecualian' });
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-40 p-4">
      <div className="bg-white rounded-md shadow-2xl w-full max-w-2xl overflow-hidden flex flex-col max-h-[90vh] animate-in zoom-in duration-200">
        <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between">
          <div>
            <h3 className="text-base font-bold text-[#006E62]">
              {initialData ? 'Ubah Jadwal' : 'Tambah Jadwal Baru'}
            </h3>
            <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Konfigurasi Jam & Lokasi</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={20} />
          </button>
        </div>

        {/* Tabs */}
        <div className="flex border-b border-gray-100 bg-gray-50/50 overflow-x-auto">
           {tabs.map(tab => (
              <button 
                key={tab.id}
                onClick={() => setActiveTab(tab.id as any)}
                className={`px-6 py-3 text-[10px] font-bold uppercase tracking-widest border-b-2 transition-all whitespace-nowrap ${
                  activeTab === tab.id ? 'border-[#006E62] text-[#006E62] bg-white' : 'border-transparent text-gray-400 hover:text-gray-600'
                }`}
              >
                {tab.label}
              </button>
           ))}
        </div>

        <form onSubmit={handleSubmit} className="flex-1 overflow-y-auto p-6 space-y-6">
          {activeTab === 'info' && (
            <div className="space-y-4 animate-in fade-in duration-200">
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-500 uppercase">Nama Jadwal</label>
                <input 
                  required 
                  name="name" 
                  value={formData.name} 
                  onChange={handleChange} 
                  placeholder="cth: Office Hour Pusat"
                  className="w-full px-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none" 
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-500 uppercase">Tipe Jadwal</label>
                <div className="grid grid-cols-2 gap-2">
                  {[
                    { id: 1, label: '1. Hari Kerja' },
                    { id: 2, label: '2. Shift Kerja' },
                    { id: 3, label: '3. Libur Khusus' },
                    { id: 4, label: '4. Hari Kerja Khusus' }
                  ].map(type => (
                    <button
                      key={type.id}
                      type="button"
                      onClick={() => handleChange({ target: { name: 'type', value: type.id.toString() } } as any)}
                      className={`px-3 py-2 text-[11px] font-bold text-left border rounded transition-all ${
                        formData.type === type.id 
                          ? 'bg-emerald-50 border-[#006E62] text-[#006E62]' 
                          : 'bg-white border-gray-200 text-gray-500 hover:border-gray-300'
                      }`}
                    >
                      {type.label}
                    </button>
                  ))}
                </div>
              </div>

              {formData.type >= 3 && (
                <div className="grid grid-cols-2 gap-4 p-4 bg-orange-50 border border-orange-100 rounded animate-in slide-in-from-top duration-200 mt-2">
                  <div className="space-y-1">
                    <label className="text-[10px] font-bold text-orange-600 uppercase">Mulai Tanggal</label>
                    <input type="date" required name="start_date" value={formData.start_date} onChange={handleChange} className="w-full px-3 py-2 text-xs border border-orange-200 rounded outline-none" />
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] font-bold text-orange-600 uppercase">Sampai Tanggal</label>
                    <input type="date" required name="end_date" value={formData.end_date} onChange={handleChange} className="w-full px-3 py-2 text-xs border border-orange-200 rounded outline-none" />
                  </div>
                </div>
               )}
            </div>
          )}

          {activeTab === 'toleransi' && formData.type !== 3 && (
            <div className="space-y-4 animate-in fade-in duration-200">
              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1">
                  <label className="text-[10px] font-bold text-gray-500 uppercase">Toleransi Datang (Menit)</label>
                  <input type="number" name="tolerance_checkin_minutes" value={formData.tolerance_checkin_minutes} onChange={handleChange} className="w-full px-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none" />
                </div>
                <div className="space-y-1">
                  <label className="text-[10px] font-bold text-gray-500 uppercase">Toleransi Pulang (Menit)</label>
                  <input type="number" name="tolerance_minutes" value={formData.tolerance_minutes} onChange={handleChange} className="w-full px-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none" />
                </div>
              </div>
              <div className="p-4 bg-blue-50 border border-blue-100 rounded">
                 <p className="text-[10px] text-blue-700 font-medium">Batas waktu keterlambatan datang dan batas akhir presensi pulang setelah jam kerja berakhir.</p>
              </div>
            </div>
          )}

          {activeTab === 'rules' && formData.type !== 3 && (
            <div className="space-y-4 animate-in fade-in duration-200">
                 <div className="space-y-2">
                   <p className="text-[10px] font-bold text-gray-400 uppercase mb-2">Tentukan Hari & Jam Kerja</p>
                   {days.map((day, idx) => (
                      <div key={idx} className={`flex items-center gap-4 p-3 rounded border ${formData.rules[idx]?.is_holiday ? 'bg-rose-50 border-rose-100 opacity-60' : 'bg-white border-gray-100'}`}>
                         <div className="w-20 shrink-0 text-[11px] font-bold text-gray-700">{day}</div>
                         <div className="flex-1 flex items-center gap-4">
                            <input type="time" disabled={formData.rules[idx]?.is_holiday} value={formData.rules[idx]?.check_in_time} onChange={(e) => handleRuleChange(idx, 'check_in_time', e.target.value)} className="flex-1 px-2 py-1 text-xs border rounded disabled:bg-transparent" />
                            <span className="text-gray-300">-</span>
                            <input type="time" disabled={formData.rules[idx]?.is_holiday} value={formData.rules[idx]?.check_out_time} onChange={(e) => handleRuleChange(idx, 'check_out_time', e.target.value)} className="flex-1 px-2 py-1 text-xs border rounded disabled:bg-transparent" />
                         </div>
                         {formData.type !== 2 && (
                           <label className="flex items-center gap-2 cursor-pointer shrink-0">
                              <input type="checkbox" checked={formData.rules[idx]?.is_holiday} onChange={(e) => handleRuleChange(idx, 'is_holiday', e.target.checked)} className="rounded border-gray-300 text-[#006E62]" />
                              <span className="text-[10px] font-bold text-gray-400 uppercase">Libur</span>
                           </label>
                         )}
                      </div>
                   ))}
                 </div>
            </div>
          )}

          {activeTab === 'locations' && (
            <div className="animate-in fade-in duration-200 space-y-4">
               <p className="text-[10px] font-bold text-gray-400 uppercase">Pilih Lokasi yang Terpengaruh Jadwal Ini</p>
               
               <div className="space-y-3">
                  <div className="relative">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={14} />
                    <input 
                      type="text"
                      placeholder="Cari Lokasi..."
                      value={locationSearch}
                      onChange={(e) => setLocationSearch(e.target.value)}
                      className="w-full pl-9 pr-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none"
                    />
                  </div>

                  <div className="border border-gray-100 rounded-md overflow-hidden bg-gray-50/30">
                    <div className="max-h-60 overflow-y-auto p-1 space-y-1">
                       {locations
                        .filter(loc => loc.name.toLowerCase().includes(locationSearch.toLowerCase()))
                        .map(loc => (
                          <div 
                            key={loc.id}
                            onClick={() => toggleLocation(loc.id)}
                            className={`flex items-center gap-3 p-2 rounded cursor-pointer text-xs transition-colors ${
                              formData.location_ids.includes(loc.id) 
                                ? 'bg-emerald-50 text-[#006E62] font-bold border border-emerald-100' 
                                : 'text-gray-600 hover:bg-white border border-transparent'
                            }`}
                          >
                             <div className={`w-4 h-4 border rounded flex items-center justify-center shrink-0 ${
                               formData.location_ids.includes(loc.id) ? 'bg-[#006E62] border-[#006E62] text-white' : 'border-gray-300 bg-white'
                             }`}>
                                {formData.location_ids.includes(loc.id) && <Check size={10} />}
                             </div>
                             <span className="truncate">{loc.name}</span>
                          </div>
                       ))}
                       {locations.filter(loc => loc.name.toLowerCase().includes(locationSearch.toLowerCase())).length === 0 && (
                         <div className="py-8 text-center text-gray-400 italic text-[11px]">Tidak ada lokasi ditemukan.</div>
                       )}
                    </div>
                  </div>
               </div>

               <div className="pt-4 border-t border-gray-50">
                  <p className="text-[9px] font-bold text-gray-400 uppercase mb-2 tracking-widest">{formData.location_ids.length} Lokasi Terpilih</p>
                  <div className="flex flex-wrap gap-2">
                    {formData.location_ids.map((lid: string) => {
                      const loc = locations.find(l => l.id === lid);
                      return loc ? (
                        <span key={lid} className="px-2 py-1 bg-emerald-50 text-[#006E62] text-[10px] font-bold rounded flex items-center gap-1 border border-emerald-100">
                          {loc.name}
                          <X size={10} className="cursor-pointer" onClick={() => toggleLocation(lid)} />
                        </span>
                      ) : null;
                    })}
                  </div>
               </div>
            </div>
          )}

          {activeTab === 'exclusions' && (
            <div className="animate-in fade-in duration-200 space-y-4">
               <div className="bg-blue-50 p-3 border border-blue-100 rounded">
                  <p className="text-[10px] text-blue-700 font-medium italic">Pilih user di bawah untuk MENGECEUALIKAN mereka dari jadwal tipe ini.</p>
               </div>

               <div className="space-y-3">
                  <div className="relative">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={14} />
                    <input 
                      type="text"
                      placeholder="Cari Nama Karyawan..."
                      value={exclusionSearch}
                      onChange={(e) => setExclusionSearch(e.target.value)}
                      className="w-full pl-9 pr-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none"
                    />
                  </div>

                  <div className="border border-gray-100 rounded-md overflow-hidden bg-gray-50/30">
                    <div className="max-h-60 overflow-y-auto p-1 space-y-1">
                       {filteredAccounts
                        .filter(acc => acc.full_name.toLowerCase().includes(exclusionSearch.toLowerCase()))
                        .map(acc => (
                          <div 
                            key={acc.id}
                            onClick={() => toggleExclusion(acc.id)}
                            className={`flex items-center gap-3 p-2 rounded cursor-pointer text-xs transition-colors ${
                              formData.excluded_account_ids.includes(acc.id) 
                                ? 'bg-rose-50 text-rose-600 font-bold border border-rose-100' 
                                : 'text-gray-600 hover:bg-white border border-transparent'
                            }`}
                          >
                             <div className={`w-4 h-4 border rounded flex items-center justify-center shrink-0 ${
                               formData.excluded_account_ids.includes(acc.id) ? 'bg-rose-500 border-rose-500 text-white' : 'border-gray-300 bg-white'
                             }`}>
                                {formData.excluded_account_ids.includes(acc.id) && <Check size={10} />}
                             </div>
                             <div>
                               <p className="leading-tight">{acc.full_name}</p>
                               <p className="text-[9px] font-normal opacity-70">{(acc as any).location?.name || 'Tanpa Lokasi'}</p>
                             </div>
                          </div>
                       ))}
                       {filteredAccounts.filter(acc => acc.full_name.toLowerCase().includes(exclusionSearch.toLowerCase())).length === 0 && (
                         <div className="py-8 text-center text-gray-400 italic text-[11px]">Tidak ada karyawan ditemukan.</div>
                       )}
                    </div>
                  </div>
               </div>

               <div className="pt-4 border-t border-gray-50">
                  <p className="text-[9px] font-bold text-gray-400 uppercase mb-2 tracking-widest">{formData.excluded_account_ids.length} User Dikecualikan</p>
                  <div className="flex flex-wrap gap-2">
                    {formData.excluded_account_ids.map((aid: string) => {
                      const acc = accounts.find(a => a.id === aid);
                      return acc ? (
                        <span key={aid} className="px-2 py-1 bg-rose-50 text-rose-600 text-[10px] font-bold rounded flex items-center gap-1 border border-rose-100">
                          {acc.full_name}
                          <X size={10} className="cursor-pointer" onClick={() => toggleExclusion(aid)} />
                        </span>
                      ) : null;
                    })}
                  </div>
               </div>
            </div>
          )}
        </form>

        <div className="px-6 py-4 border-t border-gray-100 flex justify-end gap-3 bg-gray-50">
          <button type="button" onClick={onClose} className="px-4 py-2 text-xs font-bold text-gray-500 uppercase">Batal</button>
          <button 
            onClick={handleSubmit}
            className="flex items-center gap-2 bg-[#006E62] text-white px-8 py-2 rounded shadow-md hover:bg-[#005a50] transition-all text-xs font-bold uppercase"
          >
            <Save size={14} /> Simpan Jadwal
          </button>
        </div>
      </div>
    </div>
  );
};

export default ScheduleForm;
