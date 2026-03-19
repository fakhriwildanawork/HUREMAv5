import React, { useState, useEffect, useRef } from 'react';
import { X, Save, Upload, FileText, Paperclip, ChevronDown, Calendar, CalendarClock, ExternalLink } from 'lucide-react';
import { locationService } from '../../services/locationService';
import { googleDriveService } from '../../services/googleDriveService';
import { accountService } from '../../services/accountService';
import { scheduleService } from '../../services/scheduleService';
import { Location, Schedule } from '../../types';
import Swal from 'sweetalert2';

interface LogFormProps {
  type: 'career' | 'health';
  accountId: string;
  initialData?: any; // Ini bisa data akun (untuk add) atau data log (untuk edit)
  isEdit?: boolean;
  onClose: () => void;
  onSubmit: (data: any) => void;
}

const LogForm: React.FC<LogFormProps> = ({ type, accountId, initialData, isEdit = false, onClose, onSubmit }) => {
  const [formData, setFormData] = useState<any>({
    account_id: accountId,
    // Career Fields
    position: initialData?.position || '',
    grade: initialData?.grade || '',
    location_id: initialData?.location_id || '',
    location_name: initialData?.location_name || (initialData?.location?.name || ''),
    schedule_id: initialData?.schedule_id || (
      initialData?.schedule_type === 'Fleksibel' ? 'FLEKSIBEL' : 
      initialData?.schedule_type === 'Shift Dinamis' ? 'DINAMIS' : ''
    ),
    file_sk_id: initialData?.file_sk_id || '',
    // Health Fields
    mcu_status: initialData?.mcu_status || '',
    health_risk: initialData?.health_risk || '',
    file_mcu_id: initialData?.file_mcu_id || '',
    // Common
    notes: initialData?.notes || '',
    change_date: initialData?.change_date ? initialData.change_date.split('T')[0] : new Date().toISOString().split('T')[0],
  });

  const [locations, setLocations] = useState<Location[]>([]);
  const [schedules, setSchedules] = useState<Schedule[]>([]);
  const [suggestions, setSuggestions] = useState<{ positions: string[], grades: string[] }>({ positions: [], grades: [] });
  const [uploading, setUploading] = useState(false);
  
  // Custom Dropdown States
  const [showPosDropdown, setShowPosDropdown] = useState(false);
  const [showGradeDropdown, setShowGradeDropdown] = useState(false);

  const posRef = useRef<HTMLDivElement>(null);
  const gradeRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (type === 'career') {
      locationService.getAll().then(setLocations);
      accountService.getDistinctAttributes().then(setSuggestions);
    }

    // Close dropdowns on click outside
    const handleClickOutside = (event: MouseEvent) => {
      if (posRef.current && !posRef.current.contains(event.target as Node)) setShowPosDropdown(false);
      if (gradeRef.current && !gradeRef.current.contains(event.target as Node)) setShowGradeDropdown(false);
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [type]);

  // Effect to handle dynamic dependent schedule dropdown in Career Log
  useEffect(() => {
    if (type === 'career' && formData.location_id) {
       scheduleService.getByLocation(formData.location_id).then(data => {
         const filtered = data.filter(s => s.type === 1 || s.type === 2);
         setSchedules(filtered);
       });
    } else {
       setSchedules([]);
    }
  }, [type, formData.location_id]);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) => {
    const { name, value } = e.target;
    
    if (name === 'location_id') {
      const selected = locations.find(l => l.id === value);
      setFormData(prev => {
        const updated = { 
          ...prev, 
          location_id: value, 
          location_name: selected ? selected.name : ''
        };
        
        // UX: Hanya reset schedule jika lokasi diubah secara manual dan berbeda dari data awal
        if (value !== initialData?.location_id) {
           updated.schedule_id = '';
        }
        
        return updated;
      });
    } else {
      setFormData(prev => ({ ...prev, [name]: value }));
    }
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setUploading(true);
    try {
      const folderName = type === 'career' ? 'SK_KARYAWAN' : 'MCU_KARYAWAN';
      const fileId = await googleDriveService.uploadFile(file, folderName);
      setFormData(prev => ({ ...prev, [type === 'career' ? 'file_sk_id' : 'file_mcu_id']: fileId }));
      Swal.fire({ title: 'Berhasil!', text: 'File telah diunggah ke Google Drive.', icon: 'success', timer: 1000, showConfirmButton: false });
    } catch (error) {
      Swal.fire('Gagal', 'Gagal mengunggah file', 'error');
    } finally {
      setUploading(false);
    }
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    
    // Filtrasi Payload Akhir sebelum dikirim ke Service (Mencegah polusi data antar tabel)
    // Penanganan string kosong menjadi null akan diproses di accountService via sanitizePayload
    let finalPayload: any = {
      account_id: formData.account_id,
      notes: formData.notes,
      change_date: formData.change_date
    };

    if (type === 'career') {
      // Map virtual types for schedule_type logic (handled in accountService/createCareerLog)
      let finalScheduleType = '';
      let finalScheduleId = formData.schedule_id;

      if (finalScheduleId === 'FLEKSIBEL') {
        finalScheduleType = 'Fleksibel';
        finalScheduleId = '';
      } else if (finalScheduleId === 'DINAMIS') {
        finalScheduleType = 'Shift Dinamis';
        finalScheduleId = '';
      } else {
        const sch = schedules.find(s => s.id === finalScheduleId);
        if (sch) finalScheduleType = sch.name;
      }

      finalPayload = {
        ...finalPayload,
        position: formData.position,
        grade: formData.grade,
        location_id: formData.location_id,
        location_name: formData.location_name,
        schedule_id: finalScheduleId || null,
        schedule_type: finalScheduleType, // Service will update the main account schedule_type
        file_sk_id: formData.file_sk_id
      };
    } else if (type === 'health') {
      finalPayload = {
        ...finalPayload,
        mcu_status: formData.mcu_status,
        health_risk: formData.health_risk,
        file_mcu_id: formData.file_mcu_id
      };
    }

    if (isEdit) {
      onSubmit({ ...finalPayload, id: initialData.id });
    } else {
      onSubmit(finalPayload);
    }
  };

  const filteredPositions = suggestions.positions.filter(p => 
    p.toLowerCase().includes(formData.position.toLowerCase())
  );

  const filteredGrades = suggestions.grades.filter(g => 
    g.toLowerCase().includes(formData.grade.toLowerCase())
  );

  // Cek ketersediaan tipe shift (type 2) di lokasi terpilih
  const hasShiftSchedules = schedules.some(s => s.type === 2);

  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black bg-opacity-40 p-4">
      <div className="bg-white rounded-md shadow-2xl w-full max-w-md overflow-hidden animate-in zoom-in duration-200">
        <div className="px-5 py-3 border-b border-gray-100 flex items-center justify-between">
          <div>
            <h3 className="text-sm font-bold text-[#006E62]">
              {isEdit ? 'Ubah' : 'Tambah'} Riwayat {type === 'career' ? 'Karir' : 'Kesehatan'}
            </h3>
            <p className="text-[9px] text-gray-400 font-bold uppercase tracking-widest">Pencatatan Riwayat Manual</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600">
            <X size={18} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-5 space-y-4">
          <div className="space-y-1">
            <label htmlFor="change_date" className="text-[9px] font-bold text-gray-500 uppercase">Tanggal Perubahan</label>
            <div className="relative">
              <input 
                id="change_date"
                type="date"
                required 
                name="change_date" 
                value={formData.change_date} 
                onChange={handleChange} 
                className="w-full px-2 py-1.5 text-xs border border-gray-200 rounded outline-none focus:ring-1 focus:ring-[#006E62] bg-gray-50" 
              />
            </div>
          </div>

          {type === 'career' ? (
            <>
          <div className="grid grid-cols-2 gap-3">
                <div className="space-y-1 relative" ref={posRef}>
                  <label htmlFor="position" className="text-[9px] font-bold text-gray-500 uppercase">Jabatan</label>
                  <div className="relative">
                    <input 
                      id="position"
                      required 
                      name="position" 
                      autoComplete="off"
                      value={formData.position} 
                      onChange={(e) => { handleChange(e); setShowPosDropdown(true); }}
                      onFocus={() => setShowPosDropdown(true)}
                      className="w-full px-2 py-1.5 text-xs border border-gray-200 rounded outline-none focus:ring-1 focus:ring-[#006E62] pr-7 bg-white" 
                      placeholder="Pilih atau Ketik"
                    />
                    <button 
                      type="button"
                      onClick={() => setShowPosDropdown(!showPosDropdown)}
                      className="absolute right-0 top-0 bottom-0 px-2 flex items-center text-gray-400 hover:text-[#006E62]"
                    >
                      <ChevronDown size={14} />
                    </button>
                  </div>
                  {showPosDropdown && filteredPositions.length > 0 && (
                    <div className="absolute z-[70] w-full mt-1 bg-white border border-gray-100 rounded shadow-lg max-h-40 overflow-y-auto">
                      {filteredPositions.map(p => (
                        <div 
                          key={p} 
                          className="px-3 py-2 text-xs hover:bg-gray-50 cursor-pointer text-gray-700"
                          onClick={() => {
                            setFormData(prev => ({ ...prev, position: p }));
                            setShowPosDropdown(false);
                          }}
                        >
                          {p}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
                <div className="space-y-1 relative" ref={gradeRef}>
                  <label htmlFor="grade" className="text-[9px] font-bold text-gray-500 uppercase">Golongan</label>
                  <div className="relative">
                    <input 
                      id="grade"
                      name="grade" 
                      autoComplete="off"
                      value={formData.grade} 
                      onChange={(e) => { handleChange(e); setShowGradeDropdown(true); }}
                      onFocus={() => setShowGradeDropdown(true)}
                      className="w-full px-2 py-1.5 text-xs border border-gray-200 rounded outline-none focus:ring-1 focus:ring-[#006E62] pr-7 bg-white" 
                      placeholder="Pilih atau Ketik"
                    />
                    <button 
                      type="button"
                      onClick={() => setShowGradeDropdown(!showGradeDropdown)}
                      className="absolute right-0 top-0 bottom-0 px-2 flex items-center text-gray-400 hover:text-[#006E62]"
                    >
                      <ChevronDown size={14} />
                    </button>
                  </div>
                  {showGradeDropdown && filteredGrades.length > 0 && (
                    <div className="absolute z-[70] w-full mt-1 bg-white border border-gray-100 rounded shadow-lg max-h-40 overflow-y-auto">
                      {filteredGrades.map(g => (
                        <div 
                          key={g} 
                          className="px-3 py-2 text-xs hover:bg-gray-50 cursor-pointer text-gray-700"
                          onClick={() => {
                            setFormData(prev => ({ ...prev, grade: g }));
                            setShowGradeDropdown(false);
                          }}
                        >
                          {g}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
              <div className="space-y-1">
                <label htmlFor="location_id" className="text-[9px] font-bold text-gray-500 uppercase">Lokasi Penempatan</label>
                <div className="relative">
                  <select id="location_id" required name="location_id" value={formData.location_id} onChange={handleChange} className="w-full px-2 py-1.5 text-xs border border-gray-200 rounded outline-none focus:ring-1 focus:ring-[#006E62] appearance-none bg-white">
                    <option value="">-- Pilih Lokasi --</option>
                    {locations.map(l => <option key={l.id} value={l.id}>{l.name}</option>)}
                  </select>
                  <ChevronDown size={14} className="absolute right-2 top-1/2 -translate-y-1/2 text-gray-400 pointer-events-none" />
                </div>
              </div>
              <div className="space-y-1">
                <label htmlFor="schedule_id" className="text-[9px] font-bold text-gray-500 uppercase">Jadwal Kerja Baru</label>
                <div className="relative">
                  <select 
                    id="schedule_id"
                    required 
                    name="schedule_id" 
                    value={formData.schedule_id} 
                    onChange={handleChange} 
                    disabled={!formData.location_id}
                    className="w-full px-2 py-1.5 text-xs border border-gray-200 rounded outline-none focus:ring-1 focus:ring-[#006E62] appearance-none bg-white disabled:bg-gray-50"
                  >
                    <option value="">-- {formData.location_id ? 'Pilih Jadwal' : 'Pilih Lokasi Terlebih Dahulu'} --</option>
                    <option value="FLEKSIBEL">Fleksibel</option>
                    <option value="DINAMIS" disabled={!hasShiftSchedules}>
                      Shift Dinamis{!hasShiftSchedules ? '(Shift Tidak Tersedia)' : '(Pilih Saat Presensi)'}
                    </option>
                    {schedules.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
                  </select>
                  <CalendarClock size={14} className="absolute right-2 top-1/2 -translate-y-1/2 text-gray-400 pointer-events-none" />
                </div>
              </div>
              <div className="space-y-1">
                <label htmlFor="file_sk_id" className="text-[9px] font-bold text-gray-500 uppercase">Dokumen SK (PDF/Gambar)</label>
                <div className={`flex items-center gap-3 p-2 bg-gray-50 border border-dashed rounded cursor-pointer hover:bg-white transition-colors ${formData.file_sk_id ? 'border-[#006E62]' : 'border-gray-200'}`}>
                  <label htmlFor="file_sk_id" className="flex items-center gap-2 cursor-pointer w-full">
                    <div className="p-2 bg-white rounded border border-gray-100 shrink-0">
                      <Upload size={14} className={formData.file_sk_id ? 'text-[#006E62]' : 'text-gray-300'} />
                    </div>
                    <div className="flex-1 min-w-0">
                      <p className="text-[10px] font-bold text-gray-600 uppercase">
                        {uploading ? 'Sedang Mengunggah...' : formData.file_sk_id ? 'SK Terunggah' : 'Upload File SK'}
                      </p>
                      <p className="text-[8px] text-gray-400 truncate">{formData.file_sk_id || 'ID akan tersimpan di G-Drive'}</p>
                    </div>
                    {formData.file_sk_id && (
                      <a 
                        href={googleDriveService.getFileUrl(formData.file_sk_id).replace('=s1600', '=s0')} 
                        target="_blank" 
                        rel="noreferrer"
                        className="p-1.5 bg-white border border-gray-100 rounded text-[#006E62] hover:bg-emerald-50 transition-colors"
                        onClick={(e) => e.stopPropagation()}
                      >
                        <ExternalLink size={12} />
                      </a>
                    )}
                    <input id="file_sk_id" type="file" className="hidden" onChange={handleFileUpload} disabled={uploading} />
                  </label>
                </div>
              </div>
            </>
          ) : (
            <>
              <div className="grid grid-cols-2 gap-3">
                <div className="space-y-1">
                  <label htmlFor="mcu_status" className="text-[9px] font-bold text-gray-500 uppercase">Status Medis / MCU</label>
                  <select id="mcu_status" required name="mcu_status" value={formData.mcu_status} onChange={handleChange} className="w-full px-2 py-1.5 text-xs border border-gray-200 rounded outline-none focus:ring-1 focus:ring-[#006E62] bg-white">
                    <option value="">-- Pilih Status --</option>
                    <option value="FIT">FIT</option>
                    <option value="FIT WITH NOTE">FIT WITH NOTE</option>
                    <option value="UNFIT">UNFIT</option>
                    <option value="PENDING">PENDING</option>
                  </select>
                </div>
                <div className="space-y-1">
                  <label htmlFor="health_risk" className="text-[9px] font-bold text-gray-500 uppercase">Risiko Kesehatan</label>
                  <select id="health_risk" required name="health_risk" value={formData.health_risk} onChange={handleChange} className="w-full px-2 py-1.5 text-xs border border-gray-200 rounded outline-none focus:ring-1 focus:ring-[#006E62] bg-white">
                    <option value="">-- Pilih Risiko --</option>
                    <option value="LOW">LOW</option>
                    <option value="MEDIUM">MEDIUM</option>
                    <option value="HIGH">HIGH</option>
                    <option value="VERY HIGH">VERY HIGH</option>
                  </select>
                </div>
              </div>
              <div className="space-y-1">
                <label htmlFor="file_mcu_id" className="text-[9px] font-bold text-gray-500 uppercase">Upload Hasil MCU (PDF/Gambar)</label>
                <div className={`flex items-center gap-3 p-2 bg-gray-50 border border-dashed rounded cursor-pointer hover:bg-white transition-colors ${formData.file_mcu_id ? 'border-red-500' : 'border-gray-200'}`}>
                  <label htmlFor="file_mcu_id" className="flex items-center gap-2 cursor-pointer w-full">
                    <div className="p-2 bg-white rounded border border-gray-100 shrink-0">
                      <Upload size={14} className={formData.file_mcu_id ? 'text-red-500' : 'text-gray-300'} />
                    </div>
                    <div className="flex-1 min-w-0">
                      <p className="text-[10px] font-bold text-gray-600 uppercase">
                        {uploading ? 'Sedang Mengunggah...' : formData.file_mcu_id ? 'Hasil MCU Terunggah' : 'Upload Hasil MCU'}
                      </p>
                      <p className="text-[8px] text-gray-400 truncate">{formData.file_mcu_id || 'ID akan tersimpan di G-Drive'}</p>
                    </div>
                    {formData.file_mcu_id && (
                      <a 
                        href={googleDriveService.getFileUrl(formData.file_mcu_id).replace('=s1600', '=s0')} 
                        target="_blank" 
                        rel="noreferrer"
                        className="p-1.5 bg-white border border-gray-100 rounded text-red-500 hover:bg-red-50 transition-colors"
                        onClick={(e) => e.stopPropagation()}
                      >
                        <ExternalLink size={12} />
                      </a>
                    )}
                    <input id="file_mcu_id" type="file" className="hidden" onChange={handleFileUpload} disabled={uploading} />
                  </label>
                </div>
              </div>
            </>
          )}

          <div className="space-y-1">
            <label htmlFor="notes" className="text-[9px] font-bold text-gray-500 uppercase">Catatan Tambahan</label>
            <textarea id="notes" name="notes" value={formData.notes} onChange={handleChange} rows={2} className="w-full px-2 py-1.5 text-xs border border-gray-200 rounded outline-none focus:ring-1 focus:ring-[#006E62] resize-none" />
          </div>

          <div className="flex justify-end gap-3 pt-2">
            <button type="button" onClick={onClose} className="text-[10px] font-bold text-gray-400 uppercase">Batal</button>
            <button 
              type="submit" 
              disabled={uploading}
              className="flex items-center gap-2 bg-[#006E62] text-white px-5 py-1.5 rounded text-[10px] font-bold uppercase shadow-md hover:bg-[#005a50] disabled:opacity-50"
            >
              <Save size={12} /> {isEdit ? 'Simpan Perubahan' : 'Simpan Log'}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};

export default LogForm;