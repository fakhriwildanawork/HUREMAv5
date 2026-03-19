
import React, { useState, useEffect, useRef } from 'react';
import { X, Save, Upload, Award, Calendar, ChevronDown, Check, UserCircle, Loader2 } from 'lucide-react';
import Swal from 'sweetalert2';
import { certificationService } from '../../services/certificationService';
import { accountService } from '../../services/accountService';
import { googleDriveService } from '../../services/googleDriveService';
import { AccountCertificationInput, Account } from '../../types';

interface CertificationFormModalProps {
  onClose: () => void;
  onSuccess: () => void;
  initialData?: any;
}

const CertificationFormModal: React.FC<CertificationFormModalProps> = ({ onClose, onSuccess, initialData }) => {
  const [formData, setFormData] = useState<any>({
    account_id: initialData?.account_id || '',
    entry_date: initialData?.entry_date || new Date().toISOString().split('T')[0],
    cert_type: initialData?.cert_type || '',
    cert_name: initialData?.cert_name || '',
    cert_date: initialData?.cert_date || '',
    file_id: initialData?.file_id || '',
    notes: initialData?.notes || ''
  });

  const [accounts, setAccounts] = useState<Account[]>([]);
  const [existingTypes, setExistingTypes] = useState<string[]>([]);
  const [showTypeDropdown, setShowTypeDropdown] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  
  const typeRef = useRef<HTMLDivElement>(null);

  // Menentukan apakah dropdown karyawan harus muncul
  const isAccountPredefined = !!initialData?.account_id && !initialData?.id;

  useEffect(() => {
    if (!isAccountPredefined) {
       accountService.getAll().then(data => setAccounts(data as any));
    }
    certificationService.getUniqueCertTypes().then((types: any) => setExistingTypes(types));

    const handleClickOutside = (event: MouseEvent) => {
      if (typeRef.current && !typeRef.current.contains(event.target as Node)) setShowTypeDropdown(false);
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) => {
    const { name, value } = e.target;
    setFormData(prev => ({ ...prev, [name]: value }));
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      setUploading(true);
      const fileId = await googleDriveService.uploadFile(file);
      setFormData(prev => ({ ...prev, file_id: fileId }));
    } catch (error) {
      Swal.fire('Gagal', 'Gagal mengunggah file.', 'error');
    } finally {
      setUploading(false);
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.account_id) return Swal.fire('Peringatan', 'Pilih karyawan terlebih dahulu.', 'warning');
    
    try {
      setIsSaving(true);
      if (initialData?.id) {
        await certificationService.update(initialData.id, formData);
      } else {
        await certificationService.create(formData);
      }
      Swal.fire({ title: 'Berhasil!', text: 'Data sertifikasi telah disimpan.', icon: 'success', timer: 1500, showConfirmButton: false });
      onSuccess();
    } catch (error) {
      Swal.fire('Gagal', 'Gagal menyimpan data.', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const filteredTypes = existingTypes.filter(t => 
    t.toLowerCase().includes(formData.cert_type.toLowerCase())
  );

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-40 p-4">
      <div className="bg-white rounded-md shadow-2xl w-full max-w-lg overflow-hidden flex flex-col max-h-[90vh] animate-in zoom-in duration-200">
        <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between">
          <div>
            <h3 className="text-base font-bold text-[#006E62]">{initialData?.id ? 'Ubah' : 'Tambah'} Sertifikasi</h3>
            <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Pencatatan Dokumen Sertifikat</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={20} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="flex-1 overflow-y-auto p-6 space-y-4">
          {!isAccountPredefined && (
            <div className="space-y-1">
              <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Pilih Karyawan</label>
              <div className="relative">
                <select 
                  required
                  name="account_id"
                  value={formData.account_id}
                  onChange={handleChange}
                  className="w-full px-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none bg-gray-50 appearance-none"
                >
                  <option value="">-- Cari Karyawan --</option>
                  {accounts.map(acc => (
                    <option key={acc.id} value={acc.id}>{acc.full_name} ({acc.internal_nik})</option>
                  ))}
                </select>
                <UserCircle className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 pointer-events-none" size={14} />
              </div>
            </div>
          )}

          <div className="grid grid-cols-2 gap-4">
            <div className="space-y-1">
              <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Tanggal Input</label>
              <input 
                type="date"
                required
                name="entry_date"
                value={formData.entry_date}
                onChange={handleChange}
                className="w-full px-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none bg-emerald-50/30"
              />
            </div>
            <div className="space-y-1">
              <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Tanggal Sertifikasi</label>
              <input 
                type="date"
                required
                name="cert_date"
                value={formData.cert_date}
                onChange={handleChange}
                className="w-full px-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none"
              />
            </div>
          </div>

          <div className="space-y-1 relative" ref={typeRef}>
            <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Jenis Sertifikasi</label>
            <div className="relative">
              <input 
                autoComplete="off"
                required
                name="cert_type"
                value={formData.cert_type}
                onChange={(e) => { handleChange(e); setShowTypeDropdown(true); }}
                onFocus={() => setShowTypeDropdown(true)}
                className="w-full px-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none bg-white pr-8"
                placeholder="Pilih atau ketik jenis baru..."
              />
              <button 
                type="button"
                onClick={() => setShowTypeDropdown(!showTypeDropdown)}
                className="absolute right-0 top-0 bottom-0 px-3 flex items-center text-gray-400"
              >
                <ChevronDown size={14} />
              </button>
            </div>
            {showTypeDropdown && filteredTypes.length > 0 && (
              <div className="absolute z-50 w-full mt-1 bg-white border border-gray-100 rounded shadow-xl max-h-40 overflow-y-auto">
                {filteredTypes.map(type => (
                  <div 
                    key={type}
                    onClick={() => { setFormData(prev => ({ ...prev, cert_type: type })); setShowTypeDropdown(false); }}
                    className="px-4 py-2 text-xs hover:bg-gray-50 cursor-pointer flex items-center justify-between text-gray-700"
                  >
                    <span>{type}</span>
                    {formData.cert_type === type && <Check size={12} className="text-[#006E62]" />}
                  </div>
                ))}
              </div>
            )}
          </div>

          <div className="space-y-1">
            <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Nama Sertifikasi</label>
            <input 
              required
              name="cert_name"
              value={formData.cert_name}
              onChange={handleChange}
              className="w-full px-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none"
              placeholder="cth: Ahli K3 Umum, TOEFL, etc."
            />
          </div>

          <div className="space-y-1">
            <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Upload Sertifikat (G-Drive)</label>
            <div className={`flex items-center gap-3 p-3 bg-gray-50 border border-dashed rounded-md transition-colors ${formData.file_id ? 'border-[#006E62] bg-emerald-50/20' : 'border-gray-200'}`}>
              <label className="flex items-center gap-3 cursor-pointer flex-1">
                <div className="p-2 bg-white rounded border border-gray-100 shrink-0">
                  <Upload size={16} className={formData.file_id ? 'text-[#006E62]' : 'text-gray-300'} />
                </div>
                <div className="flex-1">
                  <p className="text-[10px] font-bold text-gray-600 uppercase">
                    {uploading ? 'Mengunggah...' : formData.file_id ? 'DOKUMEN TERUNGGAH' : 'PILIH FILE (PDF/GAMBAR)'}
                  </p>
                  <p className="text-[8px] text-gray-400 truncate tracking-tighter">ID: {formData.file_id || 'Otomatis Simpan di Drive'}</p>
                </div>
                <input type="file" className="hidden" accept="image/*,application/pdf" onChange={handleFileUpload} disabled={uploading} />
              </label>
            </div>
          </div>

          <div className="space-y-1">
            <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Keterangan</label>
            <textarea 
              name="notes"
              value={formData.notes}
              onChange={handleChange}
              rows={2}
              className="w-full px-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none resize-none"
            />
          </div>
        </form>

        <div className="px-6 py-4 border-t border-gray-100 flex justify-end gap-3 bg-gray-50">
          <button type="button" onClick={onClose} className="px-4 py-2 text-xs font-bold text-gray-500 uppercase">Batal</button>
          <button 
            onClick={handleSubmit}
            disabled={uploading || isSaving}
            className="flex items-center gap-2 bg-[#006E62] text-white px-8 py-2 rounded shadow-md hover:bg-[#005a50] transition-all text-xs font-bold uppercase disabled:opacity-50"
          >
            {isSaving ? <Loader2 size={14} className="animate-spin" /> : <Save size={14} />}
            {isSaving ? 'Menyimpan...' : 'Simpan Sertifikasi'}
          </button>
        </div>
      </div>
    </div>
  );
};

export default CertificationFormModal;
