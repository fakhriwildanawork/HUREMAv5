import React, { useState, useEffect, useRef } from 'react';
import { X, Save, Upload, FileText, ChevronDown, Check, Users, Search, FolderKanban } from 'lucide-react';
import Swal from 'sweetalert2';
import { documentService } from '../../services/documentService';
import { accountService } from '../../services/accountService';
import { googleDriveService } from '../../services/googleDriveService';
import { DigitalDocument, Account } from '../../types';

interface DocumentFormProps {
  onClose: () => void;
  onSuccess: () => void;
  initialData?: DigitalDocument;
}

const DocumentForm: React.FC<DocumentFormProps> = ({ onClose, onSuccess, initialData }) => {
  const [formData, setFormData] = useState<any>({
    name: initialData?.name || '',
    doc_type: initialData?.doc_type || 'SOP',
    file_id: initialData?.file_id || '',
    description: initialData?.description || '',
    allowed_account_ids: initialData?.allowed_account_ids || []
  });

  const [accounts, setAccounts] = useState<Account[]>([]);
  const [existingTypes, setExistingTypes] = useState<string[]>([]);
  const [showTypeDropdown, setShowTypeDropdown] = useState(false);
  const [showUserDropdown, setShowUserDropdown] = useState(false);
  const [userSearchTerm, setUserSearchTerm] = useState('');
  const [uploading, setUploading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);

  const typeRef = useRef<HTMLDivElement>(null);
  const userRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    accountService.getAll().then(data => setAccounts(data as any));
    documentService.getUniqueDocTypes().then((types: any) => setExistingTypes(types));

    const handleClickOutside = (event: MouseEvent) => {
      if (typeRef.current && !typeRef.current.contains(event.target as Node)) setShowTypeDropdown(false);
      if (userRef.current && !userRef.current.contains(event.target as Node)) setShowUserDropdown(false);
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => {
    const { name, value } = e.target;
    setFormData(prev => ({ ...prev, [name]: value }));
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      setUploading(true);
      const fileId = await googleDriveService.uploadFile(file);
      setFormData(prev => ({ 
        ...prev, 
        file_id: fileId,
        name: prev.name || file.name.split('.').slice(0, -1).join('.')
      }));
    } catch (error) {
      Swal.fire('Gagal', 'Gagal mengunggah file ke Google Drive.', 'error');
    } finally {
      setUploading(false);
    }
  };

  const toggleUserAccess = (id: string) => {
    setFormData(prev => ({
      ...prev,
      allowed_account_ids: prev.allowed_account_ids.includes(id)
        ? prev.allowed_account_ids.filter((aid: string) => aid !== id)
        : [...prev.allowed_account_ids, id]
    }));
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.file_id) return Swal.fire('Peringatan', 'Harap unggah file terlebih dahulu.', 'warning');
    
    try {
      setIsSaving(true);
      if (initialData?.id) {
        await documentService.update(initialData.id, formData);
      } else {
        await documentService.create(formData);
      }
      Swal.fire({ title: 'Berhasil!', text: 'Dokumen digital telah disimpan.', icon: 'success', timer: 1500, showConfirmButton: false });
      onSuccess();
    } catch (error) {
      Swal.fire('Gagal', 'Gagal menyimpan data.', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const filteredTypes = existingTypes.filter(t => 
    t.toLowerCase().includes(formData.doc_type.toLowerCase())
  );

  const filteredAccounts = accounts.filter(acc => 
    acc.full_name.toLowerCase().includes(userSearchTerm.toLowerCase()) ||
    acc.internal_nik.toLowerCase().includes(userSearchTerm.toLowerCase())
  );

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-40 p-4">
      <div className="bg-white rounded-md shadow-2xl w-full max-w-xl overflow-hidden flex flex-col max-h-[90vh] animate-in zoom-in duration-200">
        <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between">
          <div>
            <h3 className="text-base font-bold text-[#006E62]">
              {initialData ? 'Ubah Dokumen' : 'Unggah Dokumen Baru'}
            </h3>
            <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Digital Asset Management</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={20} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="flex-1 overflow-y-auto p-6 space-y-5">
          <div className="space-y-4">
             <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">File Dokumen (G-Drive)</label>
                <div className={`p-4 border border-dashed rounded-md transition-all ${formData.file_id ? 'border-[#006E62] bg-emerald-50/20' : 'border-gray-200 bg-gray-50'}`}>
                  <label className="flex flex-col items-center justify-center cursor-pointer min-h-[80px]">
                    {uploading ? (
                      <div className="flex flex-col items-center gap-2">
                        <div className="w-6 h-6 border-2 border-[#006E62] border-t-transparent rounded-full animate-spin"></div>
                        <span className="text-[10px] font-bold text-[#006E62] uppercase">Mengunggah...</span>
                      </div>
                    ) : (
                      <>
                        <Upload className={`mb-2 ${formData.file_id ? 'text-[#006E62]' : 'text-gray-300'}`} size={24} />
                        <span className={`text-[10px] font-bold uppercase ${formData.file_id ? 'text-[#006E62]' : 'text-gray-400'}`}>
                          {formData.file_id ? 'DOKUMEN SIAP DISIMPAN' : 'Pilih File untuk Diunggah'}
                        </span>
                        <p className="text-[8px] text-gray-400 mt-1 uppercase tracking-tighter">PDF, Image, atau Doc</p>
                      </>
                    )}
                    <input type="file" className="hidden" onChange={handleFileUpload} disabled={uploading} />
                  </label>
                  {formData.file_id && (
                    <div className="mt-2 text-center">
                       <p className="text-[9px] font-mono text-gray-400 truncate tracking-tighter">ID: {formData.file_id}</p>
                    </div>
                  )}
                </div>
             </div>

             <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-1">
                  <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Nama Dokumen</label>
                  <input 
                    required
                    name="name"
                    value={formData.name}
                    onChange={handleChange}
                    className="w-full px-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none"
                    placeholder="cth: SOP Keamanan Kantor"
                  />
                </div>
                <div className="space-y-1 relative" ref={typeRef}>
                  <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Jenis Dokumen</label>
                  <div className="relative">
                    <input 
                      autoComplete="off"
                      required
                      name="doc_type"
                      value={formData.doc_type}
                      onChange={(e) => { handleChange(e); setShowTypeDropdown(true); }}
                      onFocus={() => setShowTypeDropdown(true)}
                      className="w-full px-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none pr-8"
                      placeholder="Pilih atau ketik baru..."
                    />
                    <button type="button" onClick={() => setShowTypeDropdown(!showTypeDropdown)} className="absolute right-0 top-0 bottom-0 px-2 text-gray-400">
                       <ChevronDown size={14} />
                    </button>
                  </div>
                  {showTypeDropdown && filteredTypes.length > 0 && (
                    <div className="absolute z-50 w-full mt-1 bg-white border border-gray-100 rounded shadow-xl max-h-40 overflow-y-auto">
                      {filteredTypes.map(t => (
                        <div 
                          key={t}
                          onClick={() => { setFormData(prev => ({ ...prev, doc_type: t })); setShowTypeDropdown(false); }}
                          className="px-4 py-2 text-xs hover:bg-gray-50 cursor-pointer flex items-center justify-between text-gray-700"
                        >
                          <span>{t}</span>
                          {formData.doc_type === t && <Check size={12} className="text-[#006E62]" />}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
             </div>

             <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Keterangan Singkat</label>
                <textarea 
                  name="description"
                  value={formData.description}
                  onChange={handleChange}
                  rows={2}
                  className="w-full px-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none resize-none"
                  placeholder="Penjelasan isi dokumen..."
                />
             </div>

             <div className="space-y-1 relative" ref={userRef}>
                <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest flex items-center gap-2">
                  <Users size={12} className="text-[#00FFE4]" /> Atur Hak Akses Karyawan
                </label>
                <div 
                  onClick={() => setShowUserDropdown(!showUserDropdown)}
                  className="w-full px-3 py-2 text-xs border border-gray-200 rounded flex items-center justify-between cursor-pointer bg-white"
                >
                  <span className="text-gray-400">
                    {formData.allowed_account_ids.length === 0 
                      ? '-- Pilih Karyawan yang Dapat Mengakses --' 
                      : `${formData.allowed_account_ids.length} Karyawan Terpilih`}
                  </span>
                  <ChevronDown size={14} className="text-gray-400" />
                </div>

                {showUserDropdown && (
                  <div className="absolute z-50 w-full mt-1 bg-white border border-gray-100 rounded shadow-xl max-h-60 overflow-y-auto flex flex-col p-2">
                     <div className="relative mb-2">
                       <Search className="absolute left-2 top-1/2 -translate-y-1/2 text-gray-300" size={14} />
                       <input 
                         type="text"
                         autoFocus
                         placeholder="Cari nama/nik..."
                         className="w-full pl-8 pr-3 py-1.5 text-xs bg-gray-50 border border-gray-100 rounded outline-none focus:ring-1 focus:ring-[#006E62]"
                         value={userSearchTerm}
                         onChange={(e) => setUserSearchTerm(e.target.value)}
                         onClick={(e) => e.stopPropagation()}
                       />
                     </div>
                     <div className="flex-1 overflow-y-auto space-y-1">
                       {filteredAccounts.map(acc => (
                         <div 
                           key={acc.id}
                           onClick={(e) => { e.stopPropagation(); toggleUserAccess(acc.id); }}
                           className={`flex items-center gap-2 p-2 rounded cursor-pointer text-xs hover:bg-gray-50 transition-colors ${
                             formData.allowed_account_ids.includes(acc.id) ? 'bg-emerald-50 text-[#006E62] font-bold' : 'text-gray-600'
                           }`}
                         >
                            <div className={`w-4 h-4 border rounded flex items-center justify-center shrink-0 ${
                              formData.allowed_account_ids.includes(acc.id) ? 'bg-[#006E62] border-[#006E62] text-white' : 'border-gray-300'
                            }`}>
                               {formData.allowed_account_ids.includes(acc.id) && <Check size={10} />}
                            </div>
                            <div className="flex-1 min-w-0">
                               <p className="truncate leading-none">{acc.full_name}</p>
                               <p className="text-[9px] font-normal opacity-60 uppercase mt-0.5">{acc.internal_nik} • {acc.position}</p>
                            </div>
                         </div>
                       ))}
                     </div>
                  </div>
                )}

                <div className="flex flex-wrap gap-1.5 mt-2">
                   {formData.allowed_account_ids.map((aid: string) => {
                     const acc = accounts.find(a => a.id === aid);
                     return acc ? (
                       <span key={aid} className="px-2 py-0.5 bg-emerald-50 text-[#006E62] text-[9px] font-bold rounded flex items-center gap-1 border border-emerald-100">
                         {acc.full_name}
                         <X size={10} className="cursor-pointer hover:text-red-500" onClick={() => toggleUserAccess(aid)} />
                       </span>
                     ) : null;
                   })}
                   {formData.allowed_account_ids.length === 0 && (
                     <p className="text-[9px] text-orange-500 italic font-medium">Hanya Admin yang dapat mengakses dokumen ini.</p>
                   )}
                </div>
             </div>
          </div>
        </form>

        <div className="px-6 py-4 border-t border-gray-100 flex justify-end gap-3 bg-gray-50">
          <button type="button" onClick={onClose} className="px-4 py-2 text-xs font-bold text-gray-500 uppercase">Batal</button>
          <button 
            onClick={handleSubmit}
            disabled={uploading || isSaving}
            className="flex items-center gap-2 bg-[#006E62] text-white px-8 py-2 rounded shadow-md hover:bg-[#005a50] transition-all text-xs font-bold uppercase disabled:opacity-50"
          >
            {isSaving ? <FolderKanban size={14} className="animate-spin" /> : <Save size={14} />}
            {isSaving ? 'Menyimpan...' : 'Simpan Dokumen'}
          </button>
        </div>
      </div>
    </div>
  );
};

export default DocumentForm;