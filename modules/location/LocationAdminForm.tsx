import React, { useState } from 'react';
import { X, Save, Upload, Calendar, ChevronDown, Paperclip } from 'lucide-react';
import { LocationAdminInput } from '../../types';
import { googleDriveService } from '../../services/googleDriveService';

interface LocationAdminFormProps {
  onClose: () => void;
  onSubmit: (data: any) => void;
  initialData?: any;
}

const LocationAdminForm: React.FC<LocationAdminFormProps> = ({ onClose, onSubmit, initialData }) => {
  const [formData, setFormData] = useState<any>(initialData || {
    admin_date: new Date().toISOString().split('T')[0],
    status: 'Milik Sendiri',
    due_date: '',
    description: '',
    file_ids: []
  });

  const [uploading, setUploading] = useState(false);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => {
    const { name, value } = e.target;
    
    setFormData(prev => {
      const newData = { ...prev, [name]: value };
      
      // Jika status diubah ke Milik Sendiri, kosongkan due_date di state
      if (name === 'status' && value === 'Milik Sendiri') {
        newData.due_date = '';
      }
      
      return newData;
    });
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    setUploading(true);
    const newFileIds = [...(formData.file_ids || [])];
    
    try {
      for (let i = 0; i < files.length; i++) {
        const fileId = await googleDriveService.uploadFile(files[i]);
        // Simpan dalam format id|filename agar bisa dideteksi tipenya saat dibuka
        newFileIds.push(`${fileId}|${files[i].name}`);
      }
      setFormData(prev => ({ ...prev, file_ids: newFileIds }));
    } catch (error) {
      alert('Gagal mengunggah file.');
    } finally {
      setUploading(false);
    }
  };

  const removeFile = (idx: number) => {
    setFormData(prev => ({
      ...prev,
      file_ids: prev.file_ids.filter((_, i) => i !== idx)
    }));
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    // Sanitasi data string kosong ke null sekarang dilakukan di locationService.ts
    onSubmit(formData);
  };

  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black bg-opacity-40 p-4">
      <div className="bg-white rounded-md shadow-2xl w-full max-w-lg overflow-hidden flex flex-col animate-in zoom-in duration-200">
        <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between">
          <div>
            <h3 className="text-base font-bold text-[#006E62]">{initialData ? 'Edit' : 'Tambah'} Data Administrasi</h3>
            <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Legalitas & Status Aset</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={20} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-6 space-y-5">
          <div className="grid grid-cols-2 gap-4">
            <div className="space-y-1">
              <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Tanggal</label>
              <div className="relative">
                <input
                  type="date"
                  required
                  name="admin_date"
                  value={formData.admin_date}
                  onChange={handleChange}
                  className="w-full px-3 py-1.5 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] focus:border-[#006E62] outline-none bg-gray-50"
                />
              </div>
            </div>
            <div className="space-y-1">
              <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Status Pengelolaan</label>
              <div className="relative">
                <select
                  name="status"
                  value={formData.status}
                  onChange={handleChange}
                  className="w-full px-3 py-1.5 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] focus:border-[#006E62] outline-none bg-gray-50 appearance-none"
                >
                  <option value="Milik Sendiri">Milik Sendiri</option>
                  <option value="Sewa/Kontrak">Sewa/Kontrak</option>
                  <option value="Kerjasama">Kerjasama</option>
                </select>
                <ChevronDown className="absolute right-2 top-1/2 -translate-y-1/2 text-gray-400 pointer-events-none" size={14} />
              </div>
            </div>
          </div>

          <div className="space-y-1">
            <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Jatuh Tempo (PKS/Kontrak)</label>
            <input
              type="date"
              required={formData.status !== 'Milik Sendiri'}
              disabled={formData.status === 'Milik Sendiri'}
              name="due_date"
              value={formData.due_date}
              onChange={handleChange}
              className={`w-full px-3 py-1.5 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] focus:border-[#006E62] outline-none transition-colors ${
                formData.status === 'Milik Sendiri' 
                ? 'bg-gray-100 cursor-not-allowed text-gray-400' 
                : 'bg-white'
              }`}
            />
          </div>

          <div className="space-y-1">
            <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Keterangan / Catatan</label>
            <textarea
              name="description"
              value={formData.description}
              onChange={handleChange}
              rows={3}
              placeholder="cth: Perjanjian sewa nomor 123/ABC/2024..."
              className="w-full px-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] focus:border-[#006E62] outline-none bg-gray-50 resize-none"
            />
          </div>

          <div className="space-y-2">
            <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">File Dokumentasi (Multiple)</label>
            <div className="border border-dashed border-gray-200 rounded-md p-4 bg-gray-50">
              <div className="flex flex-wrap gap-2 mb-3">
                {formData.file_ids.map((fid: string, i: number) => (
                  <div key={i} className="flex items-center gap-2 bg-white border border-gray-200 px-2 py-1 rounded-md text-[10px] font-bold group">
                    <Paperclip size={10} className="text-[#006E62]" />
                    <span className="truncate max-w-[80px]">FILE {i + 1}</span>
                    <button type="button" onClick={() => removeFile(i)} className="text-gray-300 hover:text-red-500">
                      <X size={10} />
                    </button>
                  </div>
                ))}
              </div>
              <label className={`flex flex-col items-center justify-center p-4 bg-white border border-gray-100 rounded cursor-pointer hover:border-[#006E62] transition-colors ${uploading ? 'opacity-50 cursor-not-allowed' : ''}`}>
                <Upload className="text-[#006E62] mb-1" size={16} />
                <span className="text-[10px] font-bold uppercase text-gray-400">{uploading ? 'SEDANG MENGUNGGAH...' : 'KLIK UNTUK UNGGAH FILE PKS / DOKUMEN'}</span>
                <input type="file" multiple className="hidden" onChange={handleFileUpload} disabled={uploading} />
              </label>
            </div>
          </div>

          <div className="flex justify-end gap-3 pt-4">
            <button type="button" onClick={onClose} className="px-4 py-2 text-xs font-bold text-gray-500 uppercase">
              Batal
            </button>
            <button
              type="submit"
              disabled={uploading}
              className="flex items-center gap-2 bg-[#006E62] text-white px-6 py-2 rounded-md hover:bg-[#005a50] transition-all shadow-md text-xs font-bold uppercase disabled:opacity-50"
            >
              <Save size={14} /> Simpan Administrasi
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};

export default LocationAdminForm;