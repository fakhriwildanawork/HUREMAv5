
import React, { useState } from 'react';
import { X, Save, Upload, FileBadge, Calendar, Loader2 } from 'lucide-react';
import Swal from 'sweetalert2';
import { contractService } from '../../services/contractService';
import { googleDriveService } from '../../services/googleDriveService';
import { AccountContractInput } from '../../types';

interface ContractFormModalProps {
  onClose: () => void;
  onSuccess: () => void;
  initialData?: any;
}

const ContractFormModal: React.FC<ContractFormModalProps> = ({ onClose, onSuccess, initialData }) => {
  const [formData, setFormData] = useState<any>({
    account_id: initialData?.account_id || '',
    contract_number: initialData?.contract_number || '',
    contract_type: initialData?.contract_type || 'PKWT',
    start_date: initialData?.start_date || '',
    end_date: initialData?.end_date || '',
    file_id: initialData?.file_id || '',
    notes: initialData?.notes || ''
  });

  const [uploading, setUploading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);

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
      Swal.fire('Gagal', 'Gagal mengunggah PDF.', 'error');
    } finally {
      setUploading(false);
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    try {
      setIsSaving(true);
      if (initialData?.id) {
        await contractService.update(initialData.id, formData);
      } else {
        await contractService.create(formData);
      }
      Swal.fire({ title: 'Berhasil!', text: 'Kontrak telah disimpan.', icon: 'success', timer: 1500, showConfirmButton: false });
      onSuccess();
    } catch (error) {
      Swal.fire('Gagal', 'Gagal menyimpan data.', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-40 p-4">
      <div className="bg-white rounded-md shadow-2xl w-full max-w-md overflow-hidden flex flex-col max-h-[90vh] animate-in zoom-in duration-200">
        <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between">
          <div>
            <h3 className="text-base font-bold text-[#006E62]">{initialData?.id ? 'Ubah' : 'Tambah'} Kontrak</h3>
            <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Detail Perjanjian Kerja</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={20} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-6 space-y-4 overflow-y-auto">
          <div className="space-y-1">
            <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Nomor Kontrak</label>
            <input required name="contract_number" value={formData.contract_number} onChange={handleChange} className="w-full px-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none" />
          </div>

          <div className="space-y-1">
            <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Jenis Kontrak</label>
            <select name="contract_type" value={formData.contract_type} onChange={handleChange} className="w-full px-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none">
              <option value="PKWT">PKWT</option>
              <option value="PKWTT">PKWTT (Tetap)</option>
              <option value="Magang">Magang</option>
              <option value="Harian">Harian</option>
              <option value="Addendum">Addendum</option>
            </select>
          </div>

          <div className="grid grid-cols-2 gap-4">
            <div className="space-y-1">
              <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Mulai</label>
              <input type="date" required name="start_date" value={formData.start_date} onChange={handleChange} className="w-full px-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none" />
            </div>
            <div className="space-y-1">
              <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Berakhir</label>
              <input type="date" name="end_date" value={formData.end_date} onChange={handleChange} className="w-full px-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none" />
            </div>
          </div>

          <div className="space-y-1">
            <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Unggah PDF (G-Drive)</label>
            <div className={`p-3 bg-gray-50 border border-dashed rounded-md ${formData.file_id ? 'border-[#006E62] bg-emerald-50/20' : 'border-gray-200'}`}>
              <label className="flex items-center gap-3 cursor-pointer">
                <Upload size={16} className={formData.file_id ? 'text-[#006E62]' : 'text-gray-300'} />
                <div className="flex-1">
                  <p className="text-[10px] font-bold text-gray-600 uppercase">{uploading ? 'Mengunggah...' : formData.file_id ? 'PDF TERSIMPAN' : 'PILIH PDF'}</p>
                </div>
                <input type="file" className="hidden" accept="application/pdf" onChange={handleFileUpload} disabled={uploading} />
              </label>
            </div>
          </div>

          <div className="space-y-1">
            <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Keterangan</label>
            <textarea name="notes" value={formData.notes} onChange={handleChange} rows={2} className="w-full px-3 py-2 text-xs border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] outline-none resize-none" />
          </div>
        </form>

        <div className="px-6 py-4 border-t border-gray-100 flex justify-end gap-3 bg-gray-50">
          <button type="button" onClick={onClose} className="px-4 py-2 text-xs font-bold text-gray-500 uppercase">Batal</button>
          <button onClick={handleSubmit} disabled={uploading || isSaving} className="flex items-center gap-2 bg-[#006E62] text-white px-8 py-2 rounded shadow-md hover:bg-[#005a50] transition-all text-xs font-bold uppercase disabled:opacity-50">
            {isSaving ? <Loader2 size={14} className="animate-spin" /> : <Save size={14} />} Simpan Kontrak
          </button>
        </div>
      </div>
    </div>
  );
};

export default ContractFormModal;
