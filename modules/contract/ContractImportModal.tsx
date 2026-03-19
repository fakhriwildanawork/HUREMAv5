
import React, { useState } from 'react';
import { X, FileUp, Download, CheckCircle, AlertTriangle, Save, Loader2, Paperclip } from 'lucide-react';
import Swal from 'sweetalert2';
import { contractService } from '../../services/contractService';
import { googleDriveService } from '../../services/googleDriveService';

interface ContractImportModalProps {
  onClose: () => void;
  onSuccess: () => void;
}

const ContractImportModal: React.FC<ContractImportModalProps> = ({ onClose, onSuccess }) => {
  const [step, setStep] = useState<1 | 2>(1);
  const [previewData, setPreviewData] = useState<any[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [isUploadingAttachments, setIsUploadingAttachments] = useState(false);
  const [bulkFiles, setBulkFiles] = useState<Record<string, string>>({});

  const handleBulkFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    try {
      setIsUploadingAttachments(true);
      const mapping: Record<string, string> = {};
      
      // Upload files in parallel
      const uploadPromises = Array.from(files).map(async (f) => {
        const file = f as File;
        const fileId = await googleDriveService.uploadFile(file);
        // Store mapping: filename (without extension) -> fileId
        const fileName = file.name.split('.').slice(0, -1).join('.');
        mapping[fileName] = fileId;
      });

      await Promise.all(uploadPromises);
      setBulkFiles(prev => ({ ...prev, ...mapping }));
      
      Swal.fire({
        title: 'Berhasil!',
        text: `${files.length} lampiran berhasil diunggah dan siap dipasangkan.`,
        icon: 'success',
        timer: 1500,
        showConfirmButton: false
      });
    } catch (error) {
      Swal.fire('Gagal', 'Gagal mengunggah beberapa lampiran.', 'error');
    } finally {
      setIsUploadingAttachments(false);
      if (e.target) e.target.value = '';
    }
  };

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      setIsProcessing(true);
      const results = await contractService.processImport(file, bulkFiles) as any[];
      setPreviewData(results);
      setStep(2);
    } catch (error) {
      Swal.fire('Gagal', 'Format file tidak didukung atau rusak.', 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const handleCommit = async () => {
    const validCount = previewData.filter(d => d.isValid).length;
    if (validCount === 0) {
      Swal.fire('Peringatan', 'Tidak ada data valid untuk diimpor.', 'warning');
      return;
    }

    const confirm = await Swal.fire({
      title: 'Konfirmasi Impor',
      text: `Sistem akan memproses ${validCount} baris perpanjangan kontrak. Profil utama karyawan akan otomatis diperbarui mengikuti tanggal terbaru. Lanjutkan?`,
      icon: 'question',
      showCancelButton: true,
      confirmButtonColor: '#006E62',
      confirmButtonText: 'Ya, Proses Sekarang'
    });

    if (confirm.isConfirmed) {
      try {
        setIsUploading(true);
        await contractService.commitImport(previewData);
        Swal.fire('Berhasil!', 'Seluruh data kontrak telah diperbarui.', 'success');
        onSuccess();
      } catch (error) {
        Swal.fire('Gagal', 'Terjadi kesalahan saat memproses data.', 'error');
      } finally {
        setIsUploading(false);
      }
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-40 p-4">
      <div className="bg-white rounded-md shadow-2xl w-full max-w-4xl overflow-hidden flex flex-col max-h-[90vh] animate-in zoom-in duration-200">
        <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between">
          <div>
            <h3 className="text-base font-bold text-[#006E62]">Impor Massal Perpanjangan Kontrak</h3>
            <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Tahap {step}: {step === 1 ? 'Unggah File' : 'Pratinjau Data'}</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={20} />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-6">
          {step === 1 ? (
            <div className="space-y-6 flex flex-col items-center py-10">
              <div className="w-16 h-16 bg-emerald-50 text-[#006E62] rounded-full flex items-center justify-center mb-4">
                <FileUp size={32} />
              </div>
              <div className="text-center max-w-md">
                <h4 className="text-lg font-bold text-gray-800">Unggah Excel Kontrak</h4>
                <p className="text-xs text-gray-500 mt-2">Gunakan template resmi HUREMA. Sistem akan mencatat riwayat kontrak baru dan memperbarui masa berlaku di profil akun secara otomatis.</p>
              </div>

              <div className="flex flex-col gap-3 w-full max-w-xs">
                <button 
                  onClick={() => contractService.downloadTemplate()}
                  className="flex items-center justify-center gap-2 border border-gray-200 px-4 py-3 rounded-md hover:bg-gray-50 transition-colors text-sm font-bold text-gray-600 uppercase tracking-tighter"
                >
                  <Download size={18} /> 1. Download Template
                </button>
                
                <label className="flex items-center justify-center gap-2 border border-gray-200 px-4 py-3 rounded-md hover:bg-gray-50 transition-colors text-sm font-bold text-gray-600 uppercase tracking-tighter cursor-pointer">
                  {isUploadingAttachments ? <Loader2 size={18} className="animate-spin" /> : <Paperclip size={18} />}
                  {isUploadingAttachments ? 'Mengunggah...' : `2. Unggah Lampiran PDF (${Object.keys(bulkFiles).length} File)`}
                  <input type="file" className="hidden" accept="application/pdf" multiple onChange={handleBulkFileUpload} disabled={isUploadingAttachments} />
                </label>

                <label className="flex items-center justify-center gap-2 bg-[#006E62] text-white px-4 py-3 rounded-md hover:bg-[#005a50] transition-colors shadow-md text-sm font-bold uppercase tracking-tighter cursor-pointer">
                  {isProcessing ? <Loader2 size={18} className="animate-spin" /> : <FileUp size={18} />}
                  {isProcessing ? 'Memproses...' : '3. Unggah Excel Terisi'}
                  <input type="file" className="hidden" accept=".xlsx" onChange={handleFileChange} disabled={isProcessing} />
                </label>
              </div>
            </div>
          ) : (
            <div className="space-y-4">
              <div className="flex items-center justify-between bg-emerald-50 p-4 rounded-md border border-emerald-100">
                <div className="flex items-center gap-2 text-emerald-700">
                  <CheckCircle size={20} />
                  <p className="text-xs font-bold">Terbaca {previewData.length} baris. ({previewData.filter(d => d.isValid).length} Valid, {previewData.filter(d => !d.isValid).length} Error)</p>
                </div>
                <button onClick={() => setStep(1)} className="text-[10px] font-bold uppercase text-[#006E62] hover:underline">Ganti File</button>
              </div>

              <div className="border border-gray-100 rounded overflow-x-auto">
                <table className="w-full text-left text-[11px]">
                  <thead className="bg-gray-50 font-bold text-gray-500 uppercase">
                    <tr>
                      <th className="px-4 py-2">Status</th>
                      <th className="px-4 py-2">Nama Karyawan</th>
                      <th className="px-4 py-2">No Kontrak</th>
                      <th className="px-4 py-2">Jenis</th>
                      <th className="px-4 py-2">Mulai</th>
                      <th className="px-4 py-2">Akhir</th>
                      <th className="px-4 py-2">Lampiran</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-50">
                    {previewData.map((row, idx) => (
                      <tr key={idx} className={row.isValid ? '' : 'bg-red-50'}>
                        <td className="px-4 py-2">
                          {row.isValid ? (
                            <CheckCircle size={14} className="text-emerald-500" />
                          ) : (
                            <AlertTriangle size={14} className="text-red-500" />
                          )}
                        </td>
                        <td className="px-4 py-2 font-bold">{row.full_name}</td>
                        <td className="px-4 py-2">{row.contract_number}</td>
                        <td className="px-4 py-2">{row.contract_type}</td>
                        <td className="px-4 py-2">{row.start_date}</td>
                        <td className="px-4 py-2">{row.end_date || 'TETAP'}</td>
                        <td className="px-4 py-2">
                          {row.file_id ? (
                            <span className="flex items-center gap-1 text-emerald-600 font-bold"><Paperclip size={12} /> OK</span>
                          ) : (
                            <span className="text-gray-400 italic">Tidak ada</span>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>

        <div className="px-6 py-4 border-t border-gray-100 flex justify-end gap-3 bg-gray-50">
          <button type="button" onClick={onClose} className="px-4 py-2 text-xs font-bold text-gray-500 uppercase">Batal</button>
          {step === 2 && (
            <button 
              onClick={handleCommit}
              disabled={isUploading}
              className="flex items-center gap-2 bg-[#006E62] text-white px-8 py-2 rounded shadow-md hover:bg-[#005a50] transition-all text-xs font-bold uppercase disabled:opacity-50"
            >
              {isUploading ? <Loader2 size={14} className="animate-spin" /> : <Save size={14} />}
              {isUploading ? 'Sedang Memproses...' : 'Simpan Seluruh Data'}
            </button>
          )}
        </div>
      </div>
    </div>
  );
};

export default ContractImportModal;
