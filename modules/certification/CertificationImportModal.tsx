
import React, { useState } from 'react';
import { X, FileUp, Download, CheckCircle, AlertTriangle, Save, Loader2 } from 'lucide-react';
import Swal from 'sweetalert2';
import { certificationService } from '../../services/certificationService';

interface CertificationImportModalProps {
  onClose: () => void;
  onSuccess: () => void;
}

const CertificationImportModal: React.FC<CertificationImportModalProps> = ({ onClose, onSuccess }) => {
  const [step, setStep] = useState<1 | 2>(1);
  const [previewData, setPreviewData] = useState<any[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isUploading, setIsUploading] = useState(false);

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      setIsProcessing(true);
      const results = await certificationService.processImport(file) as any[];
      setPreviewData(results);
      setStep(2);
    } catch (error) {
      Swal.fire('Gagal', 'Format file tidak didukung.', 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const handleCommit = async () => {
    const validCount = previewData.filter(d => d.isValid).length;
    if (validCount === 0) return Swal.fire('Peringatan', 'Tidak ada data valid.', 'warning');

    const confirm = await Swal.fire({
      title: 'Konfirmasi Impor',
      text: `Impor ${validCount} sertifikasi?`,
      icon: 'question',
      showCancelButton: true,
      confirmButtonColor: '#006E62',
      confirmButtonText: 'Ya, Impor'
    });

    if (confirm.isConfirmed) {
      try {
        setIsUploading(true);
        await certificationService.commitImport(previewData);
        Swal.fire('Berhasil!', 'Data sertifikasi telah diimpor.', 'success');
        onSuccess();
      } catch (error) {
        Swal.fire('Gagal', 'Terjadi kesalahan.', 'error');
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
            <h3 className="text-base font-bold text-[#006E62]">Impor Massal Sertifikasi</h3>
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
                <h4 className="text-lg font-bold text-gray-800">Unggah Excel Sertifikasi</h4>
                <p className="text-xs text-gray-500 mt-2">Gunakan template resmi untuk mengunggah data sertifikat karyawan secara massal.</p>
              </div>
              <div className="flex flex-col gap-3 w-full max-w-xs">
                <button onClick={() => certificationService.downloadTemplate()} className="flex items-center justify-center gap-2 border border-gray-200 px-4 py-3 rounded-md hover:bg-gray-50 transition-colors text-sm font-bold text-gray-600 uppercase">
                  <Download size={18} /> 1. Download Template
                </button>
                <label className="flex items-center justify-center gap-2 bg-[#006E62] text-white px-4 py-3 rounded-md hover:bg-[#005a50] transition-colors shadow-md text-sm font-bold uppercase cursor-pointer">
                  {isProcessing ? <Loader2 size={18} className="animate-spin" /> : <FileUp size={18} />}
                  {isProcessing ? 'Memproses...' : '2. Unggah Excel'}
                  <input type="file" className="hidden" accept=".xlsx" onChange={handleFileChange} disabled={isProcessing} />
                </label>
              </div>
            </div>
          ) : (
            <div className="space-y-4">
              <div className="flex items-center justify-between bg-emerald-50 p-4 rounded-md border border-emerald-100">
                <div className="flex items-center gap-2 text-emerald-700 text-xs font-bold">
                  <CheckCircle size={20} /> Terbaca {previewData.length} baris. ({previewData.filter(d => d.isValid).length} Valid)
                </div>
                <button onClick={() => setStep(1)} className="text-[10px] font-bold uppercase text-[#006E62]">Ganti File</button>
              </div>
              <div className="border border-gray-100 rounded overflow-x-auto">
                <table className="w-full text-left text-[11px]">
                  <thead className="bg-gray-50 font-bold text-gray-500 uppercase">
                    <tr><th className="px-4 py-2">Status</th><th className="px-4 py-2">Nama</th><th className="px-4 py-2">Jenis</th><th className="px-4 py-2">Sertifikat</th><th className="px-4 py-2">Tgl</th></tr>
                  </thead>
                  <tbody className="divide-y divide-gray-50">
                    {previewData.map((row, idx) => (
                      <tr key={idx} className={row.isValid ? '' : 'bg-red-50'}>
                        <td className="px-4 py-2">{row.isValid ? <CheckCircle size={14} className="text-emerald-500" /> : <AlertTriangle size={14} className="text-red-500" />}</td>
                        <td className="px-4 py-2 font-bold">{row.full_name}</td>
                        <td className="px-4 py-2">{row.cert_type}</td>
                        <td className="px-4 py-2">{row.cert_name}</td>
                        <td className="px-4 py-2">{row.cert_date}</td>
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
            <button onClick={handleCommit} disabled={isUploading} className="flex items-center gap-2 bg-[#006E62] text-white px-8 py-2 rounded shadow-md hover:bg-[#005a50] transition-all text-xs font-bold uppercase disabled:opacity-50">
              {isUploading ? <Loader2 size={14} className="animate-spin" /> : <Save size={14} />} {isUploading ? 'Memproses...' : 'Simpan Data'}
            </button>
          )}
        </div>
      </div>
    </div>
  );
};

export default CertificationImportModal;
