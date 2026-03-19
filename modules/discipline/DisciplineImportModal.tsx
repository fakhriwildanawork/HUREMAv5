import React, { useState } from 'react';
import { X, FileUp, Download, CheckCircle, AlertTriangle, Save, Loader2 } from 'lucide-react';
import Swal from 'sweetalert2';
import { disciplineService } from '../../services/disciplineService';

interface DisciplineImportModalProps {
  onClose: () => void;
  onSuccess: () => void;
  type?: 'warning' | 'termination';
}

const DisciplineImportModal: React.FC<DisciplineImportModalProps> = ({ onClose, onSuccess, type = 'warning' }) => {
  const [step, setStep] = useState<1 | 2>(1);
  const [previewData, setPreviewData] = useState<any[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isUploading, setIsUploading] = useState(false);

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      setIsProcessing(true);
      const results = type === 'warning' 
        ? await disciplineService.processWarningImport(file) as any[]
        : await disciplineService.processTerminationImport(file) as any[];
      setPreviewData(results);
      setStep(2);
    } catch (error) {
      Swal.fire('Gagal', 'Format file tidak didukung atau kolom tidak sesuai.', 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const handleCommit = async () => {
    const validCount = previewData.filter(d => d.isValid).length;
    if (validCount === 0) return Swal.fire('Peringatan', 'Tidak ada data valid untuk diimpor.', 'warning');

    const confirm = await Swal.fire({
      title: 'Konfirmasi Impor',
      text: `Sistem akan memproses ${validCount} baris data ${type === 'warning' ? 'peringatan' : 'pemberhentian'}. Lanjutkan?`,
      icon: 'question',
      showCancelButton: true,
      confirmButtonColor: type === 'warning' ? '#006E62' : '#dc2626',
      confirmButtonText: 'Ya, Proses Sekarang'
    });

    if (confirm.isConfirmed) {
      try {
        setIsUploading(true);
        if (type === 'warning') {
          await disciplineService.commitWarningImport(previewData);
        } else {
          await disciplineService.commitTerminationImport(previewData);
        }
        Swal.fire('Berhasil!', `Seluruh data ${type === 'warning' ? 'SP' : 'Exit'} telah diperbarui.`, 'success');
        onSuccess();
      } catch (error) {
        Swal.fire('Gagal', 'Terjadi kesalahan saat memproses data ke database.', 'error');
      } finally {
        setIsUploading(false);
      }
    }
  };

  const getTemplateDownloadAction = () => {
    if (type === 'warning') return disciplineService.downloadWarningTemplate();
    return disciplineService.downloadTerminationTemplate();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-40 p-4">
      <div className="bg-white rounded-md shadow-2xl w-full max-w-4xl overflow-hidden flex flex-col max-h-[90vh] animate-in zoom-in duration-200">
        <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between">
          <div>
            <h3 className={`text-base font-bold ${type === 'warning' ? 'text-[#006E62]' : 'text-red-600'}`}>
              Impor Massal {type === 'warning' ? 'Riwayat SP' : 'Karyawan Keluar (Exit)'}
            </h3>
            <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Tahap {step}: {step === 1 ? 'Unggah File' : 'Pratinjau Data'}</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors"><X size={20} /></button>
        </div>

        <div className="flex-1 overflow-y-auto p-6">
          {step === 1 ? (
            <div className="space-y-6 flex flex-col items-center py-10 text-center">
              <FileUp size={48} className={type === 'warning' ? 'text-[#006E62]' : 'text-red-600'} />
              <div className="max-w-md">
                <h4 className="text-lg font-bold text-gray-800">
                  Unggah Excel {type === 'warning' ? 'Riwayat SP' : 'Pemberhentian Karyawan'}
                </h4>
                <p className="text-xs text-gray-500 mt-2">
                  Gunakan template untuk mencatat data {type === 'warning' ? 'Surat Peringatan' : 'Exit/Resign'} secara massal.
                </p>
              </div>
              <div className="flex flex-col gap-3 w-full max-w-xs">
                <button onClick={getTemplateDownloadAction} className="flex items-center justify-center gap-2 border border-gray-200 px-4 py-3 rounded-md hover:bg-gray-50 text-xs font-bold text-gray-600 uppercase">
                  <Download size={18} /> 1. Download Template
                </button>
                <label className={`flex items-center justify-center gap-2 text-white px-4 py-3 rounded-md transition-colors shadow-md text-xs font-bold uppercase cursor-pointer ${type === 'warning' ? 'bg-[#006E62] hover:bg-[#005a50]' : 'bg-red-600 hover:bg-red-700'}`}>
                  {isProcessing ? <Loader2 size={18} className="animate-spin" /> : <FileUp size={18} />}
                  {isProcessing ? 'Memproses...' : '2. Unggah Excel'}
                  <input type="file" className="hidden" accept=".xlsx" onChange={handleFileChange} disabled={isProcessing} />
                </label>
              </div>
            </div>
          ) : (
            <div className="space-y-4">
              <div className={`flex items-center justify-between p-4 rounded border text-xs font-bold ${type === 'warning' ? 'bg-emerald-50 border-emerald-100 text-emerald-700' : 'bg-red-50 border-red-100 text-red-700'}`}>
                <div className="flex items-center gap-2">
                  <CheckCircle size={20} /> Terbaca {previewData.length} baris. ({previewData.filter(d => d.isValid).length} Valid)
                </div>
                <button onClick={() => setStep(1)} className="uppercase hover:underline">Ganti File</button>
              </div>
              <div className="border border-gray-100 rounded overflow-x-auto text-[11px]">
                <table className="w-full text-left">
                  <thead className="bg-gray-50 font-bold text-gray-500 uppercase">
                    <tr>
                      <th className="px-4 py-2">Status</th>
                      <th className="px-4 py-2">Nama</th>
                      <th className="px-4 py-2">{type === 'warning' ? 'Tipe SP' : 'Tipe Exit'}</th>
                      <th className="px-4 py-2">Alasan</th>
                      <th className="px-4 py-2">{type === 'warning' ? 'Tgl SP' : 'Tgl Exit'}</th>
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
                        <td className={`px-4 py-2 uppercase font-bold ${type === 'warning' ? 'text-orange-600' : 'text-red-600'}`}>
                          {type === 'warning' ? row.warning_type : row.termination_type}
                        </td>
                        <td className="px-4 py-2 truncate max-w-xs">{row.reason}</td>
                        <td className="px-4 py-2">{type === 'warning' ? row.issue_date : row.termination_date}</td>
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
              className={`flex items-center gap-2 text-white px-8 py-2 rounded shadow-md transition-all text-xs font-bold uppercase disabled:opacity-50 ${type === 'warning' ? 'bg-[#006E62] hover:bg-[#005a50]' : 'bg-red-600 hover:bg-red-700'}`}
            >
              {isUploading ? <Loader2 size={14} className="animate-spin" /> : <Save size={14} />} 
              {isUploading ? 'Simpan Data...' : 'Simpan Seluruh Data'}
            </button>
          )}
        </div>
      </div>
    </div>
  );
};

export default DisciplineImportModal;