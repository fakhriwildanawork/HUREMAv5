import React, { useState } from 'react';
import { X, FileUp, Download, CheckCircle, AlertTriangle, Save, Loader2, Paperclip, Upload, User, FileBadge } from 'lucide-react';
import Swal from 'sweetalert2';
import { accountService } from '../../services/accountService';
import { googleDriveService } from '../../services/googleDriveService';

interface AccountImportModalProps {
  onClose: () => void;
  onSuccess: () => void;
}

const AccountImportModal: React.FC<AccountImportModalProps> = ({ onClose, onSuccess }) => {
  const [step, setStep] = useState<1 | 2 | 3>(1);
  const [previewData, setPreviewData] = useState<any[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [isUploadingAttachments, setIsUploadingAttachments] = useState(false);
  const [bulkFiles, setBulkFiles] = useState<Record<string, string>>({});
  const [fileList, setFileList] = useState<{ name: string; id: string }[]>([]);

  const handleBulkFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    try {
      setIsUploadingAttachments(true);
      const newFiles: { name: string; id: string }[] = [];
      
      const uploadPromises = Array.from(files).map(async (f) => {
        const file = f as File;
        const fileId = await googleDriveService.uploadFile(file);
        const fileName = file.name.split('.').slice(0, -1).join('.');
        newFiles.push({ name: fileName, id: fileId });
      });

      await Promise.all(uploadPromises);
      
      const updatedFileList = [...fileList, ...newFiles];
      setFileList(updatedFileList);
      
      const mapping: Record<string, string> = {};
      updatedFileList.forEach(f => mapping[f.name] = f.id);
      setBulkFiles(mapping);

      // Update preview data with new attachments
      setPreviewData(prev => prev.map(row => {
        const internalNik = row.internal_nik || '';
        if (!internalNik) return row;
        
        const normalizedNik = String(internalNik).replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
        
        const updates: any = {};
        updatedFileList.forEach(f => {
          const normalizedFileName = f.name.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
          if (normalizedFileName.includes(normalizedNik)) {
            if (normalizedFileName.includes('photo')) updates.photo_google_id = f.id;
            else if (normalizedFileName.includes('ktp')) updates.ktp_google_id = f.id;
            else if (normalizedFileName.includes('sk')) updates.file_sk_id = f.id;
            else if (normalizedFileName.includes('contract')) {
              updates.contract_initial = { ...row.contract_initial, file_id: f.id };
            }
            else if (normalizedFileName.includes('diploma')) updates.diploma_google_id = f.id;
            else if (normalizedFileName === normalizedNik) updates.photo_google_id = f.id;
          }
        });

        return {
          ...row,
          ...updates
        };
      }));
      
      Swal.fire({
        title: 'Berhasil!',
        text: `${files.length} foto berhasil ditambahkan.`,
        icon: 'success',
        timer: 1500,
        showConfirmButton: false
      });
    } catch (error) {
      Swal.fire('Gagal', 'Gagal mengunggah beberapa foto.', 'error');
    } finally {
      setIsUploadingAttachments(false);
      if (e.target) e.target.value = '';
    }
  };

  const handleDeleteFile = (fileName: string) => {
    const updatedFileList = fileList.filter(f => f.name !== fileName);
    setFileList(updatedFileList);
    
    const mapping: Record<string, string> = {};
    updatedFileList.forEach(f => mapping[f.name] = f.id);
    setBulkFiles(mapping);

    // Update preview data (remove matched ID if file deleted)
    setPreviewData(prev => prev.map(row => {
      const internalNik = row.internal_nik || '';
      if (!internalNik) return row;
      
      const normalizedNik = String(internalNik).replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
      
      const updates: any = {
        photo_google_id: null,
        ktp_google_id: null,
        file_sk_id: null,
        diploma_google_id: null
      };

      updatedFileList.forEach(f => {
        const normalizedFileName = f.name.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
        if (normalizedFileName.includes(normalizedNik)) {
          if (normalizedFileName.includes('photo')) updates.photo_google_id = f.id;
          else if (normalizedFileName.includes('ktp')) updates.ktp_google_id = f.id;
          else if (normalizedFileName.includes('sk')) updates.file_sk_id = f.id;
          else if (normalizedFileName.includes('contract')) {
            updates.contract_initial = { ...row.contract_initial, file_id: f.id };
          }
          else if (normalizedFileName.includes('diploma')) updates.diploma_google_id = f.id;
          else if (normalizedFileName === normalizedNik) updates.photo_google_id = f.id;
        }
      });

      if (updates.contract_initial === undefined) {
        updates.contract_initial = { ...row.contract_initial, file_id: null };
      }

      return {
        ...row,
        ...updates
      };
    }));
  };

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      setIsProcessing(true);
      const results = await accountService.processImport(file, bulkFiles) as any[];
      setPreviewData(results);
    } catch (error) {
      Swal.fire('Gagal', 'Format file tidak didukung atau rusak.', 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const handleCommit = async () => {
    const hasError = previewData.some(d => !d.isValid);
    if (hasError) {
      Swal.fire('Peringatan', 'Masih ada data yang error. Silakan perbaiki file Excel Anda terlebih dahulu.', 'warning');
      return;
    }

    const validCount = previewData.length;
    if (validCount === 0) {
      Swal.fire('Peringatan', 'Tidak ada data untuk diimpor.', 'warning');
      return;
    }

    const confirm = await Swal.fire({
      title: 'Konfirmasi Impor',
      text: `Sistem akan membuat ${validCount} akun karyawan baru beserta log karir, kesehatan, dan kontrak awalnya. Lanjutkan?`,
      icon: 'question',
      showCancelButton: true,
      confirmButtonColor: '#006E62',
      confirmButtonText: 'Ya, Proses Sekarang'
    });

    if (confirm.isConfirmed) {
      try {
        setIsUploading(true);
        const results = await accountService.commitImport(previewData);
        
        if (results.failed > 0) {
          Swal.fire({
            title: 'Selesai dengan Catatan',
            html: `<p>${results.success} akun berhasil, ${results.failed} gagal.</p><div class="mt-2 text-left text-[10px] bg-red-50 p-2 max-h-40 overflow-y-auto">${results.errors.join('<br/>')}</div>`,
            icon: 'warning'
          });
        } else {
          Swal.fire('Berhasil!', `${results.success} akun karyawan baru telah dibuat.`, 'success');
        }
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
      <div className="bg-white rounded-md shadow-2xl w-full max-w-5xl overflow-hidden flex flex-col max-h-[90vh] animate-in zoom-in duration-200">
        <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between">
          <div>
            <h3 className="text-base font-bold text-[#006E62]">Impor Massal Akun Karyawan</h3>
            <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Tahap {step}: {step === 1 ? 'Unggah File & Pratinjau' : step === 2 ? 'Unggah Foto Profil' : 'Unggah Dokumen Pendukung'}</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={20} />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-6">
          {step === 1 ? (
            <div className="space-y-6">
              <div className="flex flex-col items-center py-6 border-b border-gray-50">
                <div className="w-16 h-16 bg-emerald-50 text-[#006E62] rounded-full flex items-center justify-center mb-4">
                  <FileUp size={32} />
                </div>
                <div className="text-center max-w-md">
                  <h4 className="text-lg font-bold text-gray-800">1. Unggah Excel Akun</h4>
                </div>

                <div className="flex items-center gap-3 mt-6 w-full max-w-md">
                  <button 
                    onClick={() => accountService.downloadTemplate()}
                    className="flex-1 flex items-center justify-center gap-2 border border-gray-200 px-4 py-3 rounded-md hover:bg-gray-50 transition-colors text-sm font-bold text-gray-600 uppercase tracking-tighter"
                  >
                    <Download size={18} /> Download Template
                  </button>
                  
                  <label className="flex-1 flex items-center justify-center gap-2 bg-[#006E62] text-white px-4 py-3 rounded-md hover:bg-[#005a50] transition-colors shadow-md text-sm font-bold uppercase tracking-tighter cursor-pointer">
                    {isProcessing ? <Loader2 size={18} className="animate-spin" /> : <FileUp size={18} />}
                    {isProcessing ? 'Memproses...' : previewData.length > 0 ? 'Ganti Excel' : 'Unggah Excel'}
                    <input 
                      type="file" 
                      className="hidden" 
                      accept=".xlsx" 
                      onChange={handleFileChange} 
                      onClick={(e) => (e.target as HTMLInputElement).value = ''}
                      disabled={isProcessing} 
                    />
                  </label>
                </div>
              </div>

              {previewData.length > 0 && (
                <div className="space-y-4 animate-in fade-in slide-in-from-top-4 duration-300">
                  <div className="flex items-center justify-between bg-emerald-50 p-4 rounded-md border border-emerald-100">
                    <div className="flex items-center gap-2 text-emerald-700">
                      <CheckCircle size={20} />
                      <p className="text-xs font-bold">Terbaca {previewData.length} baris. ({previewData.filter(d => d.isValid).length} Valid, <span className={previewData.some(d => !d.isValid) ? 'text-red-600' : ''}>{previewData.filter(d => !d.isValid).length} Error</span>)</p>
                    </div>
                  </div>

                  <div className="border border-gray-100 rounded overflow-x-auto">
                    <table className="w-full text-left text-[10px]">
                      <thead className="bg-gray-50 font-bold text-gray-500 uppercase">
                        <tr>
                          <th className="px-4 py-2">Status</th>
                          <th className="px-4 py-2">Nama Lengkap</th>
                          <th className="px-4 py-2">NIK Internal</th>
                          <th className="px-4 py-2">Jabatan</th>
                          <th className="px-4 py-2">Departemen</th>
                          <th className="px-4 py-2">Lokasi</th>
                          <th className="px-4 py-2">Tgl Mulai</th>
                          <th className="px-4 py-2">Keterangan</th>
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
                            <td className="px-4 py-2 font-mono">{row.internal_nik}</td>
                            <td className="px-4 py-2">{row.position}</td>
                            <td className="px-4 py-2">{row.grade}</td>
                            <td className="px-4 py-2">{row.location_name}</td>
                            <td className="px-4 py-2">{row.start_date}</td>
                            <td className="px-4 py-2">
                              {!row.isValid && (
                                <span className="text-red-600 font-medium">
                                  {row.errorMsg || 'Data wajib belum lengkap'}
                                </span>
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
          ) : step === 2 ? (
            <div className="space-y-6">
              <div className="flex flex-col items-center py-6 border-b border-gray-50">
                <div className="w-16 h-16 bg-emerald-50 text-[#006E62] rounded-full flex items-center justify-center mb-4">
                  <User size={32} />
                </div>
                <div className="text-center max-w-md">
                  <h4 className="text-lg font-bold text-gray-800">2. Unggah Foto Profil (Opsional)</h4>
                  <p className="text-xs text-gray-500 mt-2">Unggah file foto karyawan. Sistem akan mencocokkan nama file dengan NIK Internal di Excel secara otomatis.</p>
                </div>

                <div className="mt-6 w-full max-w-md">
                  <label className="flex items-center justify-center gap-2 border-2 border-dashed border-gray-200 px-4 py-6 rounded-md hover:bg-gray-50 hover:border-[#006E62] transition-all text-sm font-bold text-gray-400 uppercase tracking-tighter cursor-pointer">
                    {isUploadingAttachments ? <Loader2 size={24} className="animate-spin text-[#006E62]" /> : <Upload size={24} />}
                    {isUploadingAttachments ? 'Sedang Mengunggah...' : 'Klik atau Seret Foto ke Sini'}
                    <input type="file" className="hidden" accept="image/*" multiple onChange={handleBulkFileUpload} disabled={isUploadingAttachments} />
                  </label>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                <div className="md:col-span-1 space-y-3">
                  <h5 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Foto Terunggah ({fileList.filter(f => f.name.toLowerCase().includes('photo') || !f.name.includes('_')).length})</h5>
                  <div className="bg-gray-50 rounded-md border border-gray-100 p-2 max-h-[300px] overflow-y-auto space-y-1">
                    {fileList.length === 0 ? (
                      <p className="text-[10px] text-gray-400 text-center py-4 italic">Belum ada foto</p>
                    ) : (
                      fileList.filter(f => f.name.toLowerCase().includes('photo') || !f.name.includes('_')).map((file, idx) => (
                        <div key={idx} className="flex items-center justify-between bg-white p-2 rounded border border-gray-100 group">
                          <div className="flex items-center gap-2 min-w-0">
                            <Paperclip size={12} className="text-[#006E62] shrink-0" />
                            <span className="text-[10px] font-medium text-gray-600 truncate">{file.name}</span>
                          </div>
                          <button 
                            onClick={() => handleDeleteFile(file.name)}
                            className="text-red-400 hover:text-red-600 p-1 opacity-0 group-hover:opacity-100 transition-opacity"
                          >
                            <X size={12} />
                          </button>
                        </div>
                      ))
                    )}
                  </div>
                </div>

                <div className="md:col-span-2 space-y-3">
                  <h5 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Status Pencocokan Foto</h5>
                  <div className="border border-gray-100 rounded overflow-hidden">
                    <table className="w-full text-left text-[10px]">
                      <thead className="bg-gray-50 font-bold text-gray-500 uppercase">
                        <tr>
                          <th className="px-3 py-2">Nama Karyawan</th>
                          <th className="px-3 py-2">NIK</th>
                          <th className="px-3 py-2 text-center">Status Foto</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-50">
                        {previewData.map((row, idx) => (
                          <tr key={idx}>
                            <td className="px-3 py-2 font-medium">{row.full_name}</td>
                            <td className="px-3 py-2 font-mono">{row.internal_nik}</td>
                            <td className="px-3 py-2 text-center">
                              {row.photo_google_id ? <CheckCircle size={14} className="text-emerald-500 mx-auto" /> : <X size={14} className="text-gray-300 mx-auto" />}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            </div>
          ) : (
            <div className="space-y-6">
              <div className="flex flex-col items-center py-6 border-b border-gray-50">
                <div className="w-16 h-16 bg-emerald-50 text-[#006E62] rounded-full flex items-center justify-center mb-4">
                  <FileBadge size={32} />
                </div>
                <div className="text-center max-w-md">
                  <h4 className="text-lg font-bold text-gray-800">3. Unggah Dokumen Pendukung (Opsional)</h4>
                  <p className="text-xs text-gray-500 mt-2">Unggah file dokumen karyawan (KTP, SK, Kontrak, Ijazah). Sistem akan mencocokkan nama file dengan NIK Internal secara otomatis berdasarkan akhiran nama file.</p>
                </div>

                <div className="mt-4 bg-blue-50 p-3 rounded-md border border-blue-100 w-full max-w-2xl">
                  <p className="text-[10px] font-bold text-blue-800 uppercase mb-2">Aturan Penamaan File:</p>
                  <div className="grid grid-cols-2 md:grid-cols-4 gap-2 text-[9px] text-blue-700">
                    <div><span className="font-bold">KTP:</span> NIK_ktp</div>
                    <div><span className="font-bold">SK:</span> NIK_sk</div>
                    <div><span className="font-bold">Kontrak:</span> NIK_contract</div>
                    <div><span className="font-bold">Ijazah:</span> NIK_diploma</div>
                  </div>
                </div>

                <div className="mt-6 w-full max-w-md">
                  <label className="flex items-center justify-center gap-2 border-2 border-dashed border-gray-200 px-4 py-6 rounded-md hover:bg-gray-50 hover:border-[#006E62] transition-all text-sm font-bold text-gray-400 uppercase tracking-tighter cursor-pointer">
                    {isUploadingAttachments ? <Loader2 size={24} className="animate-spin text-[#006E62]" /> : <Upload size={24} />}
                    {isUploadingAttachments ? 'Sedang Mengunggah...' : 'Klik atau Seret Dokumen ke Sini'}
                    <input 
                      type="file" 
                      className="hidden" 
                      accept="image/*,application/pdf,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document" 
                      multiple 
                      onChange={handleBulkFileUpload} 
                      disabled={isUploadingAttachments} 
                    />
                  </label>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                <div className="md:col-span-1 space-y-3">
                  <h5 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Dokumen Terunggah ({fileList.filter(f => f.name.includes('_') && !f.name.toLowerCase().includes('photo')).length})</h5>
                  <div className="bg-gray-50 rounded-md border border-gray-100 p-2 max-h-[300px] overflow-y-auto space-y-1">
                    {fileList.filter(f => f.name.includes('_') && !f.name.toLowerCase().includes('photo')).length === 0 ? (
                      <p className="text-[10px] text-gray-400 text-center py-4 italic">Belum ada dokumen</p>
                    ) : (
                      fileList.filter(f => f.name.includes('_') && !f.name.toLowerCase().includes('photo')).map((file, idx) => (
                        <div key={idx} className="flex items-center justify-between bg-white p-2 rounded border border-gray-100 group">
                          <div className="flex items-center gap-2 min-w-0">
                            <Paperclip size={12} className="text-[#006E62] shrink-0" />
                            <span className="text-[10px] font-medium text-gray-600 truncate">{file.name}</span>
                          </div>
                          <button 
                            onClick={() => handleDeleteFile(file.name)}
                            className="text-red-400 hover:text-red-600 p-1 opacity-0 group-hover:opacity-100 transition-opacity"
                          >
                            <X size={12} />
                          </button>
                        </div>
                      ))
                    )}
                  </div>
                </div>

                <div className="md:col-span-2 space-y-3">
                  <h5 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Status Pencocokan Dokumen</h5>
                  <div className="border border-gray-100 rounded overflow-hidden">
                    <table className="w-full text-left text-[10px]">
                      <thead className="bg-gray-50 font-bold text-gray-500 uppercase">
                        <tr>
                          <th className="px-3 py-2">Nama Karyawan</th>
                          <th className="px-3 py-2">NIK</th>
                          <th className="px-3 py-2 text-center">KTP</th>
                          <th className="px-3 py-2 text-center">SK</th>
                          <th className="px-3 py-2 text-center">Kontrak</th>
                          <th className="px-3 py-2 text-center">Ijazah</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-50">
                        {previewData.map((row, idx) => (
                          <tr key={idx}>
                            <td className="px-3 py-2 font-medium">{row.full_name}</td>
                            <td className="px-3 py-2 font-mono">{row.internal_nik}</td>
                            <td className="px-3 py-2 text-center">
                              {row.ktp_google_id ? <CheckCircle size={14} className="text-emerald-500 mx-auto" /> : <X size={14} className="text-gray-300 mx-auto" />}
                            </td>
                            <td className="px-3 py-2 text-center">
                              {row.file_sk_id ? <CheckCircle size={14} className="text-emerald-500 mx-auto" /> : <X size={14} className="text-gray-300 mx-auto" />}
                            </td>
                            <td className="px-3 py-2 text-center">
                              {row.contract_initial?.file_id ? <CheckCircle size={14} className="text-emerald-500 mx-auto" /> : <X size={14} className="text-gray-300 mx-auto" />}
                            </td>
                            <td className="px-3 py-2 text-center">
                              {row.diploma_google_id ? <CheckCircle size={14} className="text-emerald-500 mx-auto" /> : <X size={14} className="text-gray-300 mx-auto" />}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>

        <div className="px-6 py-4 border-t border-gray-100 flex justify-end gap-3 bg-gray-50">
          <button type="button" onClick={onClose} className="px-4 py-2 text-xs font-bold text-gray-500 uppercase">Batal</button>
          {step === 1 ? (
            <button 
              onClick={() => setStep(2)}
              disabled={previewData.length === 0 || previewData.some(d => !d.isValid)}
              className="flex items-center gap-2 bg-[#006E62] text-white px-8 py-2 rounded shadow-md hover:bg-[#005a50] transition-all text-xs font-bold uppercase disabled:opacity-50"
            >
              Lanjut
            </button>
          ) : step === 2 ? (
            <div className="flex gap-3">
              <button 
                onClick={() => setStep(1)}
                className="px-4 py-2 text-xs font-bold text-[#006E62] uppercase border border-[#006E62] rounded"
              >
                Kembali
              </button>
              <button 
                onClick={() => setStep(3)}
                className="flex items-center gap-2 bg-[#006E62] text-white px-8 py-2 rounded shadow-md hover:bg-[#005a50] transition-all text-xs font-bold uppercase disabled:opacity-50"
              >
                Lanjut
              </button>
            </div>
          ) : (
            <div className="flex gap-3">
              <button 
                onClick={() => setStep(2)}
                className="px-4 py-2 text-xs font-bold text-[#006E62] uppercase border border-[#006E62] rounded"
              >
                Kembali
              </button>
              <button 
                onClick={handleCommit}
                disabled={isUploading || previewData.some(d => !d.isValid)}
                className="flex items-center gap-2 bg-[#006E62] text-white px-8 py-2 rounded shadow-md hover:bg-[#005a50] transition-all text-xs font-bold uppercase disabled:opacity-50"
              >
                {isUploading ? <Loader2 size={14} className="animate-spin" /> : <Save size={14} />}
                {isUploading ? 'Sedang Memproses...' : 'Simpan Seluruh Data'}
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default AccountImportModal;
