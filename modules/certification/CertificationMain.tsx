
import React, { useState, useEffect } from 'react';
import { Award, Search, Paperclip, UserCircle, Upload, Calendar, FileUp, Download, Trash2, Edit2, X, Info } from 'lucide-react';
import Swal from 'sweetalert2';
import { certificationService } from '../../services/certificationService';
import { googleDriveService } from '../../services/googleDriveService';
import { AccountCertificationExtended } from '../../types';
import LoadingSpinner from '../../components/Common/LoadingSpinner';
import CertificationImportModal from './CertificationImportModal';

const CertificationMain: React.FC = () => {
  const [certs, setCerts] = useState<AccountCertificationExtended[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [searchTerm, setSearchTerm] = useState('');
  const [showImportModal, setShowImportModal] = useState(false);
  const [uploadingId, setUploadingId] = useState<string | null>(null);
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [selectedCert, setSelectedCert] = useState<AccountCertificationExtended | null>(null);
  const [editingCert, setEditingCert] = useState<AccountCertificationExtended | null>(null);

  useEffect(() => {
    fetchCerts();
  }, []);

  const fetchCerts = async () => {
    try {
      setIsLoading(true);
      const data = await certificationService.getAllGlobal();
      setCerts(data);
    } catch (error) {
      Swal.fire('Gagal', 'Gagal memuat data sertifikasi', 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleManualUpload = async (e: React.ChangeEvent<HTMLInputElement>, cert: AccountCertificationExtended) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      setUploadingId(cert.id);
      const fileId = await googleDriveService.uploadFile(file);
      await certificationService.update(cert.id, { file_id: fileId });
      
      setCerts(prev => prev.map(c => c.id === cert.id ? { ...c, file_id: fileId } : c));
      Swal.fire({ title: 'Berhasil!', text: 'Dokumen sertifikat telah dilampirkan.', icon: 'success', timer: 1500, showConfirmButton: false });
    } catch (error) {
      Swal.fire('Gagal', 'Gagal mengunggah dokumen', 'error');
    } finally {
      setUploadingId(null);
    }
  };

  const filteredCerts = certs.filter(c => {
    const searchStr = `${c.account?.full_name} ${c.account?.internal_nik} ${c.cert_name} ${c.cert_type}`.toLowerCase();
    return searchStr.includes(searchTerm.toLowerCase());
  });

  const formatDate = (dateStr?: string | null) => {
    if (!dateStr) return '-';
    return new Date(dateStr).toLocaleDateString('id-ID', {
      day: 'numeric',
      month: 'short',
      year: 'numeric'
    });
  };

  const handleDelete = async (id: string) => {
    const result = await Swal.fire({
      title: 'Hapus Sertifikasi?',
      text: "Data ini akan dihapus permanen dari sistem.",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#ef4444',
      cancelButtonColor: '#6b7280',
      confirmButtonText: 'Ya, Hapus',
      cancelButtonText: 'Batal'
    });

    if (result.isConfirmed) {
      try {
        setIsLoading(true);
        await certificationService.delete(id);
        setCerts(prev => prev.filter(c => c.id !== id));
        setSelectedIds(prev => prev.filter(sid => sid !== id));
        Swal.fire('Terhapus!', 'Sertifikasi telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Gagal menghapus sertifikasi', 'error');
      } finally {
        setIsLoading(false);
      }
    }
  };

  const handleBulkDelete = async () => {
    if (selectedIds.length === 0) return;

    const result = await Swal.fire({
      title: 'Hapus Masal?',
      text: `Apakah Anda yakin ingin menghapus ${selectedIds.length} sertifikasi terpilih secara permanen?`,
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#ef4444',
      cancelButtonColor: '#6b7280',
      confirmButtonText: 'Ya, Hapus Semua',
      cancelButtonText: 'Batal'
    });

    if (result.isConfirmed) {
      try {
        setIsLoading(true);
        await certificationService.bulkDelete(selectedIds);
        setCerts(prev => prev.filter(c => !selectedIds.includes(c.id)));
        setSelectedIds([]);
        Swal.fire('Berhasil!', 'Data terpilih telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Beberapa data mungkin gagal dihapus.', 'error');
        fetchCerts();
      } finally {
        setIsLoading(false);
      }
    }
  };

  const toggleSelect = (id: string) => {
    setSelectedIds(prev => 
      prev.includes(id) ? prev.filter(i => i !== id) : [...prev, id]
    );
  };

  const toggleSelectAll = () => {
    if (selectedIds.length === filteredCerts.length) {
      setSelectedIds([]);
    } else {
      setSelectedIds(filteredCerts.map(c => c.id));
    }
  };

  return (
    <div className="space-y-6">
      {uploadingId && <LoadingSpinner message="Mengunggah Dokumen..." />}

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">
        <div className="bg-white border border-gray-100 p-4 rounded-md shadow-sm flex items-center gap-4">
          <div className="p-3 bg-emerald-50 rounded-md text-[#006E62]"><Award size={24} /></div>
          <div>
            <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Total Sertifikasi Terdata</p>
            <p className="text-xl font-bold text-gray-800">{certs.length}</p>
          </div>
        </div>
        <div className="bg-white border border-gray-100 p-4 rounded-md shadow-sm flex items-center gap-4">
          <div className="p-3 bg-blue-50 rounded-md text-blue-600"><Calendar size={24} /></div>
          <div>
            <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Input Bulan Ini</p>
            <p className="text-xl font-bold text-gray-800">
              {certs.filter(c => {
                const inputDate = new Date(c.entry_date);
                const now = new Date();
                return inputDate.getMonth() === now.getMonth() && inputDate.getFullYear() === now.getFullYear();
              }).length}
            </p>
          </div>
        </div>
      </div>

      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="relative flex-1 max-w-md">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
          <input
            type="text"
            placeholder="Cari sertifikasi (Nama Karyawan, Jenis, Nama Sertifikat)..."
            className="w-full pl-10 pr-4 py-2 bg-gray-50 border border-gray-200 rounded-md focus:outline-none focus:ring-2 focus:ring-[#006E62] text-sm"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </div>
        
        <div className="flex items-center gap-2">
          {selectedIds.length > 0 && (
            <button 
              onClick={handleBulkDelete}
              className="flex items-center gap-2 bg-red-50 text-red-600 px-4 py-2 rounded-md border border-red-100 hover:bg-red-100 transition-all text-sm font-medium mr-2"
            >
              <Trash2 size={18} /> Hapus ({selectedIds.length})
            </button>
          )}
          <button onClick={() => certificationService.downloadTemplate()} className="flex items-center gap-2 border border-gray-200 px-4 py-2 rounded-md hover:bg-gray-50 transition-colors text-sm font-medium text-gray-600">
            <Download size={18} /> Unduh Template
          </button>
          <button onClick={() => setShowImportModal(true)} className="flex items-center gap-2 bg-[#006E62] text-white px-4 py-2 rounded-md hover:bg-[#005a50] transition-colors shadow-sm text-sm font-medium">
            <FileUp size={18} /> Impor Massal
          </button>
        </div>
      </div>

      <div className="bg-white border border-gray-100 rounded-md overflow-hidden shadow-sm">
        <table className="w-full text-left border-collapse">
          <thead className="bg-gray-50 text-[10px] font-bold text-gray-500 uppercase tracking-widest border-b border-gray-100">
            <tr>
              <th className="px-6 py-4 w-10">
                <input 
                  type="checkbox" 
                  className="rounded border-gray-300 text-[#006E62] focus:ring-[#006E62]"
                  checked={selectedIds.length === filteredCerts.length && filteredCerts.length > 0}
                  onChange={toggleSelectAll}
                />
              </th>
              <th className="px-6 py-4">Karyawan</th>
              <th className="px-6 py-4">Sertifikasi</th>
              <th className="px-6 py-4">Tanggal</th>
              <th className="px-6 py-4">Dokumen</th>
              <th className="px-6 py-4 text-right">Aksi</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-50">
            {isLoading ? (
              <tr><td colSpan={6} className="text-center py-20 text-gray-400">Memuat data sertifikasi...</td></tr>
            ) : filteredCerts.length === 0 ? (
              <tr><td colSpan={6} className="text-center py-20 text-gray-400">Tidak ada data sertifikasi ditemukan.</td></tr>
            ) : (
              filteredCerts.map(c => {
                const isSelected = selectedIds.includes(c.id);
                return (
                  <tr 
                    key={c.id} 
                    className={`hover:bg-gray-50/50 transition-colors cursor-pointer group ${isSelected ? 'bg-emerald-50/20' : ''}`}
                    onClick={() => setSelectedCert(c)}
                  >
                    <td className="px-6 py-4" onClick={(e) => e.stopPropagation()}>
                      <input 
                        type="checkbox" 
                        className="rounded border-gray-300 text-[#006E62] focus:ring-[#006E62]"
                        checked={isSelected}
                        onChange={() => toggleSelect(c.id)}
                      />
                    </td>
                    <td className="px-6 py-4">
                      <div className="flex items-center gap-3">
                        <div className="w-8 h-8 rounded-full bg-gray-100 text-gray-400 flex items-center justify-center border border-gray-200 overflow-hidden">
                          {c.account?.photo_google_id ? (
                            <img 
                              src={googleDriveService.getFileUrl(c.account.photo_google_id)} 
                              alt="" 
                              className="w-full h-full object-cover"
                              referrerPolicy="no-referrer"
                            />
                          ) : (
                            <UserCircle size={20} />
                          )}
                        </div>
                        <div>
                          <div className="text-xs font-bold text-gray-800">{c.account?.full_name}</div>
                          <div className="text-[10px] text-gray-400 font-mono uppercase">{c.account?.internal_nik}</div>
                        </div>
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <div className="text-xs font-bold text-[#006E62]">{c.cert_name}</div>
                      <div className="text-[10px] text-gray-400 font-bold uppercase tracking-tighter">{c.cert_type}</div>
                    </td>
                    <td className="px-6 py-4">
                      <div className="text-xs text-gray-600 font-medium">{formatDate(c.cert_date)}</div>
                      <div className="text-[9px] text-gray-400 uppercase font-bold">Entry: {formatDate(c.entry_date)}</div>
                    </td>
                    <td className="px-6 py-4">
                      {c.file_id ? (
                        <a 
                          href={googleDriveService.getFileUrl(c.file_id).replace('=s1600', '=s0')} 
                          target="_blank" 
                          rel="noreferrer"
                          className="inline-flex items-center gap-1.5 text-[10px] font-bold text-[#006E62] bg-emerald-50 px-2 py-1 rounded hover:bg-emerald-100 transition-colors"
                        >
                          <Paperclip size={12} /> LIHAT FILE
                        </a>
                      ) : (
                        <label className="inline-flex items-center gap-1.5 text-[10px] font-bold text-orange-500 bg-orange-50 px-2 py-1 rounded cursor-pointer hover:bg-orange-100 transition-colors">
                          <Upload size={12} /> UPLOAD
                          <input type="file" className="hidden" accept="image/*,application/pdf" onChange={(e) => handleManualUpload(e, c)} />
                        </label>
                      )}
                    </td>
                    <td className="px-6 py-4 text-right" onClick={(e) => e.stopPropagation()}>
                      <div className="flex items-center justify-end gap-2">
                        <button 
                          onClick={() => setEditingCert(c)}
                          className="p-1.5 text-[#006E62] hover:bg-emerald-50 rounded transition-colors"
                          title="Edit Sertifikasi"
                        >
                          <Edit2 size={14} />
                        </button>
                        <button 
                          onClick={() => handleDelete(c.id)}
                          className="p-1.5 text-red-400 hover:bg-red-50 rounded transition-colors"
                          title="Hapus Sertifikasi"
                        >
                          <Trash2 size={14} />
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>

      {showImportModal && (
        <CertificationImportModal 
          onClose={() => setShowImportModal(false)} 
          onSuccess={() => { setShowImportModal(false); fetchCerts(); }} 
        />
      )}

      {/* Detail Modal */}
      {selectedCert && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-lg overflow-hidden animate-in fade-in zoom-in duration-200">
            <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
              <div className="flex items-center gap-2 text-[#006E62]">
                <Info size={20} />
                <h3 className="font-bold text-gray-800">Detail Sertifikasi</h3>
              </div>
              <button onClick={() => setSelectedCert(null)} className="p-2 hover:bg-gray-200 rounded-full transition-colors">
                <X size={20} className="text-gray-500" />
              </button>
            </div>
            <div className="p-6 space-y-6">
              <div className="flex items-center gap-4 p-4 bg-emerald-50/30 rounded-lg border border-emerald-100/50">
                <div className="w-12 h-12 rounded-full bg-gray-100 flex items-center justify-center border border-gray-200 text-gray-400 shadow-sm overflow-hidden">
                  {selectedCert.account?.photo_google_id ? (
                    <img 
                      src={googleDriveService.getFileUrl(selectedCert.account.photo_google_id)} 
                      alt="" 
                      className="w-full h-full object-cover"
                      referrerPolicy="no-referrer"
                    />
                  ) : (
                    <UserCircle size={32} />
                  )}
                </div>
                <div>
                  <h4 className="font-bold text-gray-900">{selectedCert.account?.full_name}</h4>
                  <p className="text-xs font-mono text-gray-500 uppercase tracking-wider">{selectedCert.account?.internal_nik}</p>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-6">
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Nama Sertifikasi</p>
                  <p className="text-sm font-bold text-[#006E62]">{selectedCert.cert_name}</p>
                </div>
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Jenis Sertifikasi</p>
                  <p className="text-sm font-bold text-gray-800">{selectedCert.cert_type}</p>
                </div>
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tanggal Sertifikasi</p>
                  <p className="text-sm font-medium text-gray-700">{formatDate(selectedCert.cert_date)}</p>
                </div>
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tanggal Input</p>
                  <p className="text-sm font-bold text-gray-700">{formatDate(selectedCert.entry_date)}</p>
                </div>
              </div>

              <div className="space-y-2 pt-4 border-t border-gray-100">
                <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Catatan / Keterangan</p>
                <p className="text-sm text-gray-600 leading-relaxed bg-gray-50 p-3 rounded-lg border border-gray-100 italic">
                  {selectedCert.notes || 'Tidak ada catatan.'}
                </p>
              </div>

              {selectedCert.file_id && (
                <div className="pt-4">
                  <a 
                    href={googleDriveService.getFileUrl(selectedCert.file_id).replace('=s1600', '=s0')}
                    target="_blank"
                    rel="noreferrer"
                    className="flex items-center justify-center gap-2 w-full py-3 bg-emerald-50 text-[#006E62] rounded-lg font-bold text-sm hover:bg-emerald-100 transition-all border border-emerald-200"
                  >
                    <Paperclip size={18} /> LIHAT DOKUMEN SERTIFIKAT
                  </a>
                </div>
              )}
            </div>
            <div className="px-6 py-4 bg-gray-50 border-t border-gray-100 flex justify-end">
              <button 
                onClick={() => setSelectedCert(null)}
                className="px-6 py-2 bg-white border border-gray-200 text-gray-600 rounded-lg text-sm font-bold hover:bg-gray-100 transition-all"
              >
                Tutup
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Edit Modal */}
      {editingCert && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-md overflow-hidden animate-in fade-in zoom-in duration-200">
            <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
              <div className="flex items-center gap-2 text-[#006E62]">
                <Edit2 size={20} />
                <h3 className="font-bold text-gray-800">Edit Sertifikasi</h3>
              </div>
              <button onClick={() => setEditingCert(null)} className="p-2 hover:bg-gray-200 rounded-full transition-colors">
                <X size={20} className="text-gray-500" />
              </button>
            </div>
            <form onSubmit={async (e) => {
              e.preventDefault();
              const formData = new FormData(e.currentTarget);
              const file = (formData.get('file_sertifikat') as File);
              
              const data: any = {
                cert_name: formData.get('cert_name') as string,
                cert_type: formData.get('cert_type') as string,
                cert_date: formData.get('cert_date') as string,
                notes: formData.get('notes') as string,
              };
              
              try {
                setIsLoading(true);
                
                if (file && file.size > 0) {
                  if (editingCert.file_id) {
                    await googleDriveService.deleteFile(editingCert.file_id);
                  }
                  const newFileId = await googleDriveService.uploadFile(file);
                  data.file_id = newFileId;
                }

                await certificationService.update(editingCert.id, data);
                setCerts(prev => prev.map(c => c.id === editingCert.id ? { ...c, ...data } : c));
                setEditingCert(null);
                Swal.fire({ title: 'Berhasil!', text: 'Data sertifikasi telah diperbarui.', icon: 'success', timer: 1500, showConfirmButton: false });
              } catch (error) {
                Swal.fire('Gagal', 'Gagal memperbarui data sertifikasi', 'error');
              } finally {
                setIsLoading(false);
              }
            }} className="p-6 space-y-4">
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Nama Sertifikasi</label>
                <input 
                  name="cert_name"
                  defaultValue={editingCert.cert_name}
                  required
                  className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-[#006E62] text-sm font-medium"
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Jenis Sertifikasi</label>
                <select 
                  name="cert_type"
                  defaultValue={editingCert.cert_type}
                  required
                  className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-[#006E62] text-sm font-medium"
                >
                  <option value="Internal">Internal</option>
                  <option value="Eksternal">Eksternal</option>
                  <option value="Lisensi">Lisensi</option>
                  <option value="Lainnya">Lainnya</option>
                </select>
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tanggal Sertifikasi</label>
                <input 
                  type="date"
                  name="cert_date"
                  defaultValue={editingCert.cert_date}
                  required
                  className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-[#006E62] text-sm font-medium"
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Catatan</label>
                <textarea 
                  name="notes"
                  defaultValue={editingCert.notes || ''}
                  rows={3}
                  className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-[#006E62] text-sm font-medium resize-none"
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Ganti Dokumen Sertifikat (Opsional)</label>
                <input 
                  type="file"
                  name="file_sertifikat"
                  accept="image/*,application/pdf"
                  className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-[#006E62] text-sm font-medium"
                />
                {editingCert.file_id && <p className="text-[10px] text-orange-500 font-medium italic">* Mengunggah file baru akan menghapus file lama.</p>}
              </div>
              <div className="pt-4 flex gap-3">
                <button 
                  type="button"
                  onClick={() => setEditingCert(null)}
                  className="flex-1 py-2.5 bg-white border border-gray-200 text-gray-600 rounded-lg text-sm font-bold hover:bg-gray-100 transition-all"
                >
                  Batal
                </button>
                <button 
                  type="submit"
                  className="flex-1 py-2.5 bg-[#006E62] text-white rounded-lg text-sm font-bold hover:bg-[#005a50] transition-all shadow-md shadow-emerald-100"
                >
                  Simpan Perubahan
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};

export default CertificationMain;
