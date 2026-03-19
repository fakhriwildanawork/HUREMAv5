
import React, { useState, useEffect } from 'react';
import { Activity, Search, Download, FileUp, Paperclip, UserCircle, Upload, Trash2, Edit2, X, Info } from 'lucide-react';
import Swal from 'sweetalert2';
import { healthService } from '../../services/healthService';
import { googleDriveService } from '../../services/googleDriveService';
import { accountService } from '../../services/accountService';
import { HealthLogExtended } from '../../types';
import HealthImportModal from './HealthImportModal';
import LoadingSpinner from '../../components/Common/LoadingSpinner';

const HealthLogMain: React.FC = () => {
  const [logs, setLogs] = useState<HealthLogExtended[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [searchTerm, setSearchTerm] = useState('');
  const [showImportModal, setShowImportModal] = useState(false);
  const [uploadingId, setUploadingId] = useState<string | null>(null);
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [selectedLog, setSelectedLog] = useState<HealthLogExtended | null>(null);
  const [editingLog, setEditingLog] = useState<HealthLogExtended | null>(null);

  useEffect(() => {
    fetchLogs();
  }, []);

  const fetchLogs = async () => {
    try {
      setIsLoading(true);
      const data = await healthService.getAllGlobal();
      setLogs(data);
    } catch (error) {
      Swal.fire('Gagal', 'Gagal memuat log kesehatan', 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleManualUploadMCU = async (e: React.ChangeEvent<HTMLInputElement>, log: HealthLogExtended) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      setUploadingId(log.id);
      const fileId = await googleDriveService.uploadFile(file);
      await accountService.updateHealthLog(log.id, { file_mcu_id: fileId });
      
      setLogs(prev => prev.map(l => l.id === log.id ? { ...l, file_mcu_id: fileId } : l));
      Swal.fire({ title: 'Berhasil!', text: 'Dokumen MCU telah dilampirkan.', icon: 'success', timer: 1500, showConfirmButton: false });
    } catch (error) {
      Swal.fire('Gagal', 'Gagal mengunggah dokumen', 'error');
    } finally {
      setUploadingId(null);
    }
  };

  const filteredLogs = logs.filter(log => {
    const searchStr = `${log.account?.full_name} ${log.account?.internal_nik} ${log.mcu_status} ${log.health_risk}`.toLowerCase();
    return searchStr.includes(searchTerm.toLowerCase());
  });

  const formatDate = (dateStr: string) => {
    return new Date(dateStr).toLocaleDateString('id-ID', {
      day: 'numeric',
      month: 'short',
      year: 'numeric'
    });
  };

  const handleDelete = async (id: string) => {
    const result = await Swal.fire({
      title: 'Hapus Log Kesehatan?',
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
        await healthService.delete(id);
        setLogs(prev => prev.filter(l => l.id !== id));
        setSelectedIds(prev => prev.filter(sid => sid !== id));
        Swal.fire('Terhapus!', 'Log kesehatan telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Gagal menghapus log', 'error');
      } finally {
        setIsLoading(false);
      }
    }
  };

  const handleBulkDelete = async () => {
    if (selectedIds.length === 0) return;

    const result = await Swal.fire({
      title: 'Hapus Masal?',
      text: `Apakah Anda yakin ingin menghapus ${selectedIds.length} log terpilih secara permanen?`,
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
        await healthService.bulkDelete(selectedIds);
        setLogs(prev => prev.filter(l => !selectedIds.includes(l.id)));
        setSelectedIds([]);
        Swal.fire('Berhasil!', 'Data terpilih telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Beberapa data mungkin gagal dihapus.', 'error');
        fetchLogs();
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
    if (selectedIds.length === filteredLogs.length) {
      setSelectedIds([]);
    } else {
      setSelectedIds(filteredLogs.map(l => l.id));
    }
  };

  return (
    <div className="space-y-6">
      {uploadingId && <LoadingSpinner message="Mengunggah Dokumen..." />}

      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="relative flex-1 max-w-md">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
          <input
            type="text"
            placeholder="Cari log (Nama, NIK, Status MCU)..."
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
          <button 
            onClick={() => healthService.downloadTemplate()}
            className="flex items-center gap-2 border border-gray-200 px-4 py-2 rounded-md hover:bg-gray-50 transition-colors text-sm font-medium text-gray-600"
          >
            <Download size={18} /> Unduh Template
          </button>
          <button 
            onClick={() => setShowImportModal(true)}
            className="flex items-center gap-2 bg-[#006E62] text-white px-4 py-2 rounded-md hover:bg-[#005a50] transition-colors shadow-sm text-sm font-medium"
          >
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
                  checked={selectedIds.length === filteredLogs.length && filteredLogs.length > 0}
                  onChange={toggleSelectAll}
                />
              </th>
              <th className="px-6 py-4">Karyawan</th>
              <th className="px-6 py-4">Status MCU</th>
              <th className="px-6 py-4">Risiko Kesehatan</th>
              <th className="px-6 py-4">Tgl Periksa</th>
              <th className="px-6 py-4">Dokumen MCU</th>
              <th className="px-6 py-4 text-right">Aksi</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-50">
            {isLoading ? (
              <tr><td colSpan={7} className="text-center py-20 text-gray-400">Memuat data log kesehatan...</td></tr>
            ) : filteredLogs.length === 0 ? (
              <tr><td colSpan={7} className="text-center py-20 text-gray-400">Tidak ada log kesehatan ditemukan.</td></tr>
            ) : (
              filteredLogs.map(log => {
                const isSelected = selectedIds.includes(log.id);
                return (
                  <tr 
                    key={log.id} 
                    className={`hover:bg-gray-50/50 transition-colors cursor-pointer ${isSelected ? 'bg-emerald-50/20' : ''}`}
                    onClick={() => setSelectedLog(log)}
                  >
                    <td className="px-6 py-4" onClick={(e) => e.stopPropagation()}>
                      <input 
                        type="checkbox" 
                        className="rounded border-gray-300 text-[#006E62] focus:ring-[#006E62]"
                        checked={isSelected}
                        onChange={() => toggleSelect(log.id)}
                      />
                    </td>
                    <td className="px-6 py-4">
                      <div className="flex items-center gap-3">
                        <div className="w-8 h-8 rounded-full bg-gray-100 text-gray-400 flex items-center justify-center border border-gray-200 overflow-hidden">
                          {log.account?.photo_google_id ? (
                            <img 
                              src={googleDriveService.getFileUrl(log.account.photo_google_id)} 
                              alt="" 
                              className="w-full h-full object-cover"
                              referrerPolicy="no-referrer"
                            />
                          ) : (
                            <UserCircle size={20} />
                          )}
                        </div>
                        <div>
                          <div className="text-xs font-bold text-gray-800">{log.account?.full_name}</div>
                          <div className="text-[10px] text-gray-400 font-mono uppercase">{log.account?.internal_nik}</div>
                        </div>
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <div className="text-xs font-bold text-[#006E62] uppercase tracking-tighter">{log.mcu_status || '-'}</div>
                    </td>
                    <td className="px-6 py-4">
                      <div className={`text-[10px] font-bold px-2 py-0.5 rounded inline-block uppercase ${
                        log.health_risk?.toLowerCase().includes('tinggi') ? 'bg-red-50 text-red-600' :
                        log.health_risk?.toLowerCase().includes('sedang') ? 'bg-orange-50 text-orange-600' : 'bg-green-50 text-green-600'
                      }`}>
                        {log.health_risk || 'Normal'}
                      </div>
                    </td>
                    <td className="px-6 py-4 text-xs font-bold text-gray-500">
                      {formatDate(log.change_date)}
                    </td>
                    <td className="px-6 py-4">
                      {log.file_mcu_id ? (
                        <a 
                          href={googleDriveService.getFileUrl(log.file_mcu_id).replace('=s1600', '=s0')} 
                          target="_blank" 
                          rel="noreferrer"
                          className="inline-flex items-center gap-1.5 text-[10px] font-bold text-[#006E62] bg-[#006E62]/10 px-2 py-1 rounded hover:bg-[#006E62]/20 transition-colors"
                        >
                          <Paperclip size={12} /> LIHAT HASIL
                        </a>
                      ) : (
                        <label className="inline-flex items-center gap-1.5 text-[10px] font-bold text-orange-500 bg-orange-50 px-2 py-1 rounded cursor-pointer hover:bg-orange-100 transition-colors">
                          <Upload size={12} /> LAMPIRKAN
                          <input type="file" className="hidden" accept="image/*,application/pdf" onChange={(e) => handleManualUploadMCU(e, log)} />
                        </label>
                      )}
                    </td>
                    <td className="px-6 py-4 text-right" onClick={(e) => e.stopPropagation()}>
                      <div className="flex items-center justify-end gap-2">
                        <button 
                          onClick={() => setEditingLog(log)}
                          className="p-1.5 text-[#006E62] hover:bg-emerald-50 rounded transition-colors"
                          title="Edit Log Kesehatan"
                        >
                          <Edit2 size={14} />
                        </button>
                        <button 
                          onClick={() => handleDelete(log.id)}
                          className="p-1.5 text-red-400 hover:bg-red-50 rounded transition-colors"
                          title="Hapus Log"
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
        <HealthImportModal 
          onClose={() => setShowImportModal(false)} 
          onSuccess={() => { setShowImportModal(false); fetchLogs(); }} 
        />
      )}

      {/* Detail Modal */}
      {selectedLog && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-lg overflow-hidden animate-in fade-in zoom-in duration-200">
            <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
              <div className="flex items-center gap-2 text-[#006E62]">
                <Info size={20} />
                <h3 className="font-bold text-gray-800">Detail Log Kesehatan</h3>
              </div>
              <button onClick={() => setSelectedLog(null)} className="p-2 hover:bg-gray-200 rounded-full transition-colors">
                <X size={20} className="text-gray-500" />
              </button>
            </div>
            <div className="p-6 space-y-6">
              <div className="flex items-center gap-4 p-4 bg-emerald-50/30 rounded-lg border border-emerald-100/50">
                <div className="w-12 h-12 rounded-full bg-gray-100 flex items-center justify-center border border-gray-200 text-gray-400 shadow-sm overflow-hidden">
                  {selectedLog.account?.photo_google_id ? (
                    <img 
                      src={googleDriveService.getFileUrl(selectedLog.account.photo_google_id)} 
                      alt="" 
                      className="w-full h-full object-cover"
                      referrerPolicy="no-referrer"
                    />
                  ) : (
                    <UserCircle size={32} />
                  )}
                </div>
                <div>
                  <h4 className="font-bold text-gray-900">{selectedLog.account?.full_name}</h4>
                  <p className="text-xs font-mono text-gray-500 uppercase tracking-wider">{selectedLog.account?.internal_nik}</p>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-6">
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Status MCU</p>
                  <p className="text-sm font-bold text-[#006E62]">{selectedLog.mcu_status}</p>
                </div>
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Risiko Kesehatan</p>
                  <p className="text-sm font-bold text-gray-800">{selectedLog.health_risk}</p>
                </div>
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tanggal Periksa</p>
                  <p className="text-sm font-medium text-gray-700">{formatDate(selectedLog.change_date)}</p>
                </div>
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tanggal Input</p>
                  <p className="text-sm font-bold text-gray-700">{formatDate(selectedLog.entry_date)}</p>
                </div>
              </div>

              <div className="space-y-2 pt-4 border-t border-gray-100">
                <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Catatan / Keterangan</p>
                <p className="text-sm text-gray-600 leading-relaxed bg-gray-50 p-3 rounded-lg border border-gray-100 italic">
                  {selectedLog.notes || 'Tidak ada catatan.'}
                </p>
              </div>

              {selectedLog.file_mcu_id && (
                <div className="pt-4">
                  <a 
                    href={googleDriveService.getFileUrl(selectedLog.file_mcu_id, true)}
                    target="_blank"
                    rel="noreferrer"
                    className="flex items-center justify-center gap-2 w-full py-3 bg-emerald-50 text-[#006E62] rounded-lg font-bold text-sm hover:bg-emerald-100 transition-all border border-emerald-200"
                  >
                    <Paperclip size={18} /> LIHAT DOKUMEN HASIL
                  </a>
                </div>
              )}
            </div>
            <div className="px-6 py-4 bg-gray-50 border-t border-gray-100 flex justify-end">
              <button 
                onClick={() => setSelectedLog(null)}
                className="px-6 py-2 bg-white border border-gray-200 text-gray-600 rounded-lg text-sm font-bold hover:bg-gray-100 transition-all"
              >
                Tutup
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Edit Modal */}
      {editingLog && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-md overflow-hidden animate-in fade-in zoom-in duration-200">
            <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
              <div className="flex items-center gap-2 text-[#006E62]">
                <Edit2 size={20} />
                <h3 className="font-bold text-gray-800">Edit Log Kesehatan</h3>
              </div>
              <button onClick={() => setEditingLog(null)} className="p-2 hover:bg-gray-200 rounded-full transition-colors">
                <X size={20} className="text-gray-500" />
              </button>
            </div>
            <form onSubmit={async (e) => {
              e.preventDefault();
              const formData = new FormData(e.currentTarget);
              const file = (formData.get('file_mcu') as File);
              
              const data: any = {
                mcu_status: formData.get('mcu_status') as string,
                health_risk: formData.get('health_risk') as string,
                change_date: formData.get('change_date') as string,
                notes: formData.get('notes') as string,
              };
              
              try {
                setIsLoading(true);
                
                if (file && file.size > 0) {
                  if (editingLog.file_mcu_id) {
                    await googleDriveService.deleteFile(editingLog.file_mcu_id);
                  }
                  const newFileId = await googleDriveService.uploadFile(file);
                  data.file_mcu_id = newFileId;
                }

                await healthService.update(editingLog.id, data);
                setLogs(prev => prev.map(l => l.id === editingLog.id ? { ...l, ...data } : l));
                setEditingLog(null);
                Swal.fire({ title: 'Berhasil!', text: 'Data kesehatan telah diperbarui.', icon: 'success', timer: 1500, showConfirmButton: false });
              } catch (error) {
                Swal.fire('Gagal', 'Gagal memperbarui data kesehatan', 'error');
              } finally {
                setIsLoading(false);
              }
            }} className="p-6 space-y-4">
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Status MCU</label>
                <input 
                  name="mcu_status"
                  defaultValue={editingLog.mcu_status}
                  required
                  className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-[#006E62] text-sm font-medium"
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Risiko Kesehatan</label>
                <select 
                  name="health_risk"
                  defaultValue={editingLog.health_risk}
                  required
                  className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-[#006E62] text-sm font-medium"
                >
                  <option value="Normal">Normal</option>
                  <option value="Rendah">Rendah</option>
                  <option value="Sedang">Sedang</option>
                  <option value="Tinggi">Tinggi</option>
                </select>
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tanggal Periksa</label>
                <input 
                  type="date"
                  name="change_date"
                  defaultValue={editingLog.change_date}
                  required
                  className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-[#006E62] text-sm font-medium"
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Catatan</label>
                <textarea 
                  name="notes"
                  defaultValue={editingLog.notes || ''}
                  rows={3}
                  className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-[#006E62] text-sm font-medium resize-none"
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Ganti Dokumen MCU (Opsional)</label>
                <input 
                  type="file"
                  name="file_mcu"
                  accept="image/*,application/pdf"
                  className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-[#006E62] text-sm font-medium"
                />
                {editingLog.file_mcu_id && <p className="text-[10px] text-orange-500 font-medium italic">* Mengunggah file baru akan menghapus file lama.</p>}
              </div>
              <div className="pt-4 flex gap-3">
                <button 
                  type="button"
                  onClick={() => setEditingLog(null)}
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

export default HealthLogMain;
