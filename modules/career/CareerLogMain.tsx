import React, { useState, useEffect } from 'react';
import { History, Search, Filter, Download, FileUp, Paperclip, ExternalLink, UserCircle, Upload, Trash2, Edit2, X, Info } from 'lucide-react';
import Swal from 'sweetalert2';
import { careerService } from '../../services/careerService';
import { googleDriveService } from '../../services/googleDriveService';
import { accountService } from '../../services/accountService';
import { CareerLogExtended } from '../../types';
import CareerImportModal from './CareerImportModal';
import LogForm from '../account/LogForm';
import LoadingSpinner from '../../components/Common/LoadingSpinner';

const CareerLogMain: React.FC = () => {
  const [logs, setLogs] = useState<CareerLogExtended[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [searchTerm, setSearchTerm] = useState('');
  const [showImportModal, setShowImportModal] = useState(false);
  const [uploadingId, setUploadingId] = useState<string | null>(null);
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [selectedLog, setSelectedLog] = useState<CareerLogExtended | null>(null);
  const [editingLog, setEditingLog] = useState<CareerLogExtended | null>(null);

  useEffect(() => {
    fetchLogs();
  }, []);

  const fetchLogs = async () => {
    try {
      setIsLoading(true);
      const data = await careerService.getAllGlobal();
      setLogs(data);
    } catch (error) {
      Swal.fire('Gagal', 'Gagal memuat log karir', 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleManualUploadSK = async (e: React.ChangeEvent<HTMLInputElement>, log: CareerLogExtended) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      setUploadingId(log.id);
      const fileId = await googleDriveService.uploadFile(file);
      await accountService.updateCareerLog(log.id, { file_sk_id: fileId });
      
      setLogs(prev => prev.map(l => l.id === log.id ? { ...l, file_sk_id: fileId } : l));
      Swal.fire({ title: 'Berhasil!', text: 'Dokumen SK telah dilampirkan.', icon: 'success', timer: 1500, showConfirmButton: false });
    } catch (error) {
      Swal.fire('Gagal', 'Gagal mengunggah SK', 'error');
    } finally {
      setUploadingId(null);
    }
  };

  const filteredLogs = logs.filter(log => {
    const searchStr = `${log.account?.full_name} ${log.account?.internal_nik} ${log.position} ${log.location_name}`.toLowerCase();
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
      title: 'Hapus Log?',
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
        await careerService.delete(id);
        setLogs(prev => prev.filter(l => l.id !== id));
        setSelectedIds(prev => prev.filter(sid => sid !== id));
        Swal.fire('Terhapus!', 'Log karir telah dihapus.', 'success');
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
        await careerService.bulkDelete(selectedIds);
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

  const handleLogSubmit = async (data: any) => {
    try {
      setIsLoading(true);
      if (editingLog) {
        await accountService.updateCareerLog(editingLog.id, data);
        setLogs(prev => prev.map(l => l.id === editingLog.id ? { ...l, ...data, account: l.account } : l));
        setEditingLog(null);
        Swal.fire({ title: 'Berhasil!', text: 'Log karir telah diperbarui.', icon: 'success', timer: 1500, showConfirmButton: false });
      }
    } catch (error) {
      Swal.fire('Gagal', 'Gagal memperbarui log', 'error');
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="space-y-6">
      {uploadingId && <LoadingSpinner message="Mengunggah SK..." />}

      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="relative flex-1 max-w-md">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
          <input
            type="text"
            placeholder="Cari log (Nama, NIK, Jabatan, Lokasi)..."
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
            onClick={() => careerService.downloadTemplate()}
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
              <th className="px-6 py-4">Perubahan Karir</th>
              <th className="px-6 py-4">Penempatan</th>
              <th className="px-6 py-4">Tgl Efektif</th>
              <th className="px-6 py-4">Dokumen SK</th>
              <th className="px-6 py-4 text-right">Aksi</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-50">
            {isLoading ? (
              <tr><td colSpan={7} className="text-center py-20 text-gray-400">Memuat data log...</td></tr>
            ) : filteredLogs.length === 0 ? (
              <tr><td colSpan={7} className="text-center py-20 text-gray-400">Tidak ada log karir ditemukan.</td></tr>
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
                      <div className="text-xs font-bold text-[#006E62]">{log.position}</div>
                      <div className="text-[10px] text-gray-400 font-bold uppercase tracking-tighter">Grade: {log.grade || '-'}</div>
                    </td>
                    <td className="px-6 py-4">
                      <div className="text-xs text-gray-600 font-medium">{log.location_name}</div>
                    </td>
                    <td className="px-6 py-4 text-xs font-bold text-gray-500">
                      {formatDate(log.change_date)}
                    </td>
                    <td className="px-6 py-4">
                      {log.file_sk_id ? (
                        <a 
                          href={googleDriveService.getFileUrl(log.file_sk_id).replace('=s1600', '=s0')} 
                          target="_blank" 
                          rel="noreferrer"
                          className="inline-flex items-center gap-1.5 text-[10px] font-bold text-[#006E62] bg-emerald-50 px-2 py-1 rounded hover:bg-emerald-100 transition-colors"
                        >
                          <Paperclip size={12} /> LIHAT SK
                        </a>
                      ) : (
                        <label className="inline-flex items-center gap-1.5 text-[10px] font-bold text-orange-500 bg-orange-50 px-2 py-1 rounded cursor-pointer hover:bg-orange-100 transition-colors">
                          <Upload size={12} /> LAMPIRKAN SK
                          <input type="file" className="hidden" accept="image/*,application/pdf" onChange={(e) => handleManualUploadSK(e, log)} />
                        </label>
                      )}
                    </td>
                    <td className="px-6 py-4 text-right" onClick={(e) => e.stopPropagation()}>
                      <div className="flex items-center justify-end gap-2">
                        <button 
                          onClick={() => setEditingLog(log)}
                          className="p-1.5 text-[#006E62] hover:bg-emerald-50 rounded transition-colors"
                          title="Edit Log"
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
        <CareerImportModal 
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
                <h3 className="font-bold text-gray-800">Detail Log Karir</h3>
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
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Jabatan</p>
                  <p className="text-sm font-bold text-[#006E62]">{selectedLog.position}</p>
                </div>
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Grade</p>
                  <p className="text-sm font-bold text-gray-800">{selectedLog.grade || '-'}</p>
                </div>
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Lokasi</p>
                  <p className="text-sm font-medium text-gray-700">{selectedLog.location_name}</p>
                </div>
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tanggal Efektif</p>
                  <p className="text-sm font-bold text-gray-700">{formatDate(selectedLog.change_date)}</p>
                </div>
              </div>

              <div className="space-y-2 pt-4 border-t border-gray-100">
                <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Catatan / Keterangan</p>
                <p className="text-sm text-gray-600 leading-relaxed bg-gray-50 p-3 rounded-lg border border-gray-100 italic">
                  {selectedLog.notes || 'Tidak ada catatan.'}
                </p>
              </div>

              {selectedLog.file_sk_id && (
                <div className="pt-4">
                  <a 
                    href={googleDriveService.getFileUrl(selectedLog.file_sk_id).replace('=s1600', '=s0')}
                    target="_blank"
                    rel="noreferrer"
                    className="flex items-center justify-center gap-2 w-full py-3 bg-emerald-50 text-[#006E62] rounded-lg font-bold text-sm hover:bg-emerald-100 transition-all border border-emerald-200"
                  >
                    <Paperclip size={18} /> LIHAT DOKUMEN SK
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

      {/* Edit Modal - Using Standard LogForm */}
      {editingLog && (
        <LogForm 
          type="career"
          accountId={editingLog.account_id}
          initialData={editingLog}
          isEdit={true}
          onClose={() => setEditingLog(null)}
          onSubmit={handleLogSubmit}
        />
      )}
    </div>
  );
};

export default CareerLogMain;