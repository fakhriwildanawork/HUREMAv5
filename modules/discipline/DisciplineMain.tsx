import React, { useState, useEffect } from 'react';
import { ShieldAlert, LogOut, Search, Download, FileUp, Paperclip, UserCircle, Plus, Trash2, ArrowRight, Edit2, X, Info } from 'lucide-react';
import Swal from 'sweetalert2';
import { disciplineService } from '../../services/disciplineService';
import { googleDriveService } from '../../services/googleDriveService';
import { WarningLogExtended, TerminationLogExtended } from '../../types';
import LoadingSpinner from '../../components/Common/LoadingSpinner';
import DisciplineImportModal from './DisciplineImportModal';

const DisciplineMain: React.FC = () => {
  const [warnings, setWarnings] = useState<WarningLogExtended[]>([]);
  const [terminations, setTerminations] = useState<TerminationLogExtended[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [activeTab, setActiveTab] = useState<'warnings' | 'terminations'>('warnings');
  const [searchTerm, setSearchTerm] = useState('');
  const [showImportModal, setShowImportModal] = useState(false);
  const [importType, setImportType] = useState<'warning' | 'termination'>('warning');
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [selectedWarning, setSelectedWarning] = useState<WarningLogExtended | null>(null);
  const [editingWarning, setEditingWarning] = useState<WarningLogExtended | null>(null);
  const [selectedTermination, setSelectedTermination] = useState<TerminationLogExtended | null>(null);
  const [editingTermination, setEditingTermination] = useState<TerminationLogExtended | null>(null);

  useEffect(() => {
    fetchData();
  }, []);

  const fetchData = async () => {
    try {
      setIsLoading(true);
      const [w, t] = await Promise.all([
        disciplineService.getWarningsAll(),
        disciplineService.getTerminationsAll()
      ]);
      setWarnings(w);
      setTerminations(t);
    } catch (error) {
      Swal.fire('Gagal', 'Gagal memuat data kedisiplinan', 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleDeleteWarning = async (id: string) => {
    const res = await Swal.fire({ 
      title: 'Hapus log peringatan?', 
      text: "Data ini akan dihapus permanen dari sistem.",
      icon: 'warning', 
      showCancelButton: true, 
      confirmButtonColor: '#ef4444',
      cancelButtonColor: '#6b7280',
      confirmButtonText: 'Ya, Hapus',
      cancelButtonText: 'Batal'
    });
    if (res.isConfirmed) {
      try {
        setIsLoading(true);
        await disciplineService.deleteWarning(id);
        setWarnings(prev => prev.filter(w => w.id !== id));
        setSelectedIds(prev => prev.filter(sid => sid !== id));
        Swal.fire('Terhapus', 'Log peringatan telah dihapus.', 'success');
      } catch (e) { 
        Swal.fire('Gagal', 'Gagal menghapus log', 'error'); 
      } finally {
        setIsLoading(false);
      }
    }
  };

  const handleDeleteTermination = async (id: string, accountId: string) => {
    const res = await Swal.fire({ 
      title: 'Hapus log pengakhiran?', 
      text: "Data ini akan dihapus permanen dari sistem.",
      icon: 'warning', 
      showCancelButton: true, 
      confirmButtonColor: '#ef4444', 
      cancelButtonColor: '#6b7280',
      confirmButtonText: 'Ya, Hapus',
      cancelButtonText: 'Batal'
    });
    if (res.isConfirmed) {
      try {
        setIsLoading(true);
        await disciplineService.deleteTermination(id, accountId);
        setTerminations(prev => prev.filter(t => t.id !== id));
        setSelectedIds(prev => prev.filter(sid => sid !== id));
        Swal.fire('Terhapus', 'Log pengakhiran telah dihapus.', 'success');
      } catch (e) { 
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
      text: `Apakah Anda yakin ingin menghapus ${selectedIds.length} data terpilih secara permanen?`,
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
        if (activeTab === 'warnings') {
          await disciplineService.bulkDeleteWarnings(selectedIds);
          setWarnings(prev => prev.filter(w => !selectedIds.includes(w.id)));
        } else {
          const itemsToDelete = terminations
            .filter(t => selectedIds.includes(t.id))
            .map(t => ({ id: t.id, account_id: t.account_id }));
          await disciplineService.bulkDeleteTerminations(itemsToDelete);
          setTerminations(prev => prev.filter(t => !selectedIds.includes(t.id)));
        }
        setSelectedIds([]);
        Swal.fire('Berhasil!', 'Data terpilih telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Beberapa data mungkin gagal dihapus.', 'error');
        fetchData();
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
    const currentList = activeTab === 'warnings' ? filteredWarnings : filteredTerminations;
    if (selectedIds.length === currentList.length) {
      setSelectedIds([]);
    } else {
      setSelectedIds(currentList.map(item => item.id));
    }
  };

  const handleTabChange = (tab: 'warnings' | 'terminations') => {
    setActiveTab(tab);
    setSelectedIds([]);
  };

  const filteredWarnings = warnings.filter(w => 
    `${w.account?.full_name} ${w.account?.internal_nik} ${w.warning_type} ${w.reason}`.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const filteredTerminations = terminations.filter(t => 
    `${t.account?.full_name} ${t.account?.internal_nik} ${t.termination_type} ${t.reason}`.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const formatDate = (dateStr: string) => {
    return new Date(dateStr).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' });
  };

  const formatCurrency = (amount?: number) => {
    return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(amount || 0);
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex bg-gray-50 p-1 rounded-md border border-gray-100">
          <button 
            onClick={() => handleTabChange('warnings')}
            className={`px-4 py-2 text-xs font-bold uppercase transition-all rounded ${activeTab === 'warnings' ? 'bg-white text-[#006E62] shadow-sm' : 'text-gray-400'}`}
          >
            Peringatan (SP)
          </button>
          <button 
            onClick={() => handleTabChange('terminations')}
            className={`px-4 py-2 text-xs font-bold uppercase transition-all rounded ${activeTab === 'terminations' ? 'bg-white text-red-600 shadow-sm' : 'text-gray-400'}`}
          >
            Karyawan Keluar (Exit)
          </button>
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
          {activeTab === 'warnings' ? (
            <>
              <button onClick={() => disciplineService.downloadWarningTemplate()} className="flex items-center gap-2 border border-gray-200 px-4 py-2 rounded-md hover:bg-gray-50 text-xs font-bold text-gray-500 uppercase">
                <Download size={16} /> Template SP
              </button>
              <button onClick={() => { setImportType('warning'); setShowImportModal(true); }} className="flex items-center gap-2 bg-[#006E62] text-white px-4 py-2 rounded-md hover:bg-[#005a50] text-xs font-bold uppercase">
                <FileUp size={16} /> Impor SP
              </button>
            </>
          ) : (
            <>
              <button onClick={() => disciplineService.downloadTerminationTemplate()} className="flex items-center gap-2 border border-gray-200 px-4 py-2 rounded-md hover:bg-gray-50 text-xs font-bold text-gray-500 uppercase">
                <Download size={16} /> Template Exit
              </button>
              <button onClick={() => { setImportType('termination'); setShowImportModal(true); }} className="flex items-center gap-2 bg-red-600 text-white px-4 py-2 rounded-md hover:bg-red-700 text-xs font-bold uppercase">
                <FileUp size={16} /> Impor Exit
              </button>
            </>
          )}
        </div>
      </div>

      <div className="relative max-w-md">
        <ShieldAlert className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
        <input
          type="text"
          placeholder="Cari data..."
          className="w-full pl-10 pr-4 py-2 bg-gray-50 border border-gray-200 rounded-md focus:outline-none focus:ring-1 focus:ring-[#006E62] text-sm"
          value={searchTerm}
          onChange={(e) => setSearchTerm(e.target.value)}
        />
      </div>

      <div className="bg-white border border-gray-100 rounded-md overflow-hidden shadow-sm">
        {activeTab === 'warnings' ? (
          <table className="w-full text-left">
            <thead className="bg-gray-50 text-[10px] font-bold text-gray-500 uppercase tracking-widest border-b border-gray-100">
              <tr>
                <th className="px-6 py-4 w-10">
                  <input 
                    type="checkbox" 
                    className="rounded border-gray-300 text-[#006E62] focus:ring-[#006E62]"
                    checked={selectedIds.length === filteredWarnings.length && filteredWarnings.length > 0}
                    onChange={toggleSelectAll}
                  />
                </th>
                <th className="px-6 py-4">Karyawan</th>
                <th className="px-6 py-4">Jenis Peringatan</th>
                <th className="px-6 py-4">Alasan</th>
                <th className="px-6 py-4">Tgl Terbit</th>
                <th className="px-6 py-4">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {isLoading ? <tr><td colSpan={6} className="py-20 text-center text-gray-400">Memuat...</td></tr> : filteredWarnings.map(w => {
                const isSelected = selectedIds.includes(w.id);
                return (
                  <tr 
                    key={w.id} 
                    className={`hover:bg-gray-50/50 transition-colors cursor-pointer ${isSelected ? 'bg-emerald-50/20' : ''}`}
                    onClick={() => setSelectedWarning(w)}
                  >
                    <td className="px-6 py-4" onClick={(e) => e.stopPropagation()}>
                      <input 
                        type="checkbox" 
                        className="rounded border-gray-300 text-[#006E62] focus:ring-[#006E62]"
                        checked={isSelected}
                        onChange={() => toggleSelect(w.id)}
                      />
                    </td>
                    <td className="px-6 py-4">
                      <div className="flex items-center gap-3">
                        <div className="w-8 h-8 rounded-full bg-gray-100 text-gray-400 flex items-center justify-center border border-gray-200 overflow-hidden">
                          {w.account?.photo_google_id ? (
                            <img 
                              src={googleDriveService.getFileUrl(w.account.photo_google_id)} 
                              alt="" 
                              className="w-full h-full object-cover"
                              referrerPolicy="no-referrer"
                            />
                          ) : (
                            <UserCircle size={20} />
                          )}
                        </div>
                        <div>
                          <div className="text-xs font-bold text-gray-800">{w.account?.full_name}</div>
                          <div className="text-[10px] text-gray-400 font-mono">{w.account?.internal_nik}</div>
                        </div>
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-orange-50 text-orange-600 uppercase">{w.warning_type}</span>
                    </td>
                    <td className="px-6 py-4 text-xs text-gray-600 max-w-xs truncate">{w.reason}</td>
                    <td className="px-6 py-4 text-xs font-bold text-gray-500">{formatDate(w.issue_date)}</td>
                    <td className="px-6 py-4" onClick={(e) => e.stopPropagation()}>
                      <div className="flex items-center gap-2">
                        {w.file_id && (
                          <a 
                            href={googleDriveService.getFileUrl(w.file_id, true)} 
                            target="_blank" 
                            rel="noreferrer"
                            className="p-1.5 text-[#006E62] hover:bg-emerald-50 rounded transition-colors"
                            title="Lihat Dokumen"
                          >
                            <Paperclip size={14} />
                          </a>
                        )}
                        <button 
                          onClick={() => setEditingWarning(w)}
                          className="p-1.5 text-[#006E62] hover:bg-emerald-50 rounded transition-colors"
                          title="Edit Peringatan"
                        >
                          <Edit2 size={14} />
                        </button>
                        <button 
                          onClick={() => handleDeleteWarning(w.id)} 
                          className="p-1.5 text-red-400 hover:bg-red-50 rounded transition-colors"
                          title="Hapus Peringatan"
                        >
                          <Trash2 size={14} />
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        ) : (
          <table className="w-full text-left">
            <thead className="bg-gray-50 text-[10px] font-bold text-gray-500 uppercase tracking-widest border-b border-gray-100">
              <tr>
                <th className="px-6 py-4 w-10">
                  <input 
                    type="checkbox" 
                    className="rounded border-gray-300 text-[#006E62] focus:ring-[#006E62]"
                    checked={selectedIds.length === filteredTerminations.length && filteredTerminations.length > 0}
                    onChange={toggleSelectAll}
                  />
                </th>
                <th className="px-6 py-4">Karyawan</th>
                <th className="px-6 py-4">Tipe Exit</th>
                <th className="px-6 py-4">Tgl Berhenti</th>
                <th className="px-6 py-4">Alasan</th>
                <th className="px-6 py-4">Keuangan</th>
                <th className="px-6 py-4">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {isLoading ? <tr><td colSpan={7} className="py-20 text-center text-gray-400">Memuat...</td></tr> : filteredTerminations.map(t => {
                const isSelected = selectedIds.includes(t.id);
                return (
                  <tr 
                    key={t.id} 
                    className={`hover:bg-gray-50/50 transition-colors cursor-pointer border-l-2 border-l-red-500 ${isSelected ? 'bg-red-50/20' : 'bg-red-50/5'}`}
                    onClick={() => setSelectedTermination(t)}
                  >
                    <td className="px-6 py-4" onClick={(e) => e.stopPropagation()}>
                      <input 
                        type="checkbox" 
                        className="rounded border-gray-300 text-[#006E62] focus:ring-[#006E62]"
                        checked={isSelected}
                        onChange={() => toggleSelect(t.id)}
                      />
                    </td>
                    <td className="px-6 py-4">
                      <div className="flex items-center gap-3">
                        <div className="w-8 h-8 rounded-full bg-gray-100 text-gray-400 flex items-center justify-center border border-gray-200 overflow-hidden">
                          {t.account?.photo_google_id ? (
                            <img 
                              src={googleDriveService.getFileUrl(t.account.photo_google_id)} 
                              alt="" 
                              className="w-full h-full object-cover"
                              referrerPolicy="no-referrer"
                            />
                          ) : (
                            <UserCircle size={20} />
                          )}
                        </div>
                        <div>
                          <div className="text-xs font-bold text-gray-800">{t.account?.full_name}</div>
                          <div className="text-[10px] text-gray-400 font-mono">{t.account?.internal_nik}</div>
                        </div>
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-red-50 text-red-600 uppercase">{t.termination_type}</span>
                    </td>
                    <td className="px-6 py-4 text-xs font-bold text-gray-500">{formatDate(t.termination_date)}</td>
                    <td className="px-6 py-4 text-xs text-gray-600 max-w-xs truncate italic">"{t.reason}"</td>
                    <td className="px-6 py-4">
                      <div className="text-[10px] font-bold text-gray-700">
                        {t.termination_type === 'Pemecatan' ? `Pesangon: ${formatCurrency(t.severance_amount)}` : `Penalti: ${formatCurrency(t.penalty_amount)}`}
                      </div>
                    </td>
                    <td className="px-6 py-4" onClick={(e) => e.stopPropagation()}>
                      <div className="flex items-center gap-2">
                        {t.file_id && (
                          <a 
                            href={googleDriveService.getFileUrl(t.file_id, true)} 
                            target="_blank" 
                            rel="noreferrer"
                            className="p-1.5 text-[#006E62] hover:bg-emerald-50 rounded transition-colors"
                            title="Lihat Dokumen"
                          >
                            <Paperclip size={14} />
                          </a>
                        )}
                        <button 
                          onClick={() => setEditingTermination(t)}
                          className="p-1.5 text-[#006E62] hover:bg-emerald-50 rounded transition-colors"
                          title="Edit Pengakhiran"
                        >
                          <Edit2 size={14} />
                        </button>
                        <button 
                          onClick={() => handleDeleteTermination(t.id, t.account_id)} 
                          className="p-1.5 text-red-400 hover:bg-red-50 rounded transition-colors"
                          title="Hapus Pengakhiran"
                        >
                          <Trash2 size={14} />
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>

      {showImportModal && (
        <DisciplineImportModal 
          type={importType}
          onClose={() => setShowImportModal(false)} 
          onSuccess={() => { setShowImportModal(false); fetchData(); }} 
        />
      )}

      {/* Warning Detail Modal */}
      {selectedWarning && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-lg overflow-hidden animate-in fade-in zoom-in duration-200">
            <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
              <div className="flex items-center gap-2 text-orange-600">
                <Info size={20} />
                <h3 className="font-bold text-gray-800">Detail Surat Peringatan</h3>
              </div>
              <button onClick={() => setSelectedWarning(null)} className="p-2 hover:bg-gray-200 rounded-full transition-colors">
                <X size={20} className="text-gray-500" />
              </button>
            </div>
            <div className="p-6 space-y-6">
              <div className="flex items-center gap-4 p-4 bg-orange-50/30 rounded-lg border border-orange-100/50">
                <div className="w-12 h-12 rounded-full bg-gray-100 flex items-center justify-center border border-gray-200 text-gray-400 shadow-sm overflow-hidden">
                  {selectedWarning.account?.photo_google_id ? (
                    <img 
                      src={googleDriveService.getFileUrl(selectedWarning.account.photo_google_id)} 
                      alt="" 
                      className="w-full h-full object-cover"
                      referrerPolicy="no-referrer"
                    />
                  ) : (
                    <UserCircle size={32} />
                  )}
                </div>
                <div>
                  <h4 className="font-bold text-gray-900">{selectedWarning.account?.full_name}</h4>
                  <p className="text-xs font-mono text-gray-500 uppercase tracking-wider">{selectedWarning.account?.internal_nik}</p>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-6">
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Jenis Peringatan</p>
                  <p className="text-sm font-bold text-orange-600">{selectedWarning.warning_type}</p>
                </div>
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tanggal Terbit</p>
                  <p className="text-sm font-bold text-gray-800">{formatDate(selectedWarning.issue_date)}</p>
                </div>
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Berlaku Sampai</p>
                  <p className="text-sm font-medium text-gray-700">{selectedWarning.expiry_date ? formatDate(selectedWarning.expiry_date) : '-'}</p>
                </div>
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tanggal Input</p>
                  <p className="text-sm font-bold text-gray-700">{formatDate(selectedWarning.entry_date)}</p>
                </div>
              </div>

              <div className="space-y-2 pt-4 border-t border-gray-100">
                <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Alasan Peringatan</p>
                <p className="text-sm text-gray-600 leading-relaxed bg-gray-50 p-3 rounded-lg border border-gray-100 italic">
                  {selectedWarning.reason || 'Tidak ada alasan dicantumkan.'}
                </p>
              </div>

              {selectedWarning.file_id && (
                <div className="pt-4">
                  <a 
                    href={googleDriveService.getFileUrl(selectedWarning.file_id, true)}
                    target="_blank"
                    rel="noreferrer"
                    className="flex items-center justify-center gap-2 w-full py-3 bg-orange-50 text-orange-600 rounded-lg font-bold text-sm hover:bg-orange-100 transition-all border border-orange-200"
                  >
                    <Paperclip size={18} /> LIHAT DOKUMEN SP
                  </a>
                </div>
              )}
            </div>
            <div className="px-6 py-4 bg-gray-50 border-t border-gray-100 flex justify-end">
              <button 
                onClick={() => setSelectedWarning(null)}
                className="px-6 py-2 bg-white border border-gray-200 text-gray-600 rounded-lg text-sm font-bold hover:bg-gray-100 transition-all"
              >
                Tutup
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Warning Edit Modal */}
      {editingWarning && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-md overflow-hidden animate-in fade-in zoom-in duration-200">
            <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
              <div className="flex items-center gap-2 text-orange-600">
                <Edit2 size={20} />
                <h3 className="font-bold text-gray-800">Edit Surat Peringatan</h3>
              </div>
              <button onClick={() => setEditingWarning(null)} className="p-2 hover:bg-gray-200 rounded-full transition-colors">
                <X size={20} className="text-gray-500" />
              </button>
            </div>
            <form onSubmit={async (e) => {
              e.preventDefault();
              const formData = new FormData(e.currentTarget);
              const file = (formData.get('file_sp') as File);
              
              const data: any = {
                warning_type: formData.get('warning_type') as string,
                issue_date: formData.get('issue_date') as string,
                expiry_date: formData.get('expiry_date') as string,
                reason: formData.get('reason') as string,
              };
              
              try {
                setIsLoading(true);
                
                if (file && file.size > 0) {
                  if (editingWarning.file_id) {
                    await googleDriveService.deleteFile(editingWarning.file_id);
                  }
                  const newFileId = await googleDriveService.uploadFile(file);
                  data.file_id = newFileId;
                }

                await disciplineService.updateWarning(editingWarning.id, data as any);
                setWarnings(prev => prev.map(w => w.id === editingWarning.id ? { ...w, ...data } : w));
                setEditingWarning(null);
                Swal.fire({ title: 'Berhasil!', text: 'Data peringatan telah diperbarui.', icon: 'success', timer: 1500, showConfirmButton: false });
              } catch (error) {
                Swal.fire('Gagal', 'Gagal memperbarui data peringatan', 'error');
              } finally {
                setIsLoading(false);
              }
            }} className="p-6 space-y-4">
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Jenis Peringatan</label>
                <select 
                  name="warning_type"
                  defaultValue={editingWarning.warning_type}
                  required
                  className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-orange-500 text-sm font-medium"
                >
                  <option value="SP 1">SP 1</option>
                  <option value="SP 2">SP 2</option>
                  <option value="SP 3">SP 3</option>
                  <option value="Teguran Lisan">Teguran Lisan</option>
                  <option value="Teguran Tertulis">Teguran Tertulis</option>
                </select>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1">
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tgl Terbit</label>
                  <input 
                    type="date"
                    name="issue_date"
                    defaultValue={editingWarning.issue_date}
                    required
                    className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-orange-500 text-sm font-medium"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tgl Berakhir</label>
                  <input 
                    type="date"
                    name="expiry_date"
                    defaultValue={editingWarning.expiry_date || ''}
                    className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-orange-500 text-sm font-medium"
                  />
                </div>
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Alasan</label>
                <textarea 
                  name="reason"
                  defaultValue={editingWarning.reason || ''}
                  rows={3}
                  className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-orange-500 text-sm font-medium resize-none"
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Ganti Dokumen SP (Opsional)</label>
                <input 
                  type="file"
                  name="file_sp"
                  accept="image/*,application/pdf"
                  className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-orange-500 text-sm font-medium"
                />
                {editingWarning.file_id && <p className="text-[10px] text-orange-500 font-medium italic">* Mengunggah file baru akan menghapus file lama.</p>}
              </div>
              <div className="pt-4 flex gap-3">
                <button 
                  type="button"
                  onClick={() => setEditingWarning(null)}
                  className="flex-1 py-2.5 bg-white border border-gray-200 text-gray-600 rounded-lg text-sm font-bold hover:bg-gray-100 transition-all"
                >
                  Batal
                </button>
                <button 
                  type="submit"
                  className="flex-1 py-2.5 bg-orange-600 text-white rounded-lg text-sm font-bold hover:bg-orange-700 transition-all shadow-md shadow-orange-100"
                >
                  Simpan Perubahan
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Termination Detail Modal */}
      {selectedTermination && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-lg overflow-hidden animate-in fade-in zoom-in duration-200">
            <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
              <div className="flex items-center gap-2 text-red-600">
                <Info size={20} />
                <h3 className="font-bold text-gray-800">Detail Pengakhiran Kerja</h3>
              </div>
              <button onClick={() => setSelectedTermination(null)} className="p-2 hover:bg-gray-200 rounded-full transition-colors">
                <X size={20} className="text-gray-500" />
              </button>
            </div>
            <div className="p-6 space-y-6">
              <div className="flex items-center gap-4 p-4 bg-red-50/30 rounded-lg border border-red-100/50">
                <div className="w-12 h-12 rounded-full bg-gray-100 flex items-center justify-center border border-gray-200 text-gray-400 shadow-sm overflow-hidden">
                  {selectedTermination.account?.photo_google_id ? (
                    <img 
                      src={googleDriveService.getFileUrl(selectedTermination.account.photo_google_id)} 
                      alt="" 
                      className="w-full h-full object-cover"
                      referrerPolicy="no-referrer"
                    />
                  ) : (
                    <UserCircle size={32} />
                  )}
                </div>
                <div>
                  <h4 className="font-bold text-gray-900">{selectedTermination.account?.full_name}</h4>
                  <p className="text-xs font-mono text-gray-500 uppercase tracking-wider">{selectedTermination.account?.internal_nik}</p>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-6">
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tipe Exit</p>
                  <p className="text-sm font-bold text-red-600">{selectedTermination.termination_type}</p>
                </div>
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tanggal Berhenti</p>
                  <p className="text-sm font-bold text-gray-800">{formatDate(selectedTermination.termination_date)}</p>
                </div>
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Pesangon</p>
                  <p className="text-sm font-bold text-gray-700">{formatCurrency(selectedTermination.severance_amount)}</p>
                </div>
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Penalti</p>
                  <p className="text-sm font-bold text-gray-700">{formatCurrency(selectedTermination.penalty_amount)}</p>
                </div>
              </div>

              <div className="space-y-2 pt-4 border-t border-gray-100">
                <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Alasan Pengakhiran</p>
                <p className="text-sm text-gray-600 leading-relaxed bg-gray-50 p-3 rounded-lg border border-gray-100 italic">
                  {selectedTermination.reason || 'Tidak ada alasan dicantumkan.'}
                </p>
              </div>

              {selectedTermination.file_id && (
                <div className="pt-4">
                  <a 
                    href={googleDriveService.getFileUrl(selectedTermination.file_id, true)}
                    target="_blank"
                    rel="noreferrer"
                    className="flex items-center justify-center gap-2 w-full py-3 bg-red-50 text-red-600 rounded-lg font-bold text-sm hover:bg-red-100 transition-all border border-red-200"
                  >
                    <Paperclip size={18} /> LIHAT DOKUMEN EXIT
                  </a>
                </div>
              )}
            </div>
            <div className="px-6 py-4 bg-gray-50 border-t border-gray-100 flex justify-end">
              <button 
                onClick={() => setSelectedTermination(null)}
                className="px-6 py-2 bg-white border border-gray-200 text-gray-600 rounded-lg text-sm font-bold hover:bg-gray-100 transition-all"
              >
                Tutup
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Termination Edit Modal */}
      {editingTermination && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-md overflow-hidden animate-in fade-in zoom-in duration-200">
            <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
              <div className="flex items-center gap-2 text-red-600">
                <Edit2 size={20} />
                <h3 className="font-bold text-gray-800">Edit Pengakhiran Kerja</h3>
              </div>
              <button onClick={() => setEditingTermination(null)} className="p-2 hover:bg-gray-200 rounded-full transition-colors">
                <X size={20} className="text-gray-500" />
              </button>
            </div>
            <form onSubmit={async (e) => {
              e.preventDefault();
              const formData = new FormData(e.currentTarget);
              const file = (formData.get('file_exit') as File);
              
              const data: any = {
                termination_type: formData.get('termination_type') as string,
                termination_date: formData.get('termination_date') as string,
                reason: formData.get('reason') as string,
                severance_amount: Number(formData.get('severance_amount')),
                penalty_amount: Number(formData.get('penalty_amount')),
              };
              
              try {
                setIsLoading(true);
                
                if (file && file.size > 0) {
                  if (editingTermination.file_id) {
                    await googleDriveService.deleteFile(editingTermination.file_id);
                  }
                  const newFileId = await googleDriveService.uploadFile(file);
                  data.file_id = newFileId;
                }

                await disciplineService.updateTermination(editingTermination.id, data as any);
                setTerminations(prev => prev.map(t => t.id === editingTermination.id ? { ...t, ...data } : t));
                setEditingTermination(null);
                Swal.fire({ title: 'Berhasil!', text: 'Data pengakhiran telah diperbarui.', icon: 'success', timer: 1500, showConfirmButton: false });
              } catch (error) {
                Swal.fire('Gagal', 'Gagal memperbarui data pengakhiran', 'error');
              } finally {
                setIsLoading(false);
              }
            }} className="p-6 space-y-4">
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tipe Exit</label>
                <select 
                  name="termination_type"
                  defaultValue={editingTermination.termination_type}
                  required
                  className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-red-500 text-sm font-medium"
                >
                  <option value="Resign">Resign</option>
                  <option value="Pemecatan">Pemecatan</option>
                  <option value="Pensiun">Pensiun</option>
                  <option value="Habis Kontrak">Habis Kontrak</option>
                  <option value="Lainnya">Lainnya</option>
                </select>
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tanggal Berhenti</label>
                <input 
                  type="date"
                  name="termination_date"
                  defaultValue={editingTermination.termination_date}
                  required
                  className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-red-500 text-sm font-medium"
                />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1">
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Pesangon (Rp)</label>
                  <input 
                    type="number"
                    name="severance_amount"
                    defaultValue={editingTermination.severance_amount}
                    className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-red-500 text-sm font-medium"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Penalti (Rp)</label>
                  <input 
                    type="number"
                    name="penalty_amount"
                    defaultValue={editingTermination.penalty_amount}
                    className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-red-500 text-sm font-medium"
                  />
                </div>
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Alasan</label>
                <textarea 
                  name="reason"
                  defaultValue={editingTermination.reason || ''}
                  rows={3}
                  className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-red-500 text-sm font-medium resize-none"
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Ganti Dokumen Exit (Opsional)</label>
                <input 
                  type="file"
                  name="file_exit"
                  accept="image/*,application/pdf"
                  className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-red-500 text-sm font-medium"
                />
                {editingTermination.file_id && <p className="text-[10px] text-orange-500 font-medium italic">* Mengunggah file baru akan menghapus file lama.</p>}
              </div>
              <div className="pt-4 flex gap-3">
                <button 
                  type="button"
                  onClick={() => setEditingTermination(null)}
                  className="flex-1 py-2.5 bg-white border border-gray-200 text-gray-600 rounded-lg text-sm font-bold hover:bg-gray-100 transition-all"
                >
                  Batal
                </button>
                <button 
                  type="submit"
                  className="flex-1 py-2.5 bg-red-600 text-white rounded-lg text-sm font-bold hover:bg-red-700 transition-all shadow-md shadow-red-100"
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

export default DisciplineMain;