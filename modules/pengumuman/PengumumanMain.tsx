import React, { useState, useEffect } from 'react';
import { Megaphone, Plus, Search, Filter, Loader2, AlertCircle, Calendar, FileText, Info, History, Trash2, Edit2, Eye } from 'lucide-react';
import { Announcement, Account } from '../../types';
import { announcementService } from '../../services/announcementService';
import AnnouncementCard from './AnnouncementCard';
import AnnouncementDetail from './AnnouncementDetail';
import AnnouncementForm from './AnnouncementForm';
import Swal from 'sweetalert2';

interface PengumumanMainProps {
  user: Account;
}

const PengumumanMain: React.FC<PengumumanMainProps> = ({ user }) => {
  const [announcements, setAnnouncements] = useState<Announcement[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [searchQuery, setSearchQuery] = useState('');
  const [activeTab, setActiveTab] = useState<'Active' | 'Manage'>('Active');
  
  const [selectedAnnouncement, setSelectedAnnouncement] = useState<Announcement | null>(null);
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [editingAnnouncement, setEditingAnnouncement] = useState<Announcement | undefined>(undefined);

  const isAdmin = user?.role === 'admin' || user?.is_hr_admin || user?.is_performance_admin || user?.is_finance_admin;

  useEffect(() => {
    fetchAnnouncements();
  }, [activeTab, user.id, user.department]);

  const fetchAnnouncements = async () => {
    try {
      setIsLoading(true);
      let data: Announcement[] = [];
      if (activeTab === 'Active') {
        data = await announcementService.getAnnouncements(user.id, user.department);
      } else {
        data = await announcementService.getAllAnnouncementsAdmin();
      }
      setAnnouncements(data);
    } catch (error) {
      console.error('Failed to fetch announcements:', error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleDelete = async (id: string) => {
    const result = await Swal.fire({
      title: 'Hapus Pengumuman?',
      text: "Data pengumuman dan riwayat baca akan dihapus permanen.",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
      confirmButtonText: 'Ya, Hapus!'
    });

    if (result.isConfirmed) {
      try {
        await announcementService.deleteAnnouncement(id);
        await fetchAnnouncements();
        Swal.fire('Berhasil', 'Pengumuman telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Error', 'Gagal menghapus pengumuman.', 'error');
      }
    }
  };

  const filteredAnnouncements = announcements.filter(ann => 
    ann.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
    ann.content.toLowerCase().includes(searchQuery.toLowerCase())
  );

  return (
    <div className="p-8 space-y-8 animate-in fade-in duration-500">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
        <div>
          <h2 className="text-3xl font-black text-gray-800 tracking-tight">Pengumuman</h2>
          <p className="text-xs font-bold text-gray-400 uppercase tracking-widest mt-1">Informasi & Kebijakan Perusahaan</p>
        </div>
        
        <div className="flex items-center gap-3">
          <div className="relative">
            <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-gray-300" size={18} />
            <input 
              type="text"
              placeholder="Cari pengumuman..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="pl-12 pr-6 py-3 bg-white border border-gray-100 rounded-2xl focus:outline-none focus:ring-4 focus:ring-amber-600/5 focus:border-amber-600 transition-all text-sm font-medium w-64 shadow-sm"
            />
          </div>
          {isAdmin && (
            <button 
              onClick={() => {
                setEditingAnnouncement(undefined);
                setIsFormOpen(true);
              }}
              className="bg-amber-600 text-white px-6 py-3 rounded-2xl font-bold text-sm flex items-center gap-2 hover:bg-amber-700 transition-all shadow-lg shadow-amber-600/20 active:scale-95"
            >
              <Plus size={18} /> Buat Pengumuman
            </button>
          )}
        </div>
      </div>

      {isAdmin && (
        <div className="flex gap-2 p-1 bg-gray-100 rounded-2xl w-fit">
          <button 
            onClick={() => setActiveTab('Active')}
            className={`px-6 py-2 rounded-xl text-[10px] font-bold uppercase tracking-widest transition-all ${activeTab === 'Active' ? 'bg-white text-amber-600 shadow-sm' : 'text-gray-400 hover:text-gray-600'}`}
          >
            Pengumuman Aktif
          </button>
          <button 
            onClick={() => setActiveTab('Manage')}
            className={`px-6 py-2 rounded-xl text-[10px] font-bold uppercase tracking-widest transition-all ${activeTab === 'Manage' ? 'bg-white text-amber-600 shadow-sm' : 'text-gray-400 hover:text-gray-600'}`}
          >
            Kelola & Riwayat
          </button>
        </div>
      )}

      {isLoading ? (
        <div className="flex flex-col items-center justify-center py-32">
          <Loader2 size={48} className="animate-spin text-amber-600 mb-4" />
          <p className="text-xs font-bold text-gray-400 uppercase tracking-widest animate-pulse">Memuat Pengumuman...</p>
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          {activeTab === 'Active' ? (
            filteredAnnouncements.length > 0 ? (
              filteredAnnouncements.map(ann => (
                <AnnouncementCard 
                  key={ann.id} 
                  announcement={ann} 
                  onClick={() => setSelectedAnnouncement(ann)} 
                />
              ))
            ) : (
              <div className="col-span-full py-32 flex flex-col items-center justify-center text-gray-400 bg-white rounded-[40px] border border-dashed border-gray-200">
                <Megaphone size={64} strokeWidth={1} className="mb-4 opacity-20" />
                <p className="text-xs font-bold uppercase tracking-widest">Tidak ada pengumuman aktif saat ini</p>
              </div>
            )
          ) : (
            <div className="col-span-full bg-white rounded-[40px] border border-gray-100 overflow-hidden shadow-sm">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-gray-50/50 border-b border-gray-100">
                    <th className="px-8 py-5 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Pengumuman</th>
                    <th className="px-8 py-5 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Kategori</th>
                    <th className="px-8 py-5 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Target</th>
                    <th className="px-8 py-5 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Periode Tayang</th>
                    <th className="px-8 py-5 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Dibaca</th>
                    <th className="px-8 py-5 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">Aksi</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-50">
                  {filteredAnnouncements.map(ann => (
                    <tr key={ann.id} className="hover:bg-gray-50/50 transition-colors group">
                      <td className="px-8 py-5">
                        <div className="flex items-center gap-4">
                          <div className="w-10 h-10 rounded-xl bg-amber-50 text-amber-600 flex items-center justify-center shrink-0">
                            <Megaphone size={18} />
                          </div>
                          <div>
                            <p className="text-sm font-bold text-gray-800">{ann.title}</p>
                            <p className="text-[10px] font-bold text-gray-400 uppercase tracking-tight">Oleh: {ann.creator?.full_name}</p>
                          </div>
                        </div>
                      </td>
                      <td className="px-8 py-5">
                        <span className={`text-[9px] font-bold uppercase tracking-widest px-2 py-1 rounded-lg ${ann.category === 'Urgent' ? 'bg-rose-50 text-rose-600' : ann.category === 'Event' ? 'bg-emerald-50 text-emerald-600' : 'bg-blue-50 text-blue-600'}`}>
                          {ann.category}
                        </span>
                      </td>
                      <td className="px-8 py-5">
                        <p className="text-[10px] font-bold text-gray-600 uppercase tracking-tight">{ann.target_type}</p>
                        {ann.target_type !== 'All' && (
                          <p className="text-[8px] font-bold text-gray-400 uppercase">{ann.target_ids.length} Target</p>
                        )}
                      </td>
                      <td className="px-8 py-5">
                        <div className="flex flex-col gap-1">
                          <p className="text-[9px] font-bold text-gray-600 uppercase tracking-tight">Mulai: {new Date(ann.publish_start).toLocaleDateString('id-ID')}</p>
                          <p className="text-[9px] font-bold text-gray-400 uppercase tracking-tight">Selesai: {new Date(ann.publish_end).toLocaleDateString('id-ID')}</p>
                        </div>
                      </td>
                      <td className="px-8 py-5 text-center">
                        <div className="inline-flex items-center gap-1.5 px-3 py-1 bg-emerald-50 text-emerald-600 rounded-full">
                          <Eye size={12} />
                          <span className="text-[10px] font-black">{ann.read_count}</span>
                        </div>
                      </td>
                      <td className="px-8 py-5 text-right">
                        <div className="flex items-center justify-end gap-2">
                          <button 
                            onClick={() => setSelectedAnnouncement(ann)}
                            className="p-2 text-gray-400 hover:text-amber-600 hover:bg-amber-50 rounded-xl transition-all"
                          >
                            <Eye size={16} />
                          </button>
                          <button 
                            onClick={() => {
                              setEditingAnnouncement(ann);
                              setIsFormOpen(true);
                            }}
                            className="p-2 text-gray-400 hover:text-blue-600 hover:bg-blue-50 rounded-xl transition-all"
                          >
                            <Edit2 size={16} />
                          </button>
                          <button 
                            onClick={() => handleDelete(ann.id)}
                            className="p-2 text-gray-400 hover:text-rose-600 hover:bg-rose-50 rounded-xl transition-all"
                          >
                            <Trash2 size={16} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {filteredAnnouncements.length === 0 && (
                <div className="py-20 flex flex-col items-center justify-center text-gray-400">
                  <History size={48} strokeWidth={1} className="mb-4 opacity-20" />
                  <p className="text-xs font-bold uppercase tracking-widest">Belum ada riwayat pengumuman</p>
                </div>
              )}
            </div>
          )}
        </div>
      )}

      {selectedAnnouncement && (
        <AnnouncementDetail 
          announcement={selectedAnnouncement}
          userId={user.id}
          isAdmin={isAdmin}
          onClose={() => setSelectedAnnouncement(null)}
          onRead={fetchAnnouncements}
        />
      )}

      {isFormOpen && (
        <AnnouncementForm 
          announcement={editingAnnouncement}
          userId={user.id}
          onClose={() => setIsFormOpen(false)}
          onSave={() => {
            setIsFormOpen(false);
            fetchAnnouncements();
          }}
        />
      )}
    </div>
  );
};

export default PengumumanMain;
