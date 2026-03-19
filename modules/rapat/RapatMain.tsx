import React, { useState, useEffect } from 'react';
import { Video, Plus, Search, Calendar, Clock, MapPin, Users, Eye, Trash2, Play, CheckCircle2, AlertCircle, FileText, ExternalLink } from 'lucide-react';
import Swal from 'sweetalert2';
import { meetingService } from '../../services/meetingService';
import { authService } from '../../services/authService';
import { accountService } from '../../services/accountService';
import { Meeting, AuthUser, Account } from '../../types';
import MeetingForm from './MeetingForm';
import MeetingSession from './MeetingSession';
import MeetingDetail from './MeetingDetail';
import LoadingSpinner from '../../components/Common/LoadingSpinner';
import { CardSkeleton } from '../../components/Common/Skeleton';

const RapatMain: React.FC = () => {
  const [meetings, setMeetings] = useState<Meeting[]>([]);
  const [accounts, setAccounts] = useState<Account[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [showForm, setShowForm] = useState(false);
  const [selectedMeeting, setSelectedMeeting] = useState<Meeting | null>(null);
  const [showSession, setShowSession] = useState(false);
  const [showDetail, setShowDetail] = useState(false);
  const [user, setUser] = useState<AuthUser | null>(null);
  const isAdmin = user?.role === 'admin' || user?.is_hr_admin || user?.is_performance_admin || user?.is_finance_admin;

  useEffect(() => {
    const currentUser = authService.getCurrentUser();
    setUser(currentUser);
    fetchData();
    fetchAccounts();
  }, []);

  const fetchData = async () => {
    try {
      setIsLoading(true);
      const data = await meetingService.getAll();
      setMeetings(data);
    } catch (error) {
      console.error(error);
    } finally {
      setIsLoading(false);
    }
  };

  const fetchAccounts = async () => {
    try {
      const data = await accountService.getAll();
      setAccounts(data || []);
    } catch (error) {
      console.error('Failed to fetch accounts:', error);
    }
  };

  const handleCreateMeeting = async (input: any) => {
    if (!user) return;
    try {
      setIsSaving(true);
      await meetingService.create({
        ...input,
        created_by: user.id,
        status: 'Scheduled'
      });
      await fetchData();
      setShowForm(false);
      Swal.fire('Berhasil', 'Rapat telah dijadwalkan.', 'success');
    } catch (error) {
      Swal.fire('Gagal', 'Terjadi kesalahan saat menjadwalkan rapat.', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleDelete = async (id: string) => {
    const result = await Swal.fire({
      title: 'Hapus Rapat?',
      text: "Data rapat akan dihapus permanen.",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
      confirmButtonText: 'Ya, Hapus!'
    });

    if (result.isConfirmed) {
      try {
        setIsSaving(true);
        await meetingService.delete(id);
        await fetchData();
        Swal.fire('Terhapus!', 'Rapat telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Terjadi kesalahan saat menghapus rapat.', 'error');
      } finally {
        setIsSaving(false);
      }
    }
  };

  const handleStartMeeting = async (meeting: Meeting) => {
    const result = await Swal.fire({
      title: 'Mulai Rapat?',
      text: "Sesi rapat akan dimulai sekarang.",
      icon: 'question',
      showCancelButton: true,
      confirmButtonText: 'Ya, Mulai!'
    });

    if (result.isConfirmed) {
      try {
        setIsSaving(true);
        await meetingService.startMeeting(meeting.id);
        await fetchData();
        setSelectedMeeting({ ...meeting, status: 'In Progress', started_at: new Date().toISOString() });
        setShowSession(true);
      } catch (error) {
        Swal.fire('Gagal', 'Gagal memulai rapat.', 'error');
      } finally {
        setIsSaving(false);
      }
    }
  };

  const handleEndMeeting = async () => {
    if (!selectedMeeting) return;
    try {
      setIsSaving(true);
      await meetingService.endMeeting(selectedMeeting.id);
      await fetchData();
      setShowSession(false);
      setSelectedMeeting(null);
      Swal.fire('Selesai', 'Rapat telah diakhiri dan notulensi disimpan.', 'success');
    } catch (error) {
      Swal.fire('Gagal', 'Gagal mengakhiri rapat.', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const filteredMeetings = meetings.filter(m => {
    const matchesSearch = m.title.toLowerCase().includes(searchTerm.toLowerCase()) ||
                         m.description.toLowerCase().includes(searchTerm.toLowerCase());
    
    // Superadmin (SP) can see all meetings
    if (user?.access_code?.startsWith('SP')) {
      return matchesSearch;
    }

    // Others only see meetings where they are participants, notulens, or the creator
    const isInvolved = m.participant_ids.includes(user?.id || '') || 
                      m.notulen_ids.includes(user?.id || '') || 
                      m.created_by === user?.id;
    
    return matchesSearch && isInvolved;
  });

  const canManageSession = (meeting: Meeting) => {
    return user && (meeting.notulen_ids.includes(user.id) || isAdmin);
  };

  return (
    <div className="space-y-6">
      {isSaving && <LoadingSpinner message="Memproses Data..." />}

      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-xl font-bold text-gray-800 tracking-tight">Notulensi Rapat</h2>
          <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Manajemen Penjadwalan & Dokumentasi Rapat</p>
        </div>
        
        <div className="flex items-center gap-3">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
            <input
              type="text"
              placeholder="Cari rapat..."
              className="pl-9 pr-4 py-2 bg-gray-50 border border-gray-100 rounded-md focus:outline-none focus:ring-2 focus:ring-[#006E62] text-xs w-64"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
          <button 
            onClick={() => setShowForm(true)}
            className="flex items-center gap-2 bg-[#006E62] text-white px-4 py-2 rounded-md hover:bg-[#005a50] transition-colors shadow-sm text-xs font-bold uppercase"
          >
            <Plus size={16} />
            Jadwalkan Rapat
          </button>
        </div>
      </div>

      {isLoading ? (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {[1, 2, 3].map(i => <CardSkeleton key={i} />)}
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {filteredMeetings.map(meeting => (
            <div key={meeting.id} className="bg-white border border-gray-100 p-5 rounded-2xl shadow-sm hover:shadow-md transition-all border-l-4 border-l-transparent hover:border-l-[#006E62] group flex flex-col h-full">
              <div className="flex justify-between items-start mb-4">
                <div className="flex flex-col">
                  <span className={`px-2 py-0.5 rounded text-[9px] font-bold uppercase mb-1 w-fit ${
                    meeting.status === 'Scheduled' ? 'bg-blue-50 text-blue-600' :
                    meeting.status === 'In Progress' ? 'bg-amber-50 text-amber-600 animate-pulse' :
                    meeting.status === 'Completed' ? 'bg-emerald-50 text-emerald-600' :
                    'bg-gray-50 text-gray-600'
                  }`}>
                    {meeting.status}
                  </span>
                  <h4 className="text-sm font-bold text-gray-800 leading-tight group-hover:text-[#006E62] transition-colors">{meeting.title}</h4>
                </div>
                <div className="flex items-center gap-1">
                  {meeting.status === 'Completed' ? (
                    <button 
                      onClick={() => { setSelectedMeeting(meeting); setShowDetail(true); }}
                      className="p-1.5 text-gray-400 hover:text-[#006E62] hover:bg-emerald-50 rounded-lg transition-all"
                      title="Lihat Notulensi"
                    >
                      <Eye size={14} />
                    </button>
                  ) : meeting.status === 'In Progress' && canManageSession(meeting) ? (
                    <button 
                      onClick={() => { setSelectedMeeting(meeting); setShowSession(true); }}
                      className="p-1.5 text-amber-600 hover:bg-amber-50 rounded-lg transition-all"
                      title="Lanjutkan Sesi"
                    >
                      <Play size={14} />
                    </button>
                  ) : meeting.status === 'Scheduled' && canManageSession(meeting) ? (
                    <button 
                      onClick={() => handleStartMeeting(meeting)}
                      className="p-1.5 text-blue-600 hover:bg-blue-50 rounded-lg transition-all"
                      title="Mulai Rapat"
                    >
                      <Play size={14} />
                    </button>
                  ) : null}
                  {(isAdmin || meeting.created_by === user?.id) && (
                    <button 
                      onClick={() => handleDelete(meeting.id)}
                      className="p-1.5 text-gray-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-all"
                      title="Hapus"
                    >
                      <Trash2 size={14} />
                    </button>
                  )}
                </div>
              </div>

              <div className="space-y-2 mb-4">
                <div className="flex items-center gap-2 text-[10px] text-gray-500 font-medium">
                  <Calendar size={12} className="text-gray-400" />
                  <span>{new Date(meeting.scheduled_at).toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' })}</span>
                </div>
                <div className="flex items-center gap-2 text-[10px] text-gray-500 font-medium">
                  <Clock size={12} className="text-gray-400" />
                  <span>{new Date(meeting.scheduled_at).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })} WIB</span>
                </div>
                <div className="flex items-center gap-2 text-[10px] text-gray-500 font-medium">
                  <MapPin size={12} className="text-gray-400" />
                  <span className="truncate">{meeting.location_detail}</span>
                </div>
                <div className="flex items-center gap-2 text-[10px] text-gray-500 font-medium">
                  <Users size={12} className="text-gray-400" />
                  <span>{meeting.participant_ids.length} Peserta</span>
                </div>
              </div>

              <p className="text-[11px] text-gray-500 italic line-clamp-2 mb-4 flex-1">"{meeting.description}"</p>
              
              <div className="pt-4 border-t border-gray-50 flex items-center justify-between mt-auto">
                <div className="flex items-center gap-1.5">
                  <div className="w-6 h-6 rounded-full bg-gray-100 flex items-center justify-center text-[10px] font-bold text-gray-500 border border-white">
                    {meeting.creator?.full_name?.charAt(0)}
                  </div>
                  <span className="text-[9px] font-bold text-gray-400 uppercase">Host: {meeting.creator?.full_name}</span>
                </div>
                {meeting.status === 'Completed' && (
                  <div className="flex items-center gap-1 text-emerald-600 font-bold text-[9px] uppercase">
                    <CheckCircle2 size={12} />
                    Selesai
                  </div>
                )}
              </div>
            </div>
          ))}

          {filteredMeetings.length === 0 && (
            <div className="col-span-full flex flex-col items-center justify-center py-20 text-gray-400 bg-gray-50 rounded-xl border border-dashed border-gray-200">
              <Video size={48} strokeWidth={1} className="mb-4 opacity-20" />
              <p className="text-sm font-bold uppercase tracking-widest">Belum ada jadwal rapat</p>
            </div>
          )}
        </div>
      )}

      {showForm && (
        <MeetingForm 
          onClose={() => setShowForm(false)} 
          onSubmit={handleCreateMeeting}
          accounts={accounts}
        />
      )}

      {showSession && selectedMeeting && (
        <MeetingSession 
          meeting={selectedMeeting}
          onClose={() => { setShowSession(false); setSelectedMeeting(null); }}
          onEnd={handleEndMeeting}
        />
      )}

      {showDetail && selectedMeeting && (
        <MeetingDetail 
          meeting={selectedMeeting}
          onClose={() => { setShowDetail(false); setSelectedMeeting(null); }}
        />
      )}
    </div>
  );
};

export default RapatMain;
