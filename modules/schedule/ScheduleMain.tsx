
import React, { useState, useEffect } from 'react';
import { Plus, Search, CalendarClock, Grid, List as ListIcon, Clock, Calendar } from 'lucide-react';
import Swal from 'sweetalert2';
import { scheduleService } from '../../services/scheduleService';
import { Schedule, ScheduleInput } from '../../types';
import ScheduleForm from './ScheduleForm';
import { CardSkeleton } from '../../components/Common/Skeleton';
import LoadingSpinner from '../../components/Common/LoadingSpinner';

const ScheduleMain: React.FC = () => {
  const [schedules, setSchedules] = useState<Schedule[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [showForm, setShowForm] = useState(false);
  const [editingSchedule, setEditingSchedule] = useState<Schedule | null>(null);

  useEffect(() => {
    fetchSchedules();
  }, []);

  const fetchSchedules = async () => {
    try {
      setIsLoading(true);
      const data = await scheduleService.getAll();
      setSchedules(data);
    } catch (error) {
      Swal.fire('Gagal', 'Gagal memuat data jadwal', 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleCreate = async (input: ScheduleInput) => {
    setIsSaving(true);
    try {
      await scheduleService.create(input);
      fetchSchedules();
      setShowForm(false);
      Swal.fire({
        title: 'Berhasil!',
        text: 'Jadwal baru telah ditambahkan.',
        icon: 'success',
        timer: 1500,
        showConfirmButton: false
      });
    } catch (error) {
      Swal.fire('Gagal', 'Gagal menyimpan jadwal', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleUpdate = async (id: string, input: Partial<ScheduleInput>) => {
    setIsSaving(true);
    try {
      await scheduleService.update(id, input);
      fetchSchedules();
      setEditingSchedule(null);
      setShowForm(false);
      Swal.fire({
        title: 'Terupdate!',
        text: 'Jadwal berhasil diperbarui.',
        icon: 'success',
        timer: 1500,
        showConfirmButton: false
      });
    } catch (error) {
      Swal.fire('Gagal', 'Gagal memperbarui jadwal', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleDelete = async (id: string) => {
    const result = await Swal.fire({
      title: 'Hapus Jadwal?',
      text: "Seluruh aturan jam dan relasi lokasi akan dihapus.",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#006E62',
      cancelButtonColor: '#ef4444',
      confirmButtonText: 'Ya, hapus!',
      cancelButtonText: 'Batal'
    });

    if (result.isConfirmed) {
      setIsSaving(true);
      try {
        await scheduleService.delete(id);
        setSchedules(prev => prev.filter(s => s.id !== id));
        Swal.fire('Terhapus!', 'Jadwal telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Gagal menghapus data', 'error');
      } finally {
        setIsSaving(false);
      }
    }
  };

  const filteredSchedules = schedules.filter(s => 
    s.name.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const getBadgeType = (type: number) => {
    switch(type) {
      case 1: return { label: 'Fixed', color: 'bg-emerald-50 text-emerald-600' };
      case 2: return { label: 'Shift', color: 'bg-blue-50 text-blue-600' };
      case 3: return { label: 'Libur', color: 'bg-rose-50 text-rose-600' };
      case 4: return { label: 'Khusus', color: 'bg-amber-50 text-amber-600' };
      default: return { label: 'Jadwal', color: 'bg-gray-50 text-gray-600' };
    }
  };

  return (
    <div className="space-y-6">
      {isSaving && <LoadingSpinner />}

      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="relative flex-1 max-w-md">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
          <input
            type="text"
            placeholder="Cari Nama Jadwal..."
            className="w-full pl-10 pr-4 py-2 bg-gray-50 border border-gray-200 rounded-md focus:outline-none focus:ring-2 focus:ring-[#006E62] focus:border-transparent transition-all text-sm"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </div>
        
        <div className="flex items-center gap-2">
          <button 
            onClick={() => { setEditingSchedule(null); setShowForm(true); }}
            className="flex items-center gap-2 bg-[#006E62] text-white px-4 py-2 rounded-md hover:bg-[#005a50] transition-colors shadow-sm"
          >
            <Plus size={18} />
            <span className="font-medium text-sm">Tambah Jadwal</span>
          </button>
        </div>
      </div>

      {isLoading ? (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
          {[1, 2, 3].map(i => <CardSkeleton key={i} />)}
        </div>
      ) : filteredSchedules.length === 0 ? (
        <div className="flex flex-col items-center justify-center py-20 text-gray-400">
          <CalendarClock size={48} strokeWidth={1} className="mb-4" />
          <p className="text-lg">Belum ada data jadwal.</p>
        </div>
      ) : (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
          {filteredSchedules.map(schedule => {
            const badge = getBadgeType(schedule.type);
            return (
              <div 
                key={schedule.id}
                className="bg-white border border-gray-100 p-5 rounded-md shadow-sm hover:shadow-md transition-all border-l-4 border-l-[#006E62] group"
              >
                <div className="flex justify-between items-start mb-3">
                  <h3 className="font-bold text-gray-800 text-sm group-hover:text-[#006E62]">{schedule.name}</h3>
                  <span className={`text-[10px] font-bold px-2 py-0.5 rounded uppercase ${badge.color}`}>{badge.label}</span>
                </div>
                
                <div className="space-y-2 mb-4">
                   {schedule.type !== 3 && (
                    <div className="flex items-center gap-2 text-[11px] text-gray-500">
                        <Clock size={12} className="text-gray-300" /> 
                        Toleransi: <span className="font-bold text-gray-700">{schedule.tolerance_checkin_minutes || 0} In / {schedule.tolerance_minutes || 0} Out</span>
                    </div>
                   )}
                   {schedule.type >= 3 && (
                      <div className="flex items-center gap-2 text-[11px] text-gray-500">
                        <Calendar size={12} className="text-gray-300" />
                        Periode: <span className="font-bold text-gray-700">{schedule.start_date} s/d {schedule.end_date}</span>
                      </div>
                   )}
                   <div className="text-[10px] text-gray-400 font-bold uppercase tracking-wider">
                      {(schedule as any).location_ids?.length || 0} Lokasi Terdaftar
                   </div>
                </div>

                <div className="flex gap-2 pt-3 border-t border-gray-50">
                   <button 
                    onClick={() => { setEditingSchedule(schedule); setShowForm(true); }}
                    className="flex-1 text-[10px] font-bold uppercase text-[#006E62] hover:bg-emerald-50 py-1.5 rounded transition-colors border border-emerald-100"
                   >
                     Ubah
                   </button>
                   <button 
                    onClick={() => handleDelete(schedule.id)}
                    className="flex-1 text-[10px] font-bold uppercase text-red-500 hover:bg-red-50 py-1.5 rounded transition-colors border border-red-100"
                   >
                     Hapus
                   </button>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {showForm && (
        <ScheduleForm 
          onClose={() => { setShowForm(false); setEditingSchedule(null); }}
          onSubmit={editingSchedule ? (data) => handleUpdate(editingSchedule.id, data) : handleCreate}
          initialData={editingSchedule || undefined}
        />
      )}
    </div>
  );
};

export default ScheduleMain;
