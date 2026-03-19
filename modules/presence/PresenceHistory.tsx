
import React from 'react';
import { Fingerprint, Clock, MapPin, ExternalLink, Calendar, MessageSquare } from 'lucide-react';
import { Attendance } from '../../types';
import { googleDriveService } from '../../services/googleDriveService';
import Swal from 'sweetalert2';

interface PresenceHistoryProps {
  logs: Attendance[];
  isLoading: boolean;
}

const PresenceHistory: React.FC<PresenceHistoryProps> = ({ logs, isLoading }) => {
  const formatDate = (dateStr: string) => {
    return new Date(dateStr).toLocaleDateString('id-ID', {
      weekday: 'long',
      day: 'numeric',
      month: 'long',
      year: 'numeric'
    });
  };

  const formatTime = (dateStr: string | null) => {
    if (!dateStr) return '--:--';
    return new Date(dateStr).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', hour12: false });
  };

  if (isLoading) return (
    <div className="flex flex-col items-center justify-center py-20 text-gray-400">
      <div className="w-8 h-8 border-4 border-[#006E62] border-t-transparent rounded-full animate-spin mb-4"></div>
      <p className="text-xs font-bold uppercase tracking-widest">Memuat Riwayat...</p>
    </div>
  );

  if (logs.length === 0) return (
    <div className="flex flex-col items-center justify-center py-32 text-gray-300">
      <Calendar size={48} strokeWidth={1} />
      <p className="text-sm font-bold uppercase mt-4">Belum Ada Riwayat Presensi</p>
    </div>
  );

  return (
    <div className="space-y-4 pb-20">
      {logs.map((log) => (
        <div key={log.id} className="bg-white rounded-xl border border-gray-100 p-5 shadow-sm hover:shadow-md transition-all group border-l-4 border-l-[#006E62]">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
            <div className="space-y-4 flex-1">
              <div className="flex items-center gap-2">
                <Calendar size={14} className="text-gray-400" />
                <h4 className="text-sm font-bold text-gray-800">{formatDate(log.created_at!)}</h4>
              </div>
              
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                {/* Info Datang */}
                <div className="p-3 bg-emerald-50/50 rounded-lg border border-emerald-100 space-y-2">
                  <p className="text-[8px] font-bold text-emerald-600 uppercase tracking-tighter">Check In</p>
                  <div className="flex items-baseline gap-1">
                    <span className="text-lg font-mono font-bold text-gray-800">{formatTime(log.check_in)}</span>
                    <span className="text-[9px] font-bold text-[#006E62] uppercase">{log.status_in}</span>
                  </div>
                  {log.in_address && (
                    <p className="text-[8px] text-gray-400 line-clamp-1 flex items-center gap-1">
                      <MapPin size={8} /> {log.in_address}
                    </p>
                  )}
                  {log.late_reason && (
                    <div className="flex items-start gap-1 p-1.5 bg-white rounded border border-emerald-100/50">
                       <MessageSquare size={10} className="text-rose-400 shrink-0 mt-0.5" />
                       <p className="text-[8px] text-gray-500 italic">"{log.late_reason}"</p>
                    </div>
                  )}
                </div>

                {/* Info Pulang */}
                <div className="p-3 bg-blue-50/50 rounded-lg border border-blue-100 space-y-2">
                  <p className="text-[8px] font-bold text-blue-600 uppercase tracking-tighter">Check Out</p>
                  <div className="flex items-baseline gap-1">
                    <span className="text-lg font-mono font-bold text-gray-800">{formatTime(log.check_out)}</span>
                    <span className={`text-[9px] font-bold uppercase ${log.status_out === 'Pulang Cepat' ? 'text-rose-500' : 'text-blue-500'}`}>{log.status_out}</span>
                  </div>
                  {log.out_address && (
                    <p className="text-[8px] text-gray-400 line-clamp-1 flex items-center gap-1">
                      <MapPin size={8} /> {log.out_address}
                    </p>
                  )}
                  {log.early_departure_reason && (
                    <div className="flex items-start gap-1 p-1.5 bg-white rounded border border-blue-100/50">
                       <MessageSquare size={10} className="text-rose-400 shrink-0 mt-0.5" />
                       <p className="text-[8px] text-gray-500 italic">"{log.early_departure_reason}"</p>
                    </div>
                  )}
                </div>

                <div className="p-3 bg-gray-50 rounded-lg border border-gray-100">
                  <p className="text-[8px] font-bold text-gray-400 uppercase tracking-tighter mb-1">Durasi Kerja</p>
                  <div className="flex items-center gap-1.5">
                    <Clock size={14} className="text-gray-400" />
                    <span className="text-sm font-bold text-gray-700">{log.work_duration || '-'}</span>
                  </div>
                </div>
              </div>
            </div>

            <div className="flex items-center gap-3 md:border-l md:pl-6 border-gray-50">
              {log.in_photo_id && (
                <div 
                  onClick={() => Swal.fire({ imageUrl: googleDriveService.getFileUrl(log.in_photo_id!), imageWidth: 400, showConfirmButton: false, background: '#000' })}
                  className="w-14 h-14 rounded-lg bg-gray-100 overflow-hidden cursor-pointer hover:ring-2 hover:ring-[#006E62] transition-all border border-gray-200"
                >
                  <img src={googleDriveService.getFileUrl(log.in_photo_id)} className="w-full h-full object-cover" alt="IN" />
                </div>
              )}
              {log.out_photo_id && (
                <div 
                  onClick={() => Swal.fire({ imageUrl: googleDriveService.getFileUrl(log.out_photo_id!), imageWidth: 400, showConfirmButton: false, background: '#000' })}
                  className="w-14 h-14 rounded-lg bg-gray-100 overflow-hidden cursor-pointer hover:ring-2 hover:ring-[#006E62] transition-all border border-gray-200"
                >
                  <img src={googleDriveService.getFileUrl(log.out_photo_id)} className="w-full h-full object-cover" alt="OUT" />
                </div>
              )}
              <div className="hidden group-hover:flex ml-2">
                <button className="p-2 text-gray-400 hover:text-[#006E62] hover:bg-emerald-50 rounded-full transition-all">
                  <ExternalLink size={18} />
                </button>
              </div>
            </div>
          </div>
        </div>
      ))}
    </div>
  );
};

export default PresenceHistory;
