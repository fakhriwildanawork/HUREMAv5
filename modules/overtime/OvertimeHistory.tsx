import React from 'react';
import { Timer, Clock, MapPin, ExternalLink, Calendar, MessageSquare } from 'lucide-react';
import { Overtime } from '../../types';
import { googleDriveService } from '../../services/googleDriveService';
import Swal from 'sweetalert2';

interface OvertimeHistoryProps {
  logs: Overtime[];
  isLoading: boolean;
}

const OvertimeHistory: React.FC<OvertimeHistoryProps> = ({ logs, isLoading }) => {
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

  if (isLoading) return <div className="text-center py-20">Memuat riwayat lembur...</div>;

  if (logs.length === 0) return (
    <div className="flex flex-col items-center justify-center py-32 text-gray-300">
      <Calendar size={48} strokeWidth={1} />
      <p className="text-sm font-bold uppercase mt-4">Belum Ada Riwayat Lembur</p>
    </div>
  );

  return (
    <div className="space-y-4 pb-20">
      {logs.map((log) => (
        <div key={log.id} className="bg-white rounded-xl border border-gray-100 p-5 shadow-sm hover:shadow-md transition-all group border-l-4 border-l-amber-500">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
            <div className="space-y-4 flex-1">
              <div className="flex items-center gap-2">
                <Calendar size={14} className="text-gray-400" />
                <h4 className="text-sm font-bold text-gray-800">{formatDate(log.created_at!)}</h4>
              </div>
              
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                <div className="p-3 bg-amber-50/50 rounded-lg border border-amber-100 space-y-1">
                  <p className="text-[8px] font-bold text-amber-600 uppercase tracking-tighter">Check In OT</p>
                  <span className="text-lg font-mono font-bold text-gray-800">{formatTime(log.check_in)}</span>
                  {log.in_address && <p className="text-[8px] text-gray-400 truncate flex items-center gap-1"><MapPin size={8} /> {log.in_address}</p>}
                </div>

                <div className="p-3 bg-amber-50/50 rounded-lg border border-amber-100 space-y-1">
                  <p className="text-[8px] font-bold text-amber-600 uppercase tracking-tighter">Check Out OT</p>
                  <span className="text-lg font-mono font-bold text-gray-800">{formatTime(log.check_out)}</span>
                  {log.out_address && <p className="text-[8px] text-gray-400 truncate flex items-center gap-1"><MapPin size={8} /> {log.out_address}</p>}
                </div>

                <div className="p-3 bg-gray-50 rounded-lg border border-gray-100 flex flex-col justify-between min-h-[60px]">
                  <div>
                    <p className="text-[8px] font-bold text-gray-400 uppercase tracking-tighter mb-1">Durasi Lembur</p>
                    <div className="flex items-center gap-1.5">
                      <Clock size={14} className="text-amber-600" />
                      <span className="text-sm font-bold text-gray-700">{log.work_duration || '-'}</span>
                    </div>
                  </div>
                  {log.reason && (
                    <div className="mt-2 pt-2 border-t border-gray-200/50 flex items-start gap-1.5">
                       <MessageSquare size={10} className="text-amber-500 shrink-0 mt-0.5" />
                       <p className="text-[9px] text-gray-500 italic leading-tight line-clamp-2">"{log.reason}"</p>
                    </div>
                  )}
                </div>
              </div>
            </div>

            <div className="flex items-center gap-3 md:border-l md:pl-6 border-gray-50">
              {log.in_photo_id && (
                <div 
                  onClick={() => Swal.fire({ imageUrl: googleDriveService.getFileUrl(log.in_photo_id!), background: '#000', showConfirmButton: false })}
                  className="w-14 h-14 rounded-lg bg-gray-100 overflow-hidden cursor-pointer hover:ring-2 hover:ring-amber-500 transition-all border border-gray-200"
                >
                  <img src={googleDriveService.getFileUrl(log.in_photo_id)} className="w-full h-full object-cover" alt="IN" />
                </div>
              )}
              {log.out_photo_id && (
                <div 
                  onClick={() => Swal.fire({ imageUrl: googleDriveService.getFileUrl(log.out_photo_id!), background: '#000', showConfirmButton: false })}
                  className="w-14 h-14 rounded-lg bg-gray-100 overflow-hidden cursor-pointer hover:ring-2 hover:ring-amber-500 transition-all border border-gray-200"
                >
                  <img src={googleDriveService.getFileUrl(log.out_photo_id)} className="w-full h-full object-cover" alt="OUT" />
                </div>
              )}
            </div>
          </div>
        </div>
      ))}
    </div>
  );
};

export default OvertimeHistory;