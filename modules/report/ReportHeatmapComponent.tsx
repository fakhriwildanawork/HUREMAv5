
import React from 'react';
import { format, eachDayOfInterval, parseISO, isWithinInterval } from 'date-fns';
import { id } from 'date-fns/locale';

interface HeatmapProps {
  details: any[];
  startDate: string;
  endDate: string;
}

const AttendanceHeatmap: React.FC<HeatmapProps> = ({ details, startDate, endDate }) => {
  const days = eachDayOfInterval({
    start: parseISO(startDate),
    end: parseISO(endDate)
  });

  return (
    <div className="space-y-4">
      <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Pola Kehadiran (Heatmap)</h4>
      <div className="grid grid-cols-7 gap-2">
        {days.map((day, i) => {
          const dateStr = format(day, 'yyyy-MM-dd');
          const detail = details?.find(d => d.date === dateStr);
          
          let bgColor = 'bg-gray-100 text-gray-400';
          if (detail?.status === 'present') bgColor = 'bg-[#006E62] text-white';
          else if (detail?.status === 'leave' || detail?.status === 'permission') bgColor = 'bg-amber-100 text-amber-600';
          else if (detail?.status === 'absent') bgColor = 'bg-rose-100 text-rose-600';

          return (
            <div 
              key={i} 
              className={`aspect-square rounded-md flex items-center justify-center text-[10px] font-bold ${bgColor}`}
              title={`${format(day, 'dd MMMM yyyy', { locale: id })}: ${detail?.status || 'No Data'}`}
            >
              {format(day, 'd')}
            </div>
          );
        })}
      </div>
      <div className="flex items-center gap-4 text-[10px] font-bold text-gray-400 uppercase tracking-tighter">
        <div className="flex items-center gap-1">
          <div className="w-3 h-3 bg-[#006E62] rounded-sm"></div>
          <span>Hadir</span>
        </div>
        <div className="flex items-center gap-1">
          <div className="w-3 h-3 bg-amber-100 rounded-sm"></div>
          <span>Cuti/Izin</span>
        </div>
        <div className="flex items-center gap-1">
          <div className="w-3 h-3 bg-rose-100 rounded-sm"></div>
          <span>Absen</span>
        </div>
        <div className="flex items-center gap-1">
          <div className="w-3 h-3 bg-gray-100 rounded-sm"></div>
          <span>Belum Ada Data</span>
        </div>
      </div>
    </div>
  );
};

export default AttendanceHeatmap;
