import React from 'react';
import { Megaphone, Clock, User, ChevronRight, AlertCircle, Info, Calendar, FileText } from 'lucide-react';
import { Announcement } from '../../types';

interface AnnouncementCardProps {
  announcement: Announcement;
  onClick: () => void;
}

const AnnouncementCard: React.FC<AnnouncementCardProps> = ({ announcement, onClick }) => {
  const getCategoryStyles = (category: string) => {
    switch (category) {
      case 'Urgent': return { bg: 'bg-rose-50', text: 'text-rose-600', icon: <AlertCircle size={14} /> };
      case 'Event': return { bg: 'bg-emerald-50', text: 'text-emerald-600', icon: <Calendar size={14} /> };
      case 'Policy': return { bg: 'bg-purple-50', text: 'text-purple-600', icon: <FileText size={14} /> };
      default: return { bg: 'bg-blue-50', text: 'text-blue-600', icon: <Info size={14} /> };
    }
  };

  const styles = getCategoryStyles(announcement.category);

  return (
    <div 
      onClick={onClick}
      className="bg-white border border-gray-100 p-5 rounded-3xl shadow-sm hover:shadow-md transition-all cursor-pointer group relative overflow-hidden"
    >
      {!announcement.is_read && (
        <div className="absolute top-0 right-0">
          <div className="bg-amber-500 text-white text-[8px] font-black uppercase px-3 py-1 rounded-bl-xl animate-pulse tracking-widest">
            NEW
          </div>
        </div>
      )}
      
      <div className="flex items-start gap-4">
        <div className={`w-12 h-12 rounded-2xl ${styles.bg} ${styles.text} flex items-center justify-center shrink-0 group-hover:scale-110 transition-transform`}>
          <Megaphone size={24} />
        </div>
        
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-2 mb-1">
            <span className={`text-[9px] font-bold uppercase tracking-widest px-2 py-0.5 rounded-full ${styles.bg} ${styles.text} flex items-center gap-1`}>
              {styles.icon}
              {announcement.category}
            </span>
            <span className="text-[9px] text-gray-400 font-bold uppercase tracking-widest">•</span>
            <span className="text-[9px] text-gray-400 font-bold uppercase tracking-widest flex items-center gap-1">
              <Clock size={10} />
              {new Date(announcement.created_at).toLocaleDateString('id-ID', { day: 'numeric', month: 'short' })}
            </span>
          </div>
          
          <h3 className="text-sm font-bold text-gray-800 mb-1 truncate group-hover:text-amber-600 transition-colors">
            {announcement.title}
          </h3>
          
          <p className="text-xs text-gray-500 line-clamp-2 leading-relaxed">
            {announcement.content}
          </p>
          
          <div className="mt-4 flex items-center justify-between">
            <div className="flex items-center gap-2">
              <div className="w-5 h-5 rounded-full bg-gray-100 flex items-center justify-center text-gray-400 text-[8px] font-bold">
                {announcement.creator?.full_name.charAt(0)}
              </div>
              <span className="text-[10px] font-bold text-gray-400 uppercase tracking-tight">{announcement.creator?.full_name}</span>
            </div>
            
            <div className="flex items-center gap-1 text-[10px] font-bold text-amber-600 uppercase tracking-widest opacity-0 group-hover:opacity-100 transition-opacity">
              Baca Selengkapnya <ChevronRight size={12} />
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default AnnouncementCard;
