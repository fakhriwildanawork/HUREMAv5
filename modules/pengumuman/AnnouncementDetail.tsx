import React, { useEffect, useState } from 'react';
import { X, Calendar, Clock, User, Megaphone, Paperclip, ExternalLink, Eye, CheckCircle2, Users } from 'lucide-react';
import { Announcement } from '../../types';
import { announcementService } from '../../services/announcementService';

interface AnnouncementDetailProps {
  announcement: Announcement;
  userId: string;
  isAdmin: boolean;
  onClose: () => void;
  onRead?: () => void;
}

const AnnouncementDetail: React.FC<AnnouncementDetailProps> = ({ announcement, userId, isAdmin, onClose, onRead }) => {
  const [readReceipts, setReadReceipts] = useState<any[]>([]);
  const [showReceipts, setShowReceipts] = useState(false);

  useEffect(() => {
    if (!announcement.is_read) {
      announcementService.markAsRead(announcement.id, userId).then(() => {
        if (onRead) onRead();
      });
    }
    
    if (isAdmin) {
      announcementService.getReadReceipts(announcement.id).then(setReadReceipts);
    }
  }, [announcement.id, userId, isAdmin]);

  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
      <div className="bg-white w-full max-w-3xl rounded-3xl overflow-hidden shadow-2xl flex flex-col animate-in fade-in zoom-in duration-300 max-h-[90vh]">
        <div className="p-6 border-b border-gray-100 flex items-center justify-between shrink-0 bg-gray-50/50">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-amber-100 text-amber-600 flex items-center justify-center">
              <Megaphone size={20} />
            </div>
            <div>
              <h3 className="text-lg font-bold text-gray-800 tracking-tight">{announcement.title}</h3>
              <div className="flex items-center gap-3 mt-0.5">
                <span className="text-[10px] text-gray-400 font-bold uppercase tracking-widest flex items-center gap-1">
                  <Calendar size={12} />
                  {new Date(announcement.created_at).toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' })}
                </span>
                <span className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">•</span>
                <span className="text-[10px] text-gray-400 font-bold uppercase tracking-widest flex items-center gap-1">
                  <User size={12} />
                  {announcement.creator?.full_name}
                </span>
              </div>
            </div>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-gray-100 rounded-xl transition-colors text-gray-400">
            <X size={24} />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-8 scrollbar-none">
          <div className="prose prose-sm max-w-none">
            <div className="text-gray-700 leading-relaxed whitespace-pre-wrap text-sm font-medium bg-gray-50/50 p-6 rounded-2xl border border-gray-100">
              {announcement.content}
            </div>
          </div>

          {announcement.attachments.length > 0 && (
            <div className="mt-8 space-y-3">
              <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Lampiran Resmi</h4>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                {announcement.attachments.map((fileId, idx) => (
                  <a 
                    key={idx}
                    href={`https://drive.google.com/file/d/${fileId}/view`}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="flex items-center justify-between p-3 bg-white border border-gray-100 rounded-xl hover:border-amber-200 hover:bg-amber-50/30 transition-all group"
                  >
                    <div className="flex items-center gap-3">
                      <Paperclip size={14} className="text-gray-400 group-hover:text-amber-600" />
                      <span className="text-xs font-bold text-gray-600 group-hover:text-amber-600">Dokumen_{idx + 1}</span>
                    </div>
                    <ExternalLink size={14} className="text-gray-300 group-hover:text-amber-600" />
                  </a>
                ))}
              </div>
            </div>
          )}

          {isAdmin && (
            <div className="mt-12 pt-8 border-t border-gray-100">
              <div className="flex items-center justify-between mb-4">
                <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1 flex items-center gap-2">
                  <Eye size={14} /> Statistik Pembaca
                </h4>
                <button 
                  onClick={() => setShowReceipts(!showReceipts)}
                  className="text-[10px] font-bold text-amber-600 uppercase hover:underline"
                >
                  {showReceipts ? 'Sembunyikan Daftar' : `Lihat ${readReceipts.length} Pembaca`}
                </button>
              </div>
              
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
                <div className="p-4 bg-gray-50 rounded-2xl border border-gray-100">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Total Baca</p>
                  <p className="text-xl font-black text-gray-800">{readReceipts.length}</p>
                </div>
                <div className="p-4 bg-gray-50 rounded-2xl border border-gray-100">
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Target</p>
                  <p className="text-xs font-bold text-gray-700 uppercase">{announcement.target_type}</p>
                </div>
              </div>

              {showReceipts && (
                <div className="mt-6 space-y-2 max-h-[200px] overflow-y-auto scrollbar-thin pr-2">
                  {readReceipts.map((receipt, idx) => (
                    <div key={idx} className="flex items-center justify-between p-3 bg-white border border-gray-50 rounded-xl">
                      <div className="flex items-center gap-3">
                        <div className="w-8 h-8 rounded-full bg-emerald-50 text-emerald-600 flex items-center justify-center text-[10px] font-bold">
                          {receipt.user?.full_name.charAt(0)}
                        </div>
                        <div>
                          <p className="text-[10px] font-bold text-gray-800">{receipt.user?.full_name}</p>
                          <p className="text-[8px] font-bold text-gray-400 uppercase">{receipt.user?.internal_nik}</p>
                        </div>
                      </div>
                      <div className="text-right">
                        <p className="text-[8px] font-bold text-emerald-600 uppercase flex items-center gap-1 justify-end">
                          <CheckCircle2 size={10} /> Dibaca
                        </p>
                        <p className="text-[8px] font-bold text-gray-300 uppercase">
                          {new Date(receipt.read_at).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })}
                        </p>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}
        </div>

        <div className="p-6 border-t border-gray-100 bg-gray-50/50 flex justify-end shrink-0">
          <button 
            onClick={onClose}
            className="px-8 py-2 bg-gray-800 text-white rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-black transition-all"
          >
            Selesai Membaca
          </button>
        </div>
      </div>
    </div>
  );
};

export default AnnouncementDetail;
