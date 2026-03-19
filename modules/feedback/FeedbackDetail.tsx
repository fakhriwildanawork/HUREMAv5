import React from 'react';
import { X, Calendar, User, Shield, MessageSquare, Paperclip, Link as LinkIcon, ExternalLink, Download, CheckCircle2, Clock } from 'lucide-react';
import { Feedback } from '../../types';

interface FeedbackDetailProps {
  feedback: Feedback;
  onClose: () => void;
  isAdmin?: boolean;
}

const FeedbackDetail: React.FC<FeedbackDetailProps> = ({ feedback, onClose, isAdmin }) => {
  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
      <div className="bg-white w-full max-w-2xl rounded-3xl overflow-hidden shadow-2xl flex flex-col animate-in fade-in zoom-in duration-300 max-h-[90vh]">
        <div className="p-6 border-b border-gray-100 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-emerald-50 flex items-center justify-center text-[#006E62]">
              <MessageSquare size={20} />
            </div>
            <div>
              <h3 className="text-lg font-bold text-gray-800 tracking-tight">Detail Feedback</h3>
              <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">ID: {feedback.id.substring(0, 8)}</p>
            </div>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-gray-100 rounded-xl transition-colors text-gray-400">
            <X size={24} />
          </button>
        </div>

        <div className="p-6 overflow-y-auto space-y-8 scrollbar-none">
          {/* Status & Info Header */}
          <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
            <div className="p-3 bg-gray-50 rounded-2xl border border-gray-100">
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Kategori</p>
              <p className="text-xs font-bold text-gray-700">{feedback.category}</p>
            </div>
            <div className="p-3 bg-gray-50 rounded-2xl border border-gray-100">
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Prioritas</p>
              <p className={`text-xs font-bold ${
                feedback.priority === 'High' ? 'text-rose-500' : 
                feedback.priority === 'Medium' ? 'text-amber-500' : 'text-emerald-500'
              }`}>{feedback.priority}</p>
            </div>
            <div className="p-3 bg-gray-50 rounded-2xl border border-gray-100">
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Status</p>
              <div className="flex items-center gap-1.5">
                {feedback.status === 'Read' ? (
                  <span className="text-xs font-bold text-emerald-600 flex items-center gap-1">
                    <CheckCircle2 size={12} /> Read
                  </span>
                ) : (
                  <span className="text-xs font-bold text-amber-500 flex items-center gap-1">
                    <Clock size={12} /> Unread
                  </span>
                )}
              </div>
            </div>
            <div className="p-3 bg-gray-50 rounded-2xl border border-gray-100">
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Tanggal</p>
              <p className="text-xs font-bold text-gray-700">{new Date(feedback.created_at).toLocaleDateString('id-ID')}</p>
            </div>
          </div>

          {/* Sender Info (Admin Only) */}
          {isAdmin && (
            <div className="p-4 bg-emerald-50/50 rounded-2xl border border-emerald-100">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-white border border-emerald-100 flex items-center justify-center text-[#006E62]">
                  {feedback.is_anonymous ? <Shield size={20} /> : <User size={20} />}
                </div>
                <div>
                  <p className="text-[9px] font-bold text-[#006E62] uppercase tracking-widest">Informasi Pengirim</p>
                  <p className="text-sm font-bold text-gray-800">
                    {feedback.is_anonymous ? 'Pengirim Anonim' : feedback.account?.full_name}
                  </p>
                  {!feedback.is_anonymous && (
                    <p className="text-[10px] text-gray-500 font-medium">NIK: {feedback.account?.internal_nik}</p>
                  )}
                </div>
              </div>
            </div>
          )}

          {/* Description */}
          <div className="space-y-3">
            <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Isi Feedback</h4>
            <div className="p-6 bg-gray-50 rounded-3xl border border-gray-100 relative">
              <div className="absolute top-4 left-4 text-gray-200">
                <MessageSquare size={40} strokeWidth={1} />
              </div>
              <p className="text-sm text-gray-700 leading-relaxed relative z-10 italic">
                "{feedback.description}"
              </p>
            </div>
          </div>

          {/* Attachments */}
          {feedback.attachments && feedback.attachments.length > 0 && (
            <div className="space-y-3">
              <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Lampiran File</h4>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                {feedback.attachments.map((fileId, idx) => (
                  <div key={idx} className="flex items-center justify-between p-3 bg-white border border-gray-100 rounded-xl hover:border-[#006E62] transition-colors group">
                    <div className="flex items-center gap-3 overflow-hidden">
                      <div className="w-8 h-8 rounded-lg bg-gray-50 flex items-center justify-center text-gray-400 group-hover:text-[#006E62] transition-colors">
                        <Paperclip size={16} />
                      </div>
                      <span className="text-[11px] font-bold text-gray-600 truncate">File_{idx + 1}</span>
                    </div>
                    <a 
                      href={`https://drive.google.com/file/d/${fileId}/view`} 
                      target="_blank" 
                      rel="noopener noreferrer"
                      className="p-2 text-gray-400 hover:text-[#006E62] hover:bg-emerald-50 rounded-lg transition-all"
                    >
                      <Download size={14} />
                    </a>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Links */}
          {feedback.links && feedback.links.length > 0 && (
            <div className="space-y-3">
              <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Link Referensi</h4>
              <div className="space-y-2">
                {feedback.links.map((link, idx) => (
                  <a 
                    key={idx}
                    href={link}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="flex items-center justify-between p-3 bg-white border border-gray-100 rounded-xl hover:border-[#006E62] transition-colors group"
                  >
                    <div className="flex items-center gap-3 overflow-hidden">
                      <div className="w-8 h-8 rounded-lg bg-gray-50 flex items-center justify-center text-gray-400 group-hover:text-[#006E62] transition-colors">
                        <LinkIcon size={16} />
                      </div>
                      <span className="text-[11px] font-bold text-gray-600 truncate">{link}</span>
                    </div>
                    <ExternalLink size={14} className="text-gray-300 group-hover:text-[#006E62]" />
                  </a>
                ))}
              </div>
            </div>
          )}
        </div>

        <div className="p-6 border-t border-gray-100 bg-gray-50/50 shrink-0">
          <button 
            onClick={onClose}
            className="w-full py-3 bg-white border border-gray-200 text-gray-600 rounded-xl font-bold text-xs uppercase tracking-widest hover:bg-gray-50 transition-all"
          >
            Tutup
          </button>
        </div>
      </div>
    </div>
  );
};

export default FeedbackDetail;
