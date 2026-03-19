import React from 'react';
import { X, Calendar, User, Shield, AlertTriangle, Paperclip, Link as LinkIcon, CheckCircle2, Clock, Users, ExternalLink } from 'lucide-react';
import { Whistleblowing } from '../../types';

interface LaporDetailProps {
  report: Whistleblowing;
  onClose: () => void;
  isAdmin: boolean;
}

const LaporDetail: React.FC<LaporDetailProps> = ({ report, onClose, isAdmin }) => {
  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
      <div className="bg-white w-full max-w-2xl rounded-3xl overflow-hidden shadow-2xl flex flex-col animate-in fade-in zoom-in duration-300 max-h-[90vh]">
        <div className="p-6 border-b border-gray-100 flex items-center justify-between shrink-0 bg-rose-50/30">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-rose-100 flex items-center justify-center text-rose-600">
              <AlertTriangle size={20} />
            </div>
            <div>
              <h3 className="text-lg font-bold text-gray-800 tracking-tight">Detail Laporan</h3>
              <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Informasi lengkap laporan pelanggaran</p>
            </div>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-gray-100 rounded-xl transition-colors text-gray-400">
            <X size={24} />
          </button>
        </div>

        <div className="p-6 overflow-y-auto space-y-8 scrollbar-none">
          <div className="grid grid-cols-2 gap-4">
            <div className="p-4 bg-gray-50 rounded-2xl border border-gray-100">
              <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Kategori</p>
              <p className="text-sm font-bold text-rose-600">{report.category}</p>
            </div>
            <div className="p-4 bg-gray-50 rounded-2xl border border-gray-100">
              <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Status</p>
              <div className="flex items-center gap-1.5">
                {report.status === 'Read' ? (
                  <span className="text-xs font-bold text-emerald-600 uppercase flex items-center gap-1">
                    <CheckCircle2 size={14} />
                    Ditinjau
                  </span>
                ) : (
                  <span className="text-xs font-bold text-amber-500 uppercase flex items-center gap-1">
                    <Clock size={14} />
                    Pending
                  </span>
                )}
              </div>
            </div>
          </div>

          {isAdmin && (
            <div className="space-y-3">
              <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Informasi Pelapor</h4>
              <div className="flex items-center gap-4 p-4 bg-emerald-50/30 border border-emerald-100 rounded-2xl">
                <div className="w-12 h-12 rounded-full bg-emerald-100 flex items-center justify-center text-emerald-600 font-bold text-lg">
                  {report.account?.full_name?.charAt(0)}
                </div>
                <div>
                  <p className="text-sm font-bold text-gray-800">{report.account?.full_name}</p>
                  <p className="text-[10px] font-bold text-emerald-600 uppercase tracking-widest">{report.account?.internal_nik}</p>
                </div>
              </div>
            </div>
          )}

          <div className="space-y-3">
            <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Pihak Terlapor</h4>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              {report.reported_accounts?.map(acc => (
                <div key={acc.id} className="flex items-center gap-3 p-3 bg-gray-50 border border-gray-100 rounded-xl">
                  <div className="w-8 h-8 rounded-full bg-rose-100 flex items-center justify-center text-rose-600 text-xs font-bold">
                    {acc.full_name.charAt(0)}
                  </div>
                  <div className="overflow-hidden">
                    <p className="text-xs font-bold text-gray-800 truncate">{acc.full_name}</p>
                    <p className="text-[9px] font-bold text-gray-400 uppercase">{acc.internal_nik}</p>
                  </div>
                </div>
              ))}
              {(!report.reported_accounts || report.reported_accounts.length === 0) && (
                <p className="text-xs text-gray-400 italic px-1">Tidak ada data terlapor.</p>
              )}
            </div>
          </div>

          <div className="space-y-3">
            <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Deskripsi Laporan</h4>
            <div className="p-5 bg-gray-50 rounded-2xl border border-gray-100">
              <p className="text-sm text-gray-700 leading-relaxed whitespace-pre-wrap font-medium">{report.description}</p>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="space-y-3">
              <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Lampiran Bukti</h4>
              <div className="space-y-2">
                {report.attachments && report.attachments.length > 0 ? (
                  report.attachments.map((fileId, idx) => (
                    <a 
                      key={idx}
                      href={`https://drive.google.com/file/d/${fileId}/view`}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="flex items-center justify-between p-3 bg-white border border-gray-100 rounded-xl hover:border-rose-200 hover:bg-rose-50/30 transition-all group"
                    >
                      <div className="flex items-center gap-3">
                        <Paperclip size={14} className="text-gray-400 group-hover:text-rose-600" />
                        <span className="text-xs font-bold text-gray-600 group-hover:text-rose-600">Bukti_{idx + 1}</span>
                      </div>
                      <ExternalLink size={14} className="text-gray-300 group-hover:text-rose-600" />
                    </a>
                  ))
                ) : (
                  <p className="text-xs text-gray-400 italic px-1">Tidak ada lampiran file.</p>
                )}
              </div>
            </div>

            <div className="space-y-3">
              <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Link Referensi</h4>
              <div className="space-y-2">
                {report.links && report.links.length > 0 ? (
                  report.links.map((link, idx) => (
                    <a 
                      key={idx}
                      href={link}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="flex items-center justify-between p-3 bg-white border border-gray-100 rounded-xl hover:border-rose-200 hover:bg-rose-50/30 transition-all group"
                    >
                      <div className="flex items-center gap-3">
                        <LinkIcon size={14} className="text-gray-400 group-hover:text-rose-600" />
                        <span className="text-xs font-bold text-gray-600 group-hover:text-rose-600 truncate max-w-[150px]">{link}</span>
                      </div>
                      <ExternalLink size={14} className="text-gray-300 group-hover:text-rose-600" />
                    </a>
                  ))
                ) : (
                  <p className="text-xs text-gray-400 italic px-1">Tidak ada link referensi.</p>
                )}
              </div>
            </div>
          </div>
        </div>

        <div className="p-6 border-t border-gray-100 bg-gray-50/50 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-2 text-[10px] font-bold text-gray-400 uppercase tracking-widest">
            <Calendar size={14} />
            <span>Dilaporkan pada: {new Date(report.created_at).toLocaleString('id-ID', { day: 'numeric', month: 'long', year: 'numeric', hour: '2-digit', minute: '2-digit' })}</span>
          </div>
          <button 
            onClick={onClose}
            className="px-6 py-2 bg-gray-800 text-white rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-black transition-all"
          >
            Tutup
          </button>
        </div>
      </div>
    </div>
  );
};

export default LaporDetail;
