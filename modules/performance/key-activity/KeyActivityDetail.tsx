import React from 'react';
import { X, Target, Calendar, Info, FileText, Link as LinkIcon, CheckCircle2, User, Clock, AlertTriangle, CheckSquare } from 'lucide-react';
import { KeyActivity, KeyActivityReport } from '../../../types';

interface KeyActivityDetailProps {
  onClose: () => void;
  activity: KeyActivity;
  reports: KeyActivityReport[];
  isAdmin?: boolean;
  onVerify?: (report: KeyActivityReport) => void;
}

const KeyActivityDetail: React.FC<KeyActivityDetailProps> = ({ onClose, activity, reports, isAdmin, onVerify }) => {
  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 p-4 backdrop-blur-sm">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-3xl overflow-hidden animate-in zoom-in duration-200 flex flex-col max-h-[90vh]">
        {/* Header */}
        <div className="px-6 py-4 border-b border-gray-50 flex items-center justify-between bg-gray-50/50 shrink-0">
          <div className="flex items-center gap-2">
            <CheckSquare size={18} className="text-[#006E62]" />
            <h3 className="text-sm font-bold text-gray-800 uppercase tracking-tight">Detail Key Activity</h3>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={20} />
          </button>
        </div>

        {/* Content */}
        <div className="p-6 overflow-y-auto scrollbar-thin space-y-6">
          {/* Basic Info Section */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="space-y-4">
              <div>
                <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Judul Aktivitas</p>
                <h4 className="text-sm font-bold text-gray-800 leading-tight">{activity.title}</h4>
              </div>
              <div>
                <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Jadwal & Repetisi</p>
                <div className="flex items-center gap-2">
                  <div className="w-8 h-8 rounded-lg bg-[#006E62]/10 flex items-center justify-center text-[#006E62]">
                    <Clock size={14} />
                  </div>
                  <div>
                    <p className="text-xs font-bold text-gray-800">{activity.recurrence_type}</p>
                    <p className="text-[10px] text-gray-400 font-medium">
                      {activity.recurrence_type === 'Weekly' && `Setiap: ${activity.recurrence_rule?.days_of_week?.map(d => ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'][d]).join(', ')}`}
                      {activity.recurrence_type === 'Monthly' && `Setiap Tanggal: ${activity.recurrence_rule?.dates_of_month?.join(', ')}`}
                      {activity.recurrence_type === 'Daily' && 'Dilakukan setiap hari kerja'}
                      {activity.recurrence_type === 'EndOfMonth' && 'Dilakukan setiap akhir bulan'}
                    </p>
                  </div>
                </div>
              </div>
            </div>
            <div className="space-y-4">
              <div className="flex justify-between items-start">
                <div>
                  <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Status Master</p>
                  <span className="px-2 py-1 rounded text-[10px] font-bold uppercase border text-emerald-600 bg-emerald-50 border-emerald-100">
                    {activity.status}
                  </span>
                </div>
                <div className="text-right">
                  <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Bobot Per Laporan</p>
                  <p className="text-sm font-bold text-[#006E62]">{activity.weight} Poin</p>
                </div>
              </div>
              <div className="flex items-center gap-4 pt-2">
                <div className="flex-1 p-2 bg-gray-50 rounded-xl border border-gray-100">
                  <p className="text-[8px] font-bold text-gray-400 uppercase tracking-widest mb-0.5">Mulai Berlaku</p>
                  <div className="flex items-center gap-1.5 text-[10px] font-bold text-gray-600">
                    <Calendar size={12} />
                    <span>{new Date(activity.start_date + 'T00:00:00').toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' })}</span>
                  </div>
                </div>
                <div className="flex-1 p-2 bg-gray-50 rounded-xl border border-gray-100">
                  <p className="text-[8px] font-bold text-gray-400 uppercase tracking-widest mb-0.5">Berakhir Pada</p>
                  <div className="flex items-center gap-1.5 text-[10px] font-bold text-gray-600">
                    <Calendar size={12} />
                    <span>{new Date(activity.end_date + 'T00:00:00').toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' })}</span>
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* Description & Supporting Links */}
          <div className="space-y-4 pt-4 border-t border-gray-50">
            <div>
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-2">Deskripsi & Instruksi</p>
              <div className="p-4 bg-gray-50 rounded-2xl border border-gray-100">
                <p className="text-xs text-gray-600 leading-relaxed whitespace-pre-wrap">{activity.description}</p>
              </div>
            </div>

            {activity.supporting_links && activity.supporting_links.length > 0 && (
              <div>
                <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-2">Link Referensi Admin</p>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  {activity.supporting_links.map((link, idx) => (
                    <a 
                      key={idx} 
                      href={link} 
                      target="_blank" 
                      rel="noopener noreferrer"
                      className="flex items-center gap-2 px-3 py-2 bg-white border border-gray-100 rounded-xl hover:bg-emerald-50 hover:border-emerald-200 transition-all group"
                    >
                      <LinkIcon size={14} className="text-gray-400 group-hover:text-[#006E62]" />
                      <span className="text-[10px] font-bold text-gray-500 truncate group-hover:text-[#006E62]">Referensi {idx + 1}</span>
                    </a>
                  ))}
                </div>
              </div>
            )}
          </div>

          {/* Reports History Section */}
          <div className="space-y-4 pt-4 border-t border-gray-50">
            <div className="flex items-center justify-between">
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest">Riwayat Laporan Terkirim</p>
              <span className="text-[10px] font-bold text-[#006E62] uppercase">{reports.length} Laporan</span>
            </div>

            {reports.length === 0 ? (
              <div className="p-8 text-center bg-gray-50 rounded-2xl border border-dashed border-gray-200">
                <FileText size={32} className="mx-auto text-gray-300 mb-2 opacity-50" />
                <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Belum ada laporan terkirim</p>
              </div>
            ) : (
              <div className="space-y-3">
                {reports.map(report => (
                  <div key={report.id} className="p-4 bg-white border border-gray-100 rounded-2xl shadow-sm space-y-3">
                    <div className="flex justify-between items-start">
                      <div className="flex items-center gap-2">
                        <Calendar size={12} className="text-[#006E62]" />
                        <span className="text-[10px] font-bold text-gray-800 uppercase">Untuk Tanggal: {new Date(report.due_date + 'T00:00:00').toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' })}</span>
                      </div>
                      <span className={`px-2 py-0.5 rounded text-[9px] font-bold uppercase ${report.status === 'Verified' ? 'bg-emerald-50 text-emerald-600' : 'bg-blue-50 text-blue-600'}`}>
                        {report.status}
                      </span>
                    </div>
                    
                    <p className="text-xs text-gray-600 leading-relaxed italic">"{report.description}"</p>
                    
                    {report.links && report.links.length > 0 && (
                      <div className="flex flex-wrap gap-2 pt-1">
                        {report.links.map((link, idx) => (
                          <a 
                            key={idx} 
                            href={link} 
                            target="_blank" 
                            rel="noopener noreferrer"
                            className="flex items-center gap-1.5 px-2 py-1 bg-gray-50 border border-gray-100 rounded-lg text-[9px] font-bold text-[#006E62] hover:bg-emerald-50 transition-colors"
                          >
                            <LinkIcon size={10} />
                            <span>Bukti {idx + 1}</span>
                          </a>
                        ))}
                      </div>
                    )}

                    <div className="pt-3 border-t border-gray-50 flex items-center justify-between text-[9px] text-gray-400 font-medium">
                      <div className="flex flex-col gap-1">
                        {isAdmin && report.account && (
                          <div className="flex items-center gap-1 text-[#006E62] font-bold uppercase">
                            <User size={10} />
                            <span>{report.account.full_name}</span>
                          </div>
                        )}
                        <span>Dilaporkan pada: {new Date(report.reported_at).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' })}</span>
                      </div>
                      
                      {report.status === 'Verified' && report.verification_data ? (
                        <div className="flex items-center gap-1 text-emerald-600 font-bold uppercase">
                          <CheckCircle2 size={10} />
                          <span>Skor: {report.verification_data.score}%</span>
                        </div>
                      ) : isAdmin && onVerify ? (
                        <button 
                          onClick={() => onVerify(report)}
                          className="px-3 py-1.5 bg-[#006E62] text-white rounded-lg text-[9px] font-bold uppercase hover:bg-[#005a50] transition-all flex items-center gap-1.5"
                        >
                          <CheckCircle2 size={10} /> Verifikasi
                        </button>
                      ) : null}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>

        {/* Footer */}
        <div className="px-6 py-4 border-t border-gray-50 bg-gray-50/30 flex justify-end shrink-0">
          <button 
            onClick={onClose}
            className="px-6 py-2.5 bg-white border border-gray-100 text-gray-400 rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-gray-50 transition-all shadow-sm"
          >
            Tutup Detail
          </button>
        </div>
      </div>
    </div>
  );
};

export default KeyActivityDetail;
