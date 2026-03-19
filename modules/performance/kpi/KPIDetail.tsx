import React from 'react';
import { X, Target, Calendar, Info, FileText, Link as LinkIcon, CheckCircle2, User, Clock, AlertTriangle } from 'lucide-react';
import { KPI } from '../../../types';
import { googleDriveService } from '../../../services/googleDriveService';

interface KPIDetailProps {
  onClose: () => void;
  kpi: KPI;
}

const KPIDetail: React.FC<KPIDetailProps> = ({ onClose, kpi }) => {
  const isVerified = kpi.status === 'Verified';
  const isUnverified = kpi.status === 'Unverified';
  const isOverdue = kpi.status === 'Unreported';

  const getStatusColor = () => {
    if (isOverdue) return 'text-rose-600 bg-rose-50 border-rose-100';
    if (isVerified) return 'text-emerald-600 bg-emerald-50 border-emerald-100';
    if (isUnverified) return 'text-blue-600 bg-blue-50 border-blue-100';
    return 'text-amber-600 bg-amber-50 border-amber-100';
  };

  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 p-4 backdrop-blur-sm">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in duration-200 flex flex-col max-h-[90vh]">
        {/* Header */}
        <div className="px-6 py-4 border-b border-gray-50 flex items-center justify-between bg-gray-50/50 shrink-0">
          <div className="flex items-center gap-2">
            <Target size={18} className="text-[#006E62]" />
            <h3 className="text-sm font-bold text-gray-800 uppercase tracking-tight">Detail Key Performance Indicator</h3>
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
                <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Judul KPI</p>
                <h4 className="text-sm font-bold text-gray-800 leading-tight">{kpi.title}</h4>
              </div>
              <div>
                <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Pegawai Terkait</p>
                <div className="flex items-center gap-2">
                  <div className="w-8 h-8 rounded-lg bg-[#006E62]/10 flex items-center justify-center text-[#006E62]">
                    <User size={14} />
                  </div>
                  <div>
                    <p className="text-xs font-bold text-gray-800">{kpi.account?.full_name}</p>
                    <p className="text-[10px] text-gray-400 font-medium">{kpi.account?.internal_nik}</p>
                  </div>
                </div>
              </div>
            </div>
            <div className="space-y-4">
              <div className="flex justify-between items-start">
                <div>
                  <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Status Saat Ini</p>
                  <span className={`px-2 py-1 rounded text-[10px] font-bold uppercase border ${getStatusColor()}`}>
                    {kpi.status}
                  </span>
                </div>
                <div className="text-right">
                  <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Bobot KPI</p>
                  <p className="text-sm font-bold text-[#006E62]">{kpi.weight} Poin</p>
                </div>
              </div>
              <div className="flex items-center gap-4 pt-2">
                <div className="flex-1 p-2 bg-gray-50 rounded-xl border border-gray-100">
                  <p className="text-[8px] font-bold text-gray-400 uppercase tracking-widest mb-0.5">Mulai</p>
                  <div className="flex items-center gap-1.5 text-[10px] font-bold text-gray-600">
                    <Calendar size={12} />
                    <span>{new Date(kpi.start_date + 'T00:00:00').toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' })}</span>
                  </div>
                </div>
                <div className="flex-1 p-2 bg-gray-50 rounded-xl border border-gray-100">
                  <p className="text-[8px] font-bold text-gray-400 uppercase tracking-widest mb-0.5">Deadline</p>
                  <div className="flex items-center gap-1.5 text-[10px] font-bold text-rose-600">
                    <Clock size={12} />
                    <span>{new Date(kpi.deadline + 'T00:00:00').toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' })}</span>
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* Description & Supporting Links */}
          <div className="space-y-4 pt-4 border-t border-gray-50">
            <div>
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-2">Deskripsi & Instruksi Admin</p>
              <div className="p-4 bg-gray-50 rounded-2xl border border-gray-100">
                <p className="text-xs text-gray-600 leading-relaxed whitespace-pre-wrap">{kpi.description}</p>
              </div>
            </div>

            {kpi.supporting_links && kpi.supporting_links.length > 0 && (
              <div>
                <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-2">Link Pendukung dari Admin</p>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  {kpi.supporting_links.map((link, idx) => (
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

          {/* Report Section */}
          {(isVerified || isUnverified) && kpi.report_data && (
            <div className="space-y-4 pt-4 border-t border-gray-50">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest">Laporan Capaian Pegawai</p>
                  <p className="text-[10px] font-bold text-[#006E62] uppercase mt-0.5">{kpi.account?.full_name}</p>
                </div>
                <span className="text-[10px] font-bold text-gray-400">{new Date(kpi.report_data.reported_at).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' })}</span>
              </div>
              
              <div className="p-4 bg-emerald-50/30 border border-emerald-100 rounded-2xl space-y-4">
                <p className="text-xs text-gray-700 leading-relaxed italic">"{kpi.report_data.description}"</p>
                
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div>
                    <p className="text-[8px] font-bold text-emerald-600 uppercase tracking-widest mb-2">Lampiran File</p>
                    <div className="space-y-1.5">
                      {kpi.report_data.file_ids.length > 0 ? kpi.report_data.file_ids.map((id, idx) => (
                        <a 
                          key={id} 
                          href={googleDriveService.getFileUrl(id)} 
                          target="_blank" 
                          rel="noopener noreferrer"
                          className="flex items-center gap-2 px-3 py-1.5 bg-white border border-emerald-100 rounded-lg hover:bg-emerald-100 transition-colors group"
                        >
                          <FileText size={12} className="text-emerald-400 group-hover:text-emerald-600" />
                          <span className="text-[10px] font-bold text-emerald-600 uppercase">Dokumen {idx + 1}</span>
                        </a>
                      )) : <p className="text-[10px] text-gray-400 italic">Tidak ada file</p>}
                    </div>
                  </div>
                  <div>
                    <p className="text-[8px] font-bold text-emerald-600 uppercase tracking-widest mb-2">Lampiran Link</p>
                    <div className="space-y-1.5">
                      {kpi.report_data.links.length > 0 ? kpi.report_data.links.map((link, idx) => (
                        <a 
                          key={idx} 
                          href={link} 
                          target="_blank" 
                          rel="noopener noreferrer"
                          className="flex items-center gap-2 px-3 py-1.5 bg-white border border-emerald-100 rounded-lg hover:bg-emerald-100 transition-colors group"
                        >
                          <LinkIcon size={12} className="text-emerald-400 group-hover:text-emerald-600" />
                          <span className="text-[10px] font-bold text-emerald-600 uppercase truncate">Link Bukti {idx + 1}</span>
                        </a>
                      )) : <p className="text-[10px] text-gray-400 italic">Tidak ada link</p>}
                    </div>
                  </div>
                </div>

                <div className="pt-3 border-t border-emerald-100 flex items-center justify-between">
                  <span className="text-[10px] font-bold text-emerald-600 uppercase">Klaim Capaian Mandiri</span>
                  <span className="text-sm font-bold text-emerald-600">{kpi.report_data.self_assessment}%</span>
                </div>
              </div>
            </div>
          )}

          {/* Verification Section */}
          {isVerified && kpi.verification_data && (
            <div className="space-y-4 pt-4 border-t border-gray-50">
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest">Hasil Verifikasi Admin</p>
              <div className="p-4 bg-[#006E62] rounded-2xl text-white shadow-lg shadow-[#006E62]/20">
                <div className="flex items-center justify-between mb-3">
                  <div className="flex items-center gap-2">
                    <CheckCircle2 size={18} />
                    <span className="text-xs font-bold uppercase tracking-tight">Skor Performa Final</span>
                  </div>
                  <span className="text-2xl font-bold">{kpi.verification_data.score}%</span>
                </div>
                <div className="bg-white/10 p-3 rounded-xl">
                  <p className="text-[10px] font-medium leading-relaxed italic">"{kpi.verification_data.notes}"</p>
                </div>
                <p className="text-[8px] font-bold uppercase tracking-widest mt-3 text-white/60 text-right">
                  Diverifikasi pada {new Date(kpi.verification_data.verified_at).toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' })}
                </p>
              </div>
            </div>
          )}

          {/* Unreported Warning */}
          {isOverdue && (
            <div className="p-4 bg-rose-50 border border-rose-100 rounded-2xl flex gap-4 items-start">
              <AlertTriangle size={24} className="text-rose-500 shrink-0" />
              <div>
                <p className="text-xs font-bold text-rose-600 uppercase tracking-tight mb-1">KPI Melewati Deadline</p>
                <p className="text-[10px] text-rose-500 leading-relaxed font-medium">
                  KPI ini telah melewati batas waktu pelaporan. Segera kirimkan laporan capaian Anda agar Admin dapat melakukan verifikasi dan penilaian performa.
                </p>
              </div>
            </div>
          )}
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

export default KPIDetail;
