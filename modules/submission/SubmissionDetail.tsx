import React, { useState } from 'react';
import { X, CheckCircle, XCircle, Clock, User, FileText, Paperclip, ExternalLink, Calendar, MessageSquare } from 'lucide-react';
import { Submission } from '../../types';
import { googleDriveService } from '../../services/googleDriveService';

interface SubmissionDetailProps {
  submission: Submission;
  onClose: () => void;
  onVerify: (id: string, status: 'Disetujui' | 'Ditolak', notes?: string) => void;
  canVerify: boolean;
}

const SubmissionDetail: React.FC<SubmissionDetailProps> = ({ submission, onClose, onVerify, canVerify }) => {
  const [notes, setNotes] = useState('');

  const formatDate = (dateStr?: string) => {
    if (!dateStr) return '-';
    return new Date(dateStr).toLocaleString('id-ID', {
      day: 'numeric',
      month: 'long',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });
  };

  const DataItem = ({ label, value }: { label: string, value: string }) => (
    <div className="bg-gray-50 p-3 rounded-lg border border-gray-100">
      <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">{label}</p>
      <p className="text-xs font-bold text-gray-700">{value}</p>
    </div>
  );

  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black bg-opacity-40 p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden flex flex-col max-h-[90vh] animate-in zoom-in duration-200">
        <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between">
          <div className="flex items-center gap-3">
             <div className={`p-2 rounded-lg ${
                submission.status === 'Pending' ? 'bg-orange-50 text-orange-600' :
                submission.status === 'Disetujui' ? 'bg-emerald-50 text-emerald-600' : 'bg-red-50 text-red-600'
             }`}>
                <Clock size={20} />
             </div>
             <div>
                <h3 className="text-base font-bold text-gray-800">Detail Pengajuan {submission.type}</h3>
                <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">ID: {submission.id.slice(0,8)}</p>
             </div>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={20} />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-6 space-y-6">
          <div className="grid grid-cols-2 gap-4">
             <DataItem label="Pengaju" value={submission.account?.full_name || '-'} />
             <DataItem label="NIK Internal" value={submission.account?.internal_nik || '-'} />
             <DataItem label="Tanggal Pengajuan" value={formatDate(submission.created_at)} />
             <DataItem label="Status Saat Ini" value={submission.status} />
          </div>

          <div className="space-y-2">
            <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest border-b pb-1">Data Spesifik Pengajuan</h4>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
               {Object.entries(submission.submission_data).map(([key, value]: [string, any]) => (
                 <div key={key} className="flex justify-between items-center p-2 bg-emerald-50/20 rounded border border-emerald-100/30">
                    <span className="text-[10px] font-medium text-gray-500 capitalize">{key.replace('_', ' ')}</span>
                    <span className="text-[10px] font-bold text-[#006E62]">{value}</span>
                 </div>
               ))}
            </div>
          </div>

          <div className="space-y-2">
            <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest border-b pb-1">Keterangan / Alasan</h4>
            <div className="p-4 bg-gray-50 rounded-xl italic text-xs text-gray-600 leading-relaxed border border-gray-100">
               "{submission.description}"
            </div>
          </div>

          {submission.file_id && (
            <div className="space-y-2">
              <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest border-b pb-1">Lampiran Dokumen</h4>
              <a 
                href={googleDriveService.getFileUrl(submission.file_id).replace('=s1600', '=s0')} 
                target="_blank" 
                rel="noreferrer"
                className="flex items-center justify-between p-3 bg-white border border-gray-200 rounded-xl hover:bg-gray-50 transition-all group"
              >
                 <div className="flex items-center gap-3">
                   <Paperclip size={18} className="text-[#006E62]" />
                   <span className="text-xs font-bold text-gray-700">Lihat File Lampiran</span>
                 </div>
                 <ExternalLink size={14} className="text-gray-300 group-hover:text-[#006E62]" />
              </a>
            </div>
          )}

          {submission.status !== 'Pending' && (
             <div className="space-y-2 p-4 bg-blue-50/50 rounded-xl border border-blue-100">
                <h4 className="text-[10px] font-bold text-blue-600 uppercase tracking-widest mb-2 flex items-center gap-2">
                  <CheckCircle size={14} /> Informasi Verifikasi
                </h4>
                <div className="grid grid-cols-2 gap-4 text-[10px]">
                   <div>
                     <p className="text-gray-400 uppercase">Diverifikasi Oleh</p>
                     <p className="font-bold text-gray-700">ID: {submission.verifier_id?.slice(0,8) || '-'}</p>
                   </div>
                   <div>
                     <p className="text-gray-400 uppercase">Waktu Verifikasi</p>
                     <p className="font-bold text-gray-700">{formatDate(submission.verified_at!)}</p>
                   </div>
                </div>
                {submission.verification_notes && (
                  <div className="mt-2 pt-2 border-t border-blue-100 flex gap-2">
                    <MessageSquare size={12} className="text-blue-400 shrink-0 mt-0.5" />
                    <p className="text-[10px] text-blue-600 italic">"{submission.verification_notes}"</p>
                  </div>
                )}
             </div>
          )}
        </div>

        {canVerify && (
          <div className="p-6 border-t border-gray-100 bg-gray-50 space-y-4">
             <div className="space-y-2">
                <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest px-1">Catatan Verifikator (Opsional)</label>
                <textarea 
                  value={notes}
                  onChange={(e) => setNotes(e.target.value)}
                  placeholder="Berikan alasan penyetujuan atau penolakan..."
                  className="w-full px-3 py-2 text-xs border border-gray-200 rounded-lg focus:ring-1 focus:ring-[#006E62] outline-none resize-none bg-white"
                  rows={2}
                />
             </div>
             <div className="grid grid-cols-2 gap-3">
                <button 
                  onClick={() => onVerify(submission.id, 'Ditolak', notes)}
                  className="flex items-center justify-center gap-2 py-2.5 rounded-lg border-2 border-red-500 text-red-600 text-xs font-bold uppercase hover:bg-red-50 transition-all"
                >
                  <XCircle size={16} /> Tolak Pengajuan
                </button>
                <button 
                  onClick={() => onVerify(submission.id, 'Disetujui', notes)}
                  className="flex items-center justify-center gap-2 py-2.5 rounded-lg bg-[#006E62] text-white text-xs font-bold uppercase hover:bg-[#005a50] transition-all shadow-md"
                >
                  <CheckCircle size={16} /> Setujui (ACC)
                </button>
             </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default SubmissionDetail;