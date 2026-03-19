import React, { useState } from 'react';
import { X, CheckCircle2, XCircle, AlertCircle, Calendar, User, FileText, Download, Clock, Save, Loader2, Info, MessageSquare, ClipboardList } from 'lucide-react';
import { DispensationRequest, DispensationIssue, DispensationIssueStatus } from '../../../types';
import { dispensationService } from '../../../services/dispensationService';
import Swal from 'sweetalert2';

interface DispensationDetailProps {
  request: DispensationRequest;
  onClose: () => void;
  onSuccess?: () => void;
  isAdmin: boolean;
}

const DispensationDetail: React.FC<DispensationDetailProps> = ({ request, onClose, onSuccess, isAdmin }) => {
  const [isLoading, setIsLoading] = useState(false);
  const [issues, setIssues] = useState<DispensationIssue[]>([...request.issues]);
  const [globalStatus, setGlobalStatus] = useState(request.status);

  const handleIssueStatusChange = (index: number, status: DispensationIssueStatus) => {
    const newIssues = [...issues];
    newIssues[index].status = status;
    setIssues(newIssues);
  };

  const handleAdminNotesChange = (index: number, notes: string) => {
    const newIssues = [...issues];
    newIssues[index].admin_notes = notes;
    setIssues(newIssues);
  };

  const handleProcess = async () => {
    try {
      setIsLoading(true);
      
      // Determine global status based on issues
      const allApproved = issues.every(i => i.status === 'APPROVED');
      const allRejected = issues.every(i => i.status === 'REJECTED');
      const anyApproved = issues.some(i => i.status === 'APPROVED');
      
      let finalStatus: DispensationRequest['status'] = 'PENDING';
      if (allApproved) finalStatus = 'APPROVED';
      else if (allRejected) finalStatus = 'REJECTED';
      else if (anyApproved) finalStatus = 'PARTIAL';

      await dispensationService.updateStatus(request.id, finalStatus, issues);

      Swal.fire({
        title: 'Berhasil!',
        text: 'Keputusan dispensasi telah disimpan.',
        icon: 'success',
        timer: 1500,
        showConfirmButton: false
      });
      
      if (onSuccess) onSuccess();
    } catch (error) {
      console.error(error);
      Swal.fire('Gagal', 'Terjadi kesalahan saat memproses data.', 'error');
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[70] flex items-center justify-center p-4 animate-in fade-in duration-300">
      <div className="bg-white w-full max-w-3xl rounded-3xl shadow-2xl overflow-hidden flex flex-col max-h-[90vh]">
        {/* Header */}
        <div className="px-8 py-6 bg-gray-50 border-b border-gray-100 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-[#006E62]/10 rounded-xl flex items-center justify-center text-[#006E62]">
              <ClipboardList size={24} />
            </div>
            <div>
              <h3 className="text-lg font-bold text-gray-800 tracking-tight">Detail Pengajuan Dispensasi</h3>
              <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Informasi Lengkap & Status Verifikasi</p>
            </div>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-gray-100 rounded-full transition-colors text-gray-400">
            <X size={24} />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-8 space-y-8">
          {/* Info Pegawai & Tanggal */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="bg-gray-50 p-4 rounded-2xl border border-gray-100 space-y-3">
              <div className="flex items-center gap-2 text-[10px] font-bold text-gray-400 uppercase tracking-widest">
                <User size={12} />
                <span>Data Pegawai</span>
              </div>
              <div>
                <p className="text-sm font-bold text-gray-800">{request.account?.full_name || 'N/A'}</p>
                <p className="text-[11px] text-gray-500 font-medium">{request.account?.internal_nik || 'N/A'}</p>
              </div>
            </div>

            <div className="bg-gray-50 p-4 rounded-2xl border border-gray-100 space-y-3">
              <div className="flex items-center gap-2 text-[10px] font-bold text-gray-400 uppercase tracking-widest">
                <Calendar size={12} />
                <span>Tanggal Bermasalah</span>
              </div>
              <div>
                <p className="text-sm font-bold text-gray-800">{new Date(request.date).toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}</p>
                <p className="text-[11px] text-gray-500 font-medium italic">Diajukan pada: {new Date(request.created_at).toLocaleString('id-ID')}</p>
              </div>
            </div>
          </div>

          {/* Alasan & Bukti */}
          <div className="space-y-4">
            <div className="flex items-center gap-2 text-[10px] font-bold text-gray-400 uppercase tracking-widest ml-1">
              <FileText size={14} />
              <span>Alasan Pengajuan</span>
            </div>
            <div className="p-5 bg-white border border-gray-100 rounded-2xl shadow-sm italic text-sm text-gray-600 leading-relaxed">
              "{request.reason}"
            </div>

            {request.file_id && (
              <div className="pt-2">
                <a 
                  href={`https://drive.google.com/uc?id=${request.file_id}&export=download`}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="inline-flex items-center gap-2 px-4 py-2 bg-blue-50 text-blue-600 rounded-xl text-[10px] font-bold uppercase tracking-wider hover:bg-blue-100 transition-all border border-blue-100"
                >
                  <Download size={14} />
                  Lihat File Bukti
                </a>
              </div>
            )}
          </div>

          {/* Daftar Masalah & Verifikasi */}
          <div className="space-y-4">
            <div className="flex items-center gap-2 text-[10px] font-bold text-gray-400 uppercase tracking-widest ml-1">
              <AlertCircle size={14} />
              <span>Daftar Masalah & Keputusan</span>
            </div>

            <div className="space-y-3">
              {issues.map((issue, idx) => (
                <div key={idx} className={`p-5 rounded-2xl border transition-all ${
                  issue.status === 'APPROVED' ? 'bg-emerald-50/50 border-emerald-100' :
                  issue.status === 'REJECTED' ? 'bg-rose-50/50 border-rose-100' :
                  'bg-white border-gray-100'
                }`}>
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                    <div className="flex items-center gap-3">
                      <div className={`w-8 h-8 rounded-lg flex items-center justify-center ${
                        issue.status === 'APPROVED' ? 'bg-emerald-500 text-white' :
                        issue.status === 'REJECTED' ? 'bg-rose-500 text-white' :
                        'bg-blue-500 text-white'
                      }`}>
                        <AlertCircle size={16} />
                      </div>
                      <span className="text-xs font-bold text-gray-700 uppercase tracking-wider">{issue.type.replace('_', ' ')}</span>
                    </div>

                    {isAdmin && request.status === 'PENDING' ? (
                      <div className="flex bg-gray-100 p-1 rounded-xl w-fit">
                        <button 
                          onClick={() => handleIssueStatusChange(idx, 'APPROVED')}
                          className={`px-4 py-1.5 rounded-lg text-[9px] font-bold uppercase transition-all flex items-center gap-1 ${issue.status === 'APPROVED' ? 'bg-emerald-500 text-white shadow-md' : 'text-gray-400 hover:text-gray-600'}`}
                        >
                          <CheckCircle2 size={12} /> Setujui
                        </button>
                        <button 
                          onClick={() => handleIssueStatusChange(idx, 'REJECTED')}
                          className={`px-4 py-1.5 rounded-lg text-[9px] font-bold uppercase transition-all flex items-center gap-1 ${issue.status === 'REJECTED' ? 'bg-rose-500 text-white shadow-md' : 'text-gray-400 hover:text-gray-600'}`}
                        >
                          <XCircle size={12} /> Tolak
                        </button>
                      </div>
                    ) : (
                      <div className="flex items-center gap-2">
                        {issue.status === 'APPROVED' ? (
                          <span className="text-[10px] font-bold text-emerald-600 uppercase tracking-widest flex items-center gap-1"><CheckCircle2 size={12} /> Disetujui</span>
                        ) : issue.status === 'REJECTED' ? (
                          <span className="text-[10px] font-bold text-rose-600 uppercase tracking-widest flex items-center gap-1"><XCircle size={12} /> Ditolak</span>
                        ) : (
                          <span className="text-[10px] font-bold text-blue-600 uppercase tracking-widest flex items-center gap-1"><Clock size={12} /> Pending</span>
                        )}
                      </div>
                    )}
                  </div>

                  {/* Admin Notes */}
                  {(isAdmin && request.status === 'PENDING') ? (
                    <div className="mt-4 space-y-2">
                      <div className="flex items-center gap-2 text-[9px] font-bold text-gray-400 uppercase tracking-widest ml-1">
                        <MessageSquare size={10} />
                        <span>Catatan Admin (Opsional)</span>
                      </div>
                      <input 
                        type="text"
                        value={issue.admin_notes || ''}
                        onChange={(e) => handleAdminNotesChange(idx, e.target.value)}
                        placeholder="Berikan alasan atau catatan..."
                        className="w-full px-4 py-2 bg-white border border-gray-200 rounded-xl text-xs focus:ring-2 focus:ring-[#006E62] outline-none transition-all"
                      />
                    </div>
                  ) : issue.admin_notes && (
                    <div className="mt-4 p-3 bg-gray-50 rounded-xl border border-gray-100">
                      <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Catatan Admin:</p>
                      <p className="text-xs text-gray-600 italic">"{issue.admin_notes}"</p>
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className="px-8 py-6 bg-gray-50 border-t border-gray-100 flex items-center justify-between">
          <div className="flex items-center gap-2 text-gray-400">
            <Info size={14} />
            <p className="text-[10px] font-medium italic">
              {isAdmin ? 'Pastikan bukti sudah sesuai sebelum memberikan keputusan.' : 'Keputusan admin bersifat final dan akan mempengaruhi payroll.'}
            </p>
          </div>
          <div className="flex gap-3">
            <button
              onClick={onClose}
              className="px-6 py-3 rounded-xl text-xs font-bold text-gray-500 uppercase tracking-wider hover:bg-gray-200 transition-all"
            >
              Tutup
            </button>
            {isAdmin && request.status === 'PENDING' && (
              <button
                onClick={handleProcess}
                disabled={isLoading}
                className="flex items-center gap-2 bg-[#006E62] text-white px-8 py-3 rounded-xl font-bold text-xs uppercase tracking-wider hover:bg-[#005c52] transition-all shadow-lg shadow-[#006E62]/20 disabled:opacity-50"
              >
                {isLoading ? (
                  <Loader2 className="animate-spin" size={18} />
                ) : (
                  <>
                    <Save size={18} />
                    Simpan Keputusan
                  </>
                )}
              </button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

export default DispensationDetail;
