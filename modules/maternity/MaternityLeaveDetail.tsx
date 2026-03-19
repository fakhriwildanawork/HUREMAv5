import React, { useState } from 'react';
import { X, Calendar, MessageSquare, CheckCircle2, XCircle, Clock, ArrowRight, FileText, Send, History, User, ShieldCheck, Heart } from 'lucide-react';
import Swal from 'sweetalert2';
import { maternityLeaveService } from '../../services/maternityLeaveService';
import { googleDriveService } from '../../services/googleDriveService';
import { MaternityLeaveRequest, AuthUser } from '../../types';

interface MaternityLeaveDetailProps {
  request: MaternityLeaveRequest;
  user: AuthUser;
  onClose: () => void;
  onUpdate: () => void;
}

const MaternityLeaveDetail: React.FC<MaternityLeaveDetailProps> = ({ request, user, onClose, onUpdate }) => {
  const [isNegotiating, setIsNegotiating] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [negoData, setNegoData] = useState({
    start_date: request.start_date,
    end_date: request.end_date,
    reason: ''
  });

  const isAdmin = user?.role === 'admin' || user?.is_hr_admin || user?.is_performance_admin || user?.is_finance_admin;
  const isMyTurn = request.current_negotiator_role === (isAdmin ? 'admin' : 'user');
  const isClosed = ['approved', 'rejected', 'cancelled'].includes(request.status);

  const handleAction = async (status: 'negotiating' | 'approved' | 'rejected' | 'cancelled', reason?: string) => {
    try {
      setIsSaving(true);
      await maternityLeaveService.negotiate(
        request.id,
        isAdmin ? 'admin' : 'user',
        negoData.start_date,
        negoData.end_date,
        reason || negoData.reason || (status === 'approved' ? 'Disetujui' : status === 'rejected' ? 'Ditolak' : 'Dibatalkan'),
        status
      );
      Swal.fire({
        title: 'Berhasil!',
        text: `Status pengajuan telah diperbarui menjadi ${status}.`,
        icon: 'success',
        timer: 1500,
        showConfirmButton: false
      });
      onUpdate();
      onClose();
    } catch (error) {
      Swal.fire('Gagal', 'Terjadi kesalahan saat memproses aksi.', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 p-4 backdrop-blur-sm">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden flex flex-col max-h-[90vh] animate-in slide-in-from-bottom duration-300">
        {/* Header */}
        <div className="px-6 py-4 border-b border-gray-50 flex items-center justify-between bg-gray-50/50">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-[#006E62]/10 rounded-xl flex items-center justify-center text-[#006E62]">
              <Heart size={20} />
            </div>
            <div>
              <h3 className="text-sm font-bold text-gray-800 uppercase tracking-tight">Detail Cuti Melahirkan</h3>
              <p className="text-[9px] text-gray-400 font-bold uppercase tracking-widest">ID: {request.id.split('-')[0]}</p>
            </div>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={20} />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-6 space-y-8 scrollbar-thin">
          {/* Status & Info Utama */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="space-y-4">
              <div className="p-4 bg-gray-50 rounded-2xl border border-gray-100">
                <div className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-2">Jadwal Cuti Saat Ini</div>
                <div className="flex items-center gap-3 text-gray-800 font-bold">
                  <span className="text-lg">{new Date(request.start_date).toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' })}</span>
                  <ArrowRight size={16} className="text-gray-300" />
                  <span className="text-lg">{new Date(request.end_date).toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' })}</span>
                </div>
              </div>
              
              <div className="flex items-center gap-4">
                <div className="flex-1 p-3 bg-emerald-50/50 rounded-xl border border-emerald-100 flex items-center gap-3">
                  <ShieldCheck size={18} className="text-[#006E62]" />
                  <div>
                    <div className="text-[8px] font-bold text-gray-400 uppercase">Status</div>
                    <div className="text-[10px] font-bold text-[#006E62] uppercase tracking-wider">{request.status}</div>
                  </div>
                </div>
                {request.file_id && (
                  <a 
                    href={googleDriveService.getFileUrl(request.file_id)} 
                    target="_blank" 
                    rel="noreferrer"
                    className="flex-1 p-3 bg-blue-50/50 rounded-xl border border-blue-100 flex items-center gap-3 hover:bg-blue-50 transition-colors"
                  >
                    <FileText size={18} className="text-blue-500" />
                    <div>
                      <div className="text-[8px] font-bold text-gray-400 uppercase">Dokumen</div>
                      <div className="text-[10px] font-bold text-blue-500 uppercase tracking-wider">Lihat Lampiran</div>
                    </div>
                  </a>
                )}
              </div>
            </div>

            <div className="space-y-4">
              <div className="p-4 bg-white border border-gray-100 rounded-2xl shadow-sm">
                <div className="flex items-center gap-2 mb-3">
                  <User size={14} className="text-[#006E62]" />
                  <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Informasi Karyawan</span>
                </div>
                <div className="font-bold text-gray-800 text-sm">{request.account?.full_name}</div>
                <div className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">{request.account?.internal_nik}</div>
              </div>
              
              <div className="p-4 bg-amber-50/50 border border-amber-100 rounded-2xl">
                <div className="flex items-center gap-2 mb-2 text-amber-600">
                  <Clock size={14} />
                  <span className="text-[10px] font-bold uppercase tracking-widest">Giliran Respons</span>
                </div>
                <div className="text-[11px] font-bold text-amber-700 uppercase tracking-wider">
                  {isClosed ? 'NEGOSIASI SELESAI' : (isMyTurn ? 'GILIRAN ANDA' : `MENUNGGU ${request.current_negotiator_role.toUpperCase()}`)}
                </div>
              </div>
            </div>
          </div>

          {/* Riwayat Negosiasi */}
          <div className="space-y-4">
            <div className="flex items-center gap-2 px-1">
              <History size={16} className="text-gray-400" />
              <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-[0.2em]">Riwayat Negosiasi</h4>
            </div>
            
            <div className="space-y-4 relative before:absolute before:left-4 before:top-2 before:bottom-2 before:w-0.5 before:bg-gray-50">
              {request.negotiation_data.map((nego, idx) => (
                <div key={idx} className="relative pl-10">
                  <div className={`absolute left-2 top-1 w-4 h-4 rounded-full border-2 border-white shadow-sm ${nego.role === 'admin' ? 'bg-[#006E62]' : 'bg-blue-500'}`}></div>
                  <div className="bg-white border border-gray-100 p-4 rounded-2xl shadow-sm space-y-2">
                    <div className="flex justify-between items-center">
                      <span className={`text-[9px] font-bold uppercase tracking-widest ${nego.role === 'admin' ? 'text-[#006E62]' : 'text-blue-500'}`}>
                        {nego.role === 'admin' ? 'ADMINISTRATOR' : 'KARYAWAN'}
                      </span>
                      <span className="text-[8px] text-gray-300 font-bold">{new Date(nego.timestamp).toLocaleString('id-ID')}</span>
                    </div>
                    <div className="flex items-center gap-2 text-[11px] font-bold text-gray-700">
                      <span>{new Date(nego.start_date).toLocaleDateString('id-ID')}</span>
                      <ArrowRight size={10} className="text-gray-300" />
                      <span>{new Date(nego.end_date).toLocaleDateString('id-ID')}</span>
                    </div>
                    <p className="text-[11px] text-gray-500 italic leading-relaxed">"{nego.reason}"</p>
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* Panel Aksi Negosiasi */}
          {isMyTurn && !isClosed && (
            <div className="pt-6 border-t border-gray-100 animate-in fade-in slide-in-from-top duration-500">
              {!isNegotiating ? (
                <div className="flex flex-wrap gap-3">
                  {isAdmin ? (
                    <>
                      <button 
                        onClick={() => handleAction('approved')}
                        className="flex-1 flex items-center justify-center gap-2 bg-[#006E62] text-white py-3 rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-[#005a50] shadow-lg shadow-[#006E62]/20 transition-all"
                      >
                        <CheckCircle2 size={16} /> Langsung Setujui
                      </button>
                      <button 
                        onClick={() => setIsNegotiating(true)}
                        className="flex-1 flex items-center justify-center gap-2 bg-amber-500 text-white py-3 rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-amber-600 shadow-lg shadow-amber-500/20 transition-all"
                      >
                        <MessageSquare size={16} /> Nego Tanggal
                      </button>
                      <button 
                        onClick={() => handleAction('rejected')}
                        className="flex-1 flex items-center justify-center gap-2 bg-rose-500 text-white py-3 rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-rose-600 shadow-lg shadow-rose-500/20 transition-all"
                      >
                        <XCircle size={16} /> Tolak Cuti
                      </button>
                    </>
                  ) : (
                    <>
                      <button 
                        onClick={() => handleAction('approved')}
                        className="flex-1 flex items-center justify-center gap-2 bg-[#006E62] text-white py-3 rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-[#005a50] shadow-lg shadow-[#006E62]/20 transition-all"
                      >
                        <CheckCircle2 size={16} /> Terima Hasil Nego
                      </button>
                      <button 
                        onClick={() => setIsNegotiating(true)}
                        className="flex-1 flex items-center justify-center gap-2 bg-amber-500 text-white py-3 rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-amber-600 shadow-lg shadow-amber-500/20 transition-all"
                      >
                        <MessageSquare size={16} /> Nego Ulang
                      </button>
                      <button 
                        onClick={() => handleAction('cancelled')}
                        className="flex-1 flex items-center justify-center gap-2 bg-gray-400 text-white py-3 rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-gray-500 shadow-lg shadow-gray-400/20 transition-all"
                      >
                        <XCircle size={16} /> Batalkan Pengajuan
                      </button>
                    </>
                  )}
                </div>
              ) : (
                <div className="bg-gray-50 p-6 rounded-2xl border border-gray-100 space-y-4">
                  <div className="flex items-center justify-between mb-2">
                    <h5 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Form Penawaran Baru</h5>
                    <button onClick={() => setIsNegotiating(false)} className="text-gray-400 hover:text-gray-600"><X size={16} /></button>
                  </div>
                  <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1.5">
                      <label className="text-[9px] font-bold text-gray-400 uppercase px-1">Tgl Awal Baru</label>
                      <input 
                        type="date" 
                        value={negoData.start_date} 
                        onChange={(e) => setNegoData(prev => ({ ...prev, start_date: e.target.value }))}
                        className="w-full px-3 py-2 bg-white border border-gray-100 rounded-xl text-xs font-bold focus:ring-2 focus:ring-[#006E62]/20 outline-none"
                      />
                    </div>
                    <div className="space-y-1.5">
                      <label className="text-[9px] font-bold text-gray-400 uppercase px-1">Tgl Akhir Baru</label>
                      <input 
                        type="date" 
                        value={negoData.end_date} 
                        onChange={(e) => setNegoData(prev => ({ ...prev, end_date: e.target.value }))}
                        className="w-full px-3 py-2 bg-white border border-gray-100 rounded-xl text-xs font-bold focus:ring-2 focus:ring-[#006E62]/20 outline-none"
                      />
                    </div>
                  </div>
                  <div className="space-y-1.5">
                    <label className="text-[9px] font-bold text-gray-400 uppercase px-1">Alasan Penawaran</label>
                    <textarea 
                      placeholder="Jelaskan alasan Anda menawarkan tanggal ini..."
                      value={negoData.reason}
                      onChange={(e) => setNegoData(prev => ({ ...prev, reason: e.target.value }))}
                      rows={2}
                      className="w-full px-3 py-2 bg-white border border-gray-100 rounded-xl text-xs font-medium focus:ring-2 focus:ring-[#006E62]/20 outline-none resize-none"
                    />
                  </div>
                  <button 
                    onClick={() => handleAction('negotiating')}
                    disabled={!negoData.reason}
                    className="w-full py-3 bg-[#006E62] text-white rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-[#005a50] shadow-lg shadow-[#006E62]/20 transition-all flex items-center justify-center gap-2 disabled:opacity-50"
                  >
                    <Send size={16} /> Kirim Penawaran Negosiasi
                  </button>
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default MaternityLeaveDetail;
