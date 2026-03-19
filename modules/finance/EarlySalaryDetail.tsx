import React, { useState } from 'react';
import { ArrowLeft, CheckCircle2, XCircle, Clock, Calendar, Wallet, User, FileText, AlertCircle, Upload, ExternalLink } from 'lucide-react';
import { financeService } from '../../services/financeService';
import { authService } from '../../services/authService';
import { EarlySalaryRequest, EarlySalaryStatus } from '../../types';
import Swal from 'sweetalert2';

interface EarlySalaryDetailProps {
  request: EarlySalaryRequest;
  isAdmin: boolean;
  onBack: () => void;
}

const EarlySalaryDetail: React.FC<EarlySalaryDetailProps> = ({ request, isAdmin, onBack }) => {
  const [loading, setLoading] = useState(false);
  const [rejectionReason, setRejectionReason] = useState('');
  const [paymentProofId, setPaymentProofId] = useState('');
  const [showRejectForm, setShowRejectForm] = useState(false);
  const [showPaidForm, setShowPaidForm] = useState(false);

  const user = authService.getCurrentUser();

  const handleUpdateStatus = async (status: EarlySalaryStatus) => {
    if (!user) return;

    if (status === 'Rejected' && !rejectionReason.trim()) {
      Swal.fire('Error', 'Alasan penolakan harus diisi.', 'error');
      return;
    }

    if (status === 'Paid' && !paymentProofId.trim()) {
      Swal.fire('Error', 'ID Bukti Pembayaran harus diisi.', 'error');
      return;
    }

    const confirmText = status === 'Approved' ? 'Setujui pengajuan ini?' : 
                       status === 'Rejected' ? 'Tolak pengajuan ini?' : 
                       'Tandai sebagai sudah dibayar?';

    const result = await Swal.fire({
      title: 'Konfirmasi',
      text: confirmText,
      icon: 'question',
      showCancelButton: true,
      confirmButtonColor: status === 'Rejected' ? '#e11d48' : '#006E62',
      confirmButtonText: 'Ya, Lanjutkan'
    });

    if (!result.isConfirmed) return;

    setLoading(true);
    try {
      await financeService.updateEarlySalaryStatus(request.id, {
        status,
        rejection_reason: status === 'Rejected' ? rejectionReason : null,
        payment_proof_id: status === 'Paid' ? paymentProofId : request.payment_proof_id,
        verifier_id: user.id
      });

      await Swal.fire({
        title: 'Berhasil!',
        text: `Status pengajuan telah diperbarui menjadi ${status}.`,
        icon: 'success',
        confirmButtonColor: '#006E62'
      });
      onBack();
    } catch (error) {
      console.error('Error updating early salary status:', error);
      Swal.fire('Error', 'Gagal memperbarui status.', 'error');
    } finally {
      setLoading(false);
    }
  };

  const getStatusIcon = (status: EarlySalaryStatus) => {
    switch (status) {
      case 'Approved': return <CheckCircle2 className="text-blue-500" size={24} />;
      case 'Paid': return <CheckCircle2 className="text-emerald-500" size={24} />;
      case 'Rejected': return <XCircle className="text-rose-500" size={24} />;
      default: return <Clock className="text-amber-500" size={24} />;
    }
  };

  const getStatusLabel = (status: EarlySalaryStatus) => {
    switch (status) {
      case 'Approved': return 'Disetujui (Menunggu Pembayaran)';
      case 'Paid': return 'Sudah Dibayar';
      case 'Rejected': return 'Ditolak';
      default: return 'Menunggu Verifikasi';
    }
  };

  return (
    <div className="max-w-3xl mx-auto">
      <div className="flex items-center gap-4 mb-8">
        <button 
          onClick={onBack}
          className="p-2 hover:bg-gray-100 rounded-full transition-colors"
        >
          <ArrowLeft size={20} className="text-gray-600" />
        </button>
        <div>
          <h2 className="text-2xl font-bold text-gray-800">Detail Pengajuan Gaji Awal</h2>
          <p className="text-sm text-gray-500">Informasi lengkap mengenai pengajuan kasbon gaji.</p>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <div className="md:col-span-2 space-y-6">
          <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6 space-y-6">
            <div className="flex items-center justify-between pb-4 border-b border-gray-50">
              <div className="flex items-center gap-3">
                {getStatusIcon(request.status)}
                <div>
                  <div className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Status Saat Ini</div>
                  <div className="text-sm font-bold text-gray-800">{getStatusLabel(request.status)}</div>
                </div>
              </div>
              <div className="text-right">
                <div className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">ID Pengajuan</div>
                <div className="text-xs font-mono text-gray-500">{request.id.slice(0, 8)}...</div>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-6">
              <div className="space-y-1">
                <div className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-1">
                  <Calendar size={12} />
                  Periode Gaji
                </div>
                <div className="text-sm font-bold text-gray-800">
                  {new Date(request.year, request.month - 1).toLocaleString('id-ID', { month: 'long', year: 'numeric' })}
                </div>
              </div>
              <div className="space-y-1">
                <div className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-1">
                  <Wallet size={12} />
                  Nominal Kasbon
                </div>
                <div className="text-lg font-mono font-bold text-[#006E62]">
                  Rp {request.amount.toLocaleString('id-ID')}
                </div>
              </div>
            </div>

            <div className="space-y-1">
              <div className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-1">
                <AlertCircle size={12} />
                Alasan Pengajuan
              </div>
              <div className="text-sm text-gray-600 bg-gray-50 p-4 rounded-lg italic">
                "{request.reason}"
              </div>
            </div>

            {request.status === 'Rejected' && (
              <div className="space-y-1 p-4 bg-rose-50 rounded-lg border border-rose-100">
                <div className="text-[10px] font-bold text-rose-400 uppercase tracking-widest flex items-center gap-1">
                  <XCircle size={12} />
                  Alasan Penolakan
                </div>
                <div className="text-sm text-rose-700 font-medium">
                  {request.rejection_reason}
                </div>
              </div>
            )}

            {request.status === 'Paid' && request.payment_proof_id && (
              <div className="space-y-1 p-4 bg-emerald-50 rounded-lg border border-emerald-100">
                <div className="text-[10px] font-bold text-emerald-400 uppercase tracking-widest flex items-center gap-1">
                  <CheckCircle2 size={12} />
                  Bukti Pembayaran
                </div>
                <div className="flex items-center justify-between">
                  <div className="text-sm text-emerald-700 font-medium">
                    ID: {request.payment_proof_id}
                  </div>
                  <a 
                    href={`https://drive.google.com/file/d/${request.payment_proof_id}/view`}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="flex items-center gap-1 text-xs font-bold text-emerald-600 hover:underline"
                  >
                    LIHAT FILE <ExternalLink size={12} />
                  </a>
                </div>
              </div>
            )}
          </div>

          {isAdmin && request.status === 'Pending' && !showRejectForm && (
            <div className="flex gap-3">
              <button
                onClick={() => setShowRejectForm(true)}
                className="flex-1 flex items-center justify-center gap-2 py-3 bg-white border-2 border-rose-100 text-rose-600 rounded-xl hover:bg-rose-50 transition-all font-bold text-sm"
              >
                <XCircle size={18} />
                TOLAK PENGAJUAN
              </button>
              <button
                onClick={() => handleUpdateStatus('Approved')}
                disabled={loading}
                className="flex-1 flex items-center justify-center gap-2 py-3 bg-[#006E62] text-white rounded-xl hover:bg-[#005a50] transition-all font-bold text-sm shadow-lg shadow-emerald-100"
              >
                <CheckCircle2 size={18} />
                SETUJUI PENGAJUAN
              </button>
            </div>
          )}

          {isAdmin && request.status === 'Approved' && !showPaidForm && (
            <button
              onClick={() => setShowPaidForm(true)}
              className="w-full flex items-center justify-center gap-2 py-3 bg-emerald-600 text-white rounded-xl hover:bg-emerald-700 transition-all font-bold text-sm shadow-lg shadow-emerald-100"
            >
              <Upload size={18} />
              INPUT BUKTI BAYAR & SELESAIKAN
            </button>
          )}

          {showRejectForm && (
            <div className="bg-white rounded-xl shadow-sm border border-rose-100 p-6 space-y-4">
              <div className="flex items-center gap-2 text-rose-600 font-bold text-sm">
                <XCircle size={18} />
                Konfirmasi Penolakan
              </div>
              <textarea
                placeholder="Berikan alasan penolakan..."
                rows={3}
                value={rejectionReason}
                onChange={(e) => setRejectionReason(e.target.value)}
                className="w-full bg-gray-50 border-none rounded-lg p-4 text-sm focus:ring-2 focus:ring-rose-500 transition-all resize-none"
              />
              <div className="flex justify-end gap-3">
                <button onClick={() => setShowRejectForm(false)} className="text-xs font-bold text-gray-400 hover:text-gray-600">BATAL</button>
                <button 
                  onClick={() => handleUpdateStatus('Rejected')}
                  disabled={loading}
                  className="px-6 py-2 bg-rose-600 text-white rounded-lg text-xs font-bold hover:bg-rose-700 transition-all"
                >
                  KONFIRMASI TOLAK
                </button>
              </div>
            </div>
          )}

          {showPaidForm && (
            <div className="bg-white rounded-xl shadow-sm border border-emerald-100 p-6 space-y-4">
              <div className="flex items-center gap-2 text-emerald-600 font-bold text-sm">
                <Upload size={18} />
                Input Bukti Pembayaran
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Google Drive File ID</label>
                <input
                  type="text"
                  placeholder="Masukkan ID file bukti transfer..."
                  value={paymentProofId}
                  onChange={(e) => setPaymentProofId(e.target.value)}
                  className="w-full bg-gray-50 border-none rounded-lg px-4 py-3 text-sm focus:ring-2 focus:ring-emerald-500 transition-all"
                />
              </div>
              <div className="flex justify-end gap-3">
                <button onClick={() => setShowPaidForm(false)} className="text-xs font-bold text-gray-400 hover:text-gray-600">BATAL</button>
                <button 
                  onClick={() => handleUpdateStatus('Paid')}
                  disabled={loading}
                  className="px-6 py-2 bg-emerald-600 text-white rounded-lg text-xs font-bold hover:bg-emerald-700 transition-all"
                >
                  SIMPAN & SELESAIKAN
                </button>
              </div>
            </div>
          )}
        </div>

        <div className="space-y-6">
          <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6 space-y-4">
            <h3 className="text-xs font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
              <User size={14} />
              Informasi Pegawai
            </h3>
            <div className="space-y-3">
              <div>
                <div className="text-[10px] text-gray-400 font-bold uppercase">Nama Lengkap</div>
                <div className="text-sm font-bold text-gray-800">{request.account?.full_name}</div>
              </div>
              <div>
                <div className="text-[10px] text-gray-400 font-bold uppercase">NIK Internal</div>
                <div className="text-sm font-mono text-gray-600">{request.account?.internal_nik}</div>
              </div>
            </div>
          </div>

          <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6 space-y-4">
            <h3 className="text-xs font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
              <FileText size={14} />
              Riwayat Pengajuan
            </h3>
            <div className="space-y-4">
              <div className="relative pl-6 border-l-2 border-gray-100 space-y-4">
                <div className="relative">
                  <div className="absolute -left-[31px] top-1 w-4 h-4 rounded-full bg-gray-200 border-4 border-white"></div>
                  <div className="text-[10px] font-bold text-gray-400 uppercase">Diajukan Pada</div>
                  <div className="text-xs text-gray-600">{new Date(request.created_at).toLocaleString('id-ID')}</div>
                </div>
                {request.verified_at && (
                  <div className="relative">
                    <div className={`absolute -left-[31px] top-1 w-4 h-4 rounded-full border-4 border-white ${request.status === 'Rejected' ? 'bg-rose-400' : 'bg-[#006E62]'}`}></div>
                    <div className="text-[10px] font-bold text-gray-400 uppercase">Diverifikasi Pada</div>
                    <div className="text-xs text-gray-600">{new Date(request.verified_at).toLocaleString('id-ID')}</div>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default EarlySalaryDetail;
