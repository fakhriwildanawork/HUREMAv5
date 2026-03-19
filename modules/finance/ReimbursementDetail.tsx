import React, { useState } from 'react';
import { ArrowLeft, CheckCircle2, XCircle, AlertTriangle, FileText, ExternalLink, Receipt, CreditCard, Clock, User, MessageSquare, Upload, Save, Calendar } from 'lucide-react';
import { Reimbursement, ReimbursementStatus } from '../../types';
import { financeService } from '../../services/financeService';
import { authService } from '../../services/authService';
import { googleDriveService as driveService } from '../../services/googleDriveService';
import Swal from 'sweetalert2';

interface ReimbursementDetailProps {
  reimbursement: Reimbursement;
  isAdmin: boolean;
  onBack: () => void;
}

const ReimbursementDetail: React.FC<ReimbursementDetailProps> = ({ reimbursement, isAdmin, onBack }) => {
  const [loading, setLoading] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [verificationMode, setVerificationMode] = useState<ReimbursementStatus | null>(null);
  const [adminNotes, setAdminNotes] = useState('');
  const [amountApproved, setAmountApproved] = useState(reimbursement.amount_requested);
  const [paymentProofId, setPaymentProofId] = useState<string | null>(null);

  const user = authService.getCurrentUser();

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setUploading(true);
    try {
      const fileId = await driveService.uploadFile(file);
      setPaymentProofId(fileId);
      Swal.fire({
        icon: 'success',
        title: 'Berhasil',
        text: 'Bukti transfer berhasil diunggah.',
        toast: true,
        position: 'top-end',
        showConfirmButton: false,
        timer: 3000
      });
    } catch (error) {
      console.error('Upload error:', error);
      Swal.fire('Gagal', 'Gagal mengunggah file.', 'error');
    } finally {
      setUploading(false);
    }
  };

  const handleVerify = async () => {
    if (!verificationMode) return;
    
    if ((verificationMode === 'Partially Approved' || verificationMode === 'Rejected') && !adminNotes.trim()) {
      Swal.fire('Peringatan', 'Keterangan wajib diisi untuk status ini.', 'warning');
      return;
    }

    if ((verificationMode === 'Approved' || verificationMode === 'Partially Approved') && !paymentProofId) {
      Swal.fire('Peringatan', 'Bukti pembayaran wajib diunggah.', 'warning');
      return;
    }

    if (verificationMode === 'Partially Approved' && amountApproved > reimbursement.amount_requested) {
      Swal.fire('Peringatan', 'Nominal setuju sebagian tidak boleh melebihi nominal pengajuan.', 'warning');
      return;
    }

    setLoading(true);
    try {
      await financeService.updateReimbursementStatus(reimbursement.id, {
        status: verificationMode,
        amount_approved: verificationMode === 'Approved' ? reimbursement.amount_requested : (verificationMode === 'Partially Approved' ? amountApproved : 0),
        admin_notes: adminNotes,
        payment_proof_id: paymentProofId || undefined,
        verifier_id: user?.id || ''
      });
      Swal.fire('Berhasil!', 'Verifikasi berhasil disimpan.', 'success');
      onBack();
    } catch (error) {
      console.error('Error verifying:', error);
      Swal.fire('Gagal!', 'Terjadi kesalahan saat menyimpan verifikasi.', 'error');
    } finally {
      setLoading(false);
    }
  };

  const getStatusBadge = (status: string) => {
    switch (status) {
      case 'Approved': return <span className="px-3 py-1 bg-emerald-50 text-emerald-600 border border-emerald-100 rounded-full text-xs font-bold uppercase">Disetujui</span>;
      case 'Partially Approved': return <span className="px-3 py-1 bg-blue-50 text-blue-600 border border-blue-100 rounded-full text-xs font-bold uppercase">Setuju Sebagian</span>;
      case 'Rejected': return <span className="px-3 py-1 bg-rose-50 text-rose-600 border border-rose-100 rounded-full text-xs font-bold uppercase">Ditolak</span>;
      default: return <span className="px-3 py-1 bg-amber-50 text-amber-600 border border-amber-100 rounded-full text-xs font-bold uppercase">Pending</span>;
    }
  };

  return (
    <div className="max-w-4xl mx-auto space-y-6">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-4">
          <button onClick={onBack} className="p-2 hover:bg-gray-100 rounded-full transition-colors text-gray-500">
            <ArrowLeft size={20} />
          </button>
          <div>
            <h2 className="text-2xl font-bold text-gray-800">Detail Reimburse</h2>
            <p className="text-sm text-gray-500">Informasi lengkap pengajuan dan status verifikasi.</p>
          </div>
        </div>
        {getStatusBadge(reimbursement.status)}
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        {/* Left Column: Details */}
        <div className="md:col-span-2 space-y-6">
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 space-y-6">
            <div className="flex items-center gap-2 text-[#006E62] border-b border-gray-50 pb-4">
              <Receipt size={18} />
              <h3 className="font-bold text-sm uppercase tracking-wider">Informasi Pengajuan</h3>
            </div>

            <div className="grid grid-cols-2 gap-6">
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Pegawai</label>
                <div className="flex items-center gap-2">
                  <User size={14} className="text-gray-400" />
                  <span className="text-sm font-semibold text-gray-800">{reimbursement.account?.full_name}</span>
                </div>
                <div className="text-xs text-gray-500 ml-5">{reimbursement.account?.internal_nik}</div>
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right block">Tanggal Transaksi</label>
                <div className="flex items-center gap-2 justify-end">
                  <Calendar size={14} className="text-gray-400" />
                  <span className="text-sm font-semibold text-gray-800">
                    {new Date(reimbursement.transaction_date).toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' })}
                  </span>
                </div>
              </div>
            </div>

            <div className="space-y-1">
              <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Kategori & Keterangan</label>
              <div className="text-xs font-bold text-[#006E62] uppercase mb-1">{reimbursement.category}</div>
              <p className="text-sm text-gray-600 bg-gray-50 p-3 rounded-lg leading-relaxed">
                {reimbursement.description}
              </p>
            </div>

            <div className="grid grid-cols-2 gap-6">
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Nominal Pengajuan</label>
                <div className="text-lg font-mono font-bold text-gray-900">
                  Rp {reimbursement.amount_requested.toLocaleString('id-ID')}
                </div>
              </div>
              {reimbursement.amount_approved !== null && (
                <div className="space-y-1 text-right">
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Nominal Disetujui</label>
                  <div className="text-lg font-mono font-bold text-emerald-600">
                    Rp {reimbursement.amount_approved.toLocaleString('id-ID')}
                  </div>
                </div>
              )}
            </div>
          </div>

          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 space-y-6">
            <div className="flex items-center gap-2 text-blue-600 border-b border-gray-50 pb-4">
              <CreditCard size={18} />
              <h3 className="font-bold text-sm uppercase tracking-wider">Tujuan Pembayaran</h3>
            </div>
            <div className="grid grid-cols-2 gap-6">
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Metode</label>
                <div className="text-sm font-semibold text-gray-800">{reimbursement.payment_method}</div>
              </div>
              {reimbursement.payment_method === 'Transfer' && (
                <>
                  <div className="space-y-1">
                    <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tipe & Nama</label>
                    <div className="text-sm font-semibold text-gray-800">{reimbursement.target_type} - {reimbursement.target_name}</div>
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">No. Rekening / ID</label>
                    <div className="text-sm font-mono font-bold text-gray-800">{reimbursement.account_number}</div>
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Atas Nama</label>
                    <div className="text-sm font-semibold text-gray-800">{reimbursement.account_holder}</div>
                  </div>
                </>
              )}
            </div>
          </div>
        </div>

        {/* Right Column: Evidence & Verification */}
        <div className="space-y-6">
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 space-y-4">
            <div className="flex items-center gap-2 text-amber-600 border-b border-gray-50 pb-4">
              <FileText size={18} />
              <h3 className="font-bold text-sm uppercase tracking-wider">Bukti Transaksi</h3>
            </div>
            {reimbursement.proof_file_id ? (
              <a 
                href={driveService.getFileUrl(reimbursement.proof_file_id)} 
                target="_blank" 
                rel="noreferrer"
                className="group relative block aspect-square rounded-lg overflow-hidden border border-gray-100 hover:border-[#006E62] transition-all"
              >
                <img 
                  src={driveService.getFileUrl(reimbursement.proof_file_id)} 
                  alt="Bukti Transaksi"
                  className="w-full h-full object-cover transition-transform group-hover:scale-105"
                  referrerPolicy="no-referrer"
                />
                <div className="absolute inset-0 bg-black/40 opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center">
                  <ExternalLink className="text-white" size={24} />
                </div>
              </a>
            ) : (
              <div className="aspect-square bg-gray-50 rounded-lg flex items-center justify-center text-gray-400 italic text-xs">
                Tidak ada bukti
              </div>
            )}
          </div>

          {/* Admin Verification Section */}
          {isAdmin && reimbursement.status === 'Pending' ? (
            <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 space-y-6">
              <div className="flex items-center gap-2 text-[#006E62] border-b border-gray-50 pb-4">
                <CheckCircle2 size={18} />
                <h3 className="font-bold text-sm uppercase tracking-wider">Verifikasi Admin</h3>
              </div>

              {!verificationMode ? (
                <div className="grid grid-cols-1 gap-2">
                  <button 
                    onClick={() => {
                      setVerificationMode('Approved');
                      setAmountApproved(reimbursement.amount_requested);
                    }}
                    className="w-full py-2.5 bg-emerald-50 text-emerald-600 rounded-lg text-xs font-bold uppercase hover:bg-emerald-100 transition-all border border-emerald-100"
                  >
                    Setujui Penuh
                  </button>
                  <button 
                    onClick={() => setVerificationMode('Partially Approved')}
                    className="w-full py-2.5 bg-blue-50 text-blue-600 rounded-lg text-xs font-bold uppercase hover:bg-blue-100 transition-all border border-blue-100"
                  >
                    Setuju Sebagian
                  </button>
                  <button 
                    onClick={() => setVerificationMode('Rejected')}
                    className="w-full py-2.5 bg-rose-50 text-rose-600 rounded-lg text-xs font-bold uppercase hover:bg-rose-100 transition-all border border-rose-100"
                  >
                    Tolak Pengajuan
                  </button>
                </div>
              ) : (
                <div className="space-y-4">
                  <div className="flex items-center justify-between">
                    <span className={`text-[10px] font-black uppercase tracking-widest ${
                      verificationMode === 'Approved' ? 'text-emerald-600' : 
                      verificationMode === 'Partially Approved' ? 'text-blue-600' : 'text-rose-600'
                    }`}>
                      Mode: {verificationMode === 'Partially Approved' ? 'Setuju Sebagian' : verificationMode === 'Approved' ? 'Setuju' : 'Tolak'}
                    </span>
                    <button onClick={() => setVerificationMode(null)} className="text-[10px] font-bold text-gray-400 hover:text-gray-600 uppercase">Ganti</button>
                  </div>

                  {verificationMode === 'Partially Approved' && (
                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Nominal Disetujui</label>
                      <div className="relative">
                        <div className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400 text-sm font-medium">Rp</div>
                        <input
                          type="number"
                          value={amountApproved}
                          onChange={(e) => setAmountApproved(Number(e.target.value))}
                          className="w-full pl-10 pr-4 py-2 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium"
                        />
                      </div>
                    </div>
                  )}

                  {(verificationMode === 'Partially Approved' || verificationMode === 'Rejected') && (
                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Keterangan Admin</label>
                      <textarea
                        value={adminNotes}
                        onChange={(e) => setAdminNotes(e.target.value)}
                        placeholder="Berikan alasan penolakan atau penyesuaian..."
                        className="w-full px-4 py-2 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium min-h-[80px]"
                        required
                      />
                    </div>
                  )}

                  {(verificationMode === 'Approved' || verificationMode === 'Partially Approved') && (
                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Bukti Transfer</label>
                      <input type="file" onChange={handleFileUpload} className="hidden" id="payment-upload" accept="image/*,.pdf" />
                      <label
                        htmlFor="payment-upload"
                        className={`flex flex-col items-center justify-center w-full p-3 border-2 border-dashed rounded-lg cursor-pointer transition-all ${
                          paymentProofId ? 'border-emerald-200 bg-emerald-50' : 'border-gray-200 hover:border-[#006E62]'
                        }`}
                      >
                        {uploading ? (
                          <div className="w-4 h-4 border-2 border-[#006E62] border-t-transparent rounded-full animate-spin"></div>
                        ) : paymentProofId ? (
                          <span className="text-[10px] font-bold text-emerald-700 uppercase">Bukti Terunggah</span>
                        ) : (
                          <span className="text-[10px] font-medium text-gray-500 uppercase">Klik Unggah Bukti Bayar</span>
                        )}
                      </label>
                    </div>
                  )}

                  <button
                    onClick={handleVerify}
                    disabled={loading || uploading}
                    className="w-full py-2.5 bg-[#006E62] text-white rounded-lg text-xs font-bold uppercase hover:bg-[#005a50] transition-all shadow-md disabled:opacity-50 flex items-center justify-center gap-2"
                  >
                    {loading ? <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></div> : <Save size={14} />}
                    Simpan Verifikasi
                  </button>
                </div>
              )}
            </div>
          ) : (
            <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 space-y-4">
              <div className="flex items-center gap-2 text-indigo-600 border-b border-gray-50 pb-4">
                <MessageSquare size={18} />
                <h3 className="font-bold text-sm uppercase tracking-wider">Hasil Verifikasi</h3>
              </div>
              {reimbursement.status === 'Pending' ? (
                <div className="flex flex-col items-center gap-2 py-4 text-gray-400 italic">
                  <Clock size={24} />
                  <span className="text-xs">Menunggu verifikasi admin</span>
                </div>
              ) : (
                <div className="space-y-4">
                  {reimbursement.admin_notes && (
                    <div className="space-y-1">
                      <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Catatan Admin</label>
                      <p className="text-sm text-gray-600 bg-gray-50 p-3 rounded-lg italic">"{reimbursement.admin_notes}"</p>
                    </div>
                  )}
                  {reimbursement.payment_proof_id && (
                    <div className="space-y-2">
                      <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Bukti Pembayaran</label>
                      <a 
                        href={driveService.getFileUrl(reimbursement.payment_proof_id)} 
                        target="_blank" 
                        rel="noreferrer"
                        className="flex items-center gap-2 p-2 bg-emerald-50 text-emerald-700 rounded-lg text-xs font-bold hover:bg-emerald-100 transition-all border border-emerald-100"
                      >
                        <ExternalLink size={14} />
                        LIHAT BUKTI TRANSFER
                      </a>
                    </div>
                  )}
                  <div className="pt-2 text-[10px] text-gray-400 font-medium italic">
                    Diverifikasi pada {new Date(reimbursement.verified_at!).toLocaleString('id-ID')}
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default ReimbursementDetail;
