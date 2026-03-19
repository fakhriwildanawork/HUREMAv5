import React, { useState } from 'react';
import { ArrowLeft, Save, Receipt, Calendar, CreditCard, Wallet, Upload, Info, FileText } from 'lucide-react';
import { financeService } from '../../services/financeService';
import { authService } from '../../services/authService';
import { ReimbursementInput } from '../../types';
import { googleDriveService as driveService } from '../../services/googleDriveService';
import Swal from 'sweetalert2';

interface ReimbursementFormProps {
  onBack: () => void;
}

const BANK_LIST = [
  'BCA', 'Mandiri', 'BNI', 'BRI', 'BTN', 'CIMB Niaga', 'Permata', 'Danamon', 'BSI', 'OCBC NISP'
];

const EWALLET_LIST = [
  'GoPay', 'OVO', 'Dana', 'LinkAja', 'ShopeePay'
];

const CATEGORIES = ['Operasional', 'Akomodasi', 'Inventaris', 'Lainnya'];

const ReimbursementForm: React.FC<ReimbursementFormProps> = ({ onBack }) => {
  const [loading, setLoading] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [formData, setFormData] = useState<ReimbursementInput>({
    transaction_date: new Date().toISOString().split('T')[0],
    category: 'Operasional',
    description: '',
    amount_requested: 0,
    proof_file_id: null,
    payment_method: 'Transfer',
    target_type: 'Bank',
    target_name: '',
    account_number: '',
    account_holder: '',
  });

  const user = authService.getCurrentUser();

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setUploading(true);
    try {
      const fileId = await driveService.uploadFile(file);
      setFormData(prev => ({ ...prev, proof_file_id: fileId }));
      Swal.fire({
        icon: 'success',
        title: 'Berhasil',
        text: 'Bukti transaksi berhasil diunggah.',
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

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.proof_file_id) {
      Swal.fire('Peringatan', 'Silakan unggah bukti transaksi terlebih dahulu.', 'warning');
      return;
    }

    setLoading(true);
    try {
      await financeService.createReimbursement({
        ...formData,
        account_id: user?.id || '',
      });
      Swal.fire('Berhasil!', 'Pengajuan reimburse berhasil dikirim.', 'success');
      onBack();
    } catch (error) {
      console.error('Error saving reimbursement:', error);
      Swal.fire('Gagal!', 'Terjadi kesalahan saat mengirim pengajuan.', 'error');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="max-w-4xl mx-auto space-y-6">
      <div className="flex items-center gap-4">
        <button
          onClick={onBack}
          className="p-2 hover:bg-gray-100 rounded-full transition-colors text-gray-500"
        >
          <ArrowLeft size={20} />
        </button>
        <div>
          <h2 className="text-2xl font-bold text-gray-800">Ajukan Reimburse</h2>
          <p className="text-sm text-gray-500">Isi detail transaksi dan unggah bukti untuk pengajuan reimburse.</p>
        </div>
      </div>

      <form onSubmit={handleSubmit} className="space-y-6">
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {/* Detail Transaksi */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 space-y-6">
            <div className="flex items-center gap-2 text-[#006E62] mb-2">
              <Receipt size={18} />
              <h3 className="font-bold text-sm uppercase tracking-wider">Detail Transaksi</h3>
            </div>
            
            <div className="space-y-4">
              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
                  <Calendar size={12} />
                  Tanggal Transaksi
                </label>
                <input
                  type="date"
                  value={formData.transaction_date}
                  onChange={(e) => setFormData({ ...formData, transaction_date: e.target.value })}
                  className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium"
                  required
                />
              </div>

              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Kategori</label>
                <select
                  value={formData.category}
                  onChange={(e) => setFormData({ ...formData, category: e.target.value })}
                  className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium"
                >
                  {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>

              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Nominal Transaksi</label>
                <div className="relative">
                  <div className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400 text-sm font-medium">Rp</div>
                  <input
                    type="number"
                    value={formData.amount_requested}
                    onChange={(e) => setFormData({ ...formData, amount_requested: Number(e.target.value) })}
                    className="w-full pl-10 pr-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium"
                    required
                  />
                </div>
              </div>

              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Keterangan</label>
                <textarea
                  value={formData.description}
                  onChange={(e) => setFormData({ ...formData, description: e.target.value })}
                  placeholder="Jelaskan keperluan transaksi ini..."
                  className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium min-h-[100px]"
                  required
                />
              </div>
            </div>
          </div>

          {/* Bukti & Pembayaran */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 space-y-6">
            <div className="flex items-center gap-2 text-blue-600 mb-2">
              <CreditCard size={18} />
              <h3 className="font-bold text-sm uppercase tracking-wider">Bukti & Tujuan Pembayaran</h3>
            </div>

            <div className="space-y-4">
              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
                  <Upload size={12} />
                  Bukti Transaksi (Foto/File)
                </label>
                <div className="relative">
                  <input
                    type="file"
                    onChange={handleFileUpload}
                    className="hidden"
                    id="proof-upload"
                    accept="image/*,.pdf"
                  />
                  <label
                    htmlFor="proof-upload"
                    className={`flex flex-col items-center justify-center w-full p-4 border-2 border-dashed rounded-xl cursor-pointer transition-all ${
                      formData.proof_file_id ? 'border-emerald-200 bg-emerald-50' : 'border-gray-200 hover:border-[#006E62] hover:bg-gray-50'
                    }`}
                  >
                    {uploading ? (
                      <div className="flex flex-col items-center gap-2">
                        <div className="w-6 h-6 border-2 border-[#006E62] border-t-transparent rounded-full animate-spin"></div>
                        <span className="text-xs font-medium text-gray-500">Mengunggah...</span>
                      </div>
                    ) : formData.proof_file_id ? (
                      <div className="flex flex-col items-center gap-1">
                        <FileText className="text-emerald-600" size={24} />
                        <span className="text-xs font-bold text-emerald-700 uppercase tracking-wider">File Terunggah</span>
                      </div>
                    ) : (
                      <div className="flex flex-col items-center gap-1">
                        <Upload className="text-gray-400" size={24} />
                        <span className="text-xs font-medium text-gray-500">Klik untuk pilih file</span>
                      </div>
                    )}
                  </label>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1.5">
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Metode</label>
                  <select
                    value={formData.payment_method}
                    onChange={(e) => setFormData({ ...formData, payment_method: e.target.value as 'Cash' | 'Transfer' })}
                    className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium"
                  >
                    <option value="Cash">Cash</option>
                    <option value="Transfer">Transfer</option>
                  </select>
                </div>
                {formData.payment_method === 'Transfer' && (
                  <div className="space-y-1.5">
                    <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tipe Tujuan</label>
                    <select
                      value={formData.target_type || 'Bank'}
                      onChange={(e) => setFormData({ ...formData, target_type: e.target.value as 'Bank' | 'E-Wallet' })}
                      className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium"
                    >
                      <option value="Bank">Bank</option>
                      <option value="E-Wallet">E-Wallet</option>
                    </select>
                  </div>
                )}
              </div>

              {formData.payment_method === 'Transfer' && (
                <>
                  <div className="space-y-1.5">
                    <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">
                      Nama {formData.target_type}
                    </label>
                    <div className="relative">
                      <input
                        list="target-names"
                        type="text"
                        value={formData.target_name || ''}
                        onChange={(e) => setFormData({ ...formData, target_name: e.target.value })}
                        placeholder={`Pilih atau ketik nama ${formData.target_type}...`}
                        className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium"
                        required
                      />
                      <datalist id="target-names">
                        {(formData.target_type === 'Bank' ? BANK_LIST : EWALLET_LIST).map(item => (
                          <option key={item} value={item} />
                        ))}
                      </datalist>
                    </div>
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">No. Rekening / E-Wallet</label>
                    <input
                      type="text"
                      value={formData.account_number || ''}
                      onChange={(e) => setFormData({ ...formData, account_number: e.target.value })}
                      className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium"
                      required
                    />
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Atas Nama</label>
                    <input
                      type="text"
                      value={formData.account_holder || ''}
                      onChange={(e) => setFormData({ ...formData, account_holder: e.target.value })}
                      className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium"
                      required
                    />
                  </div>
                </>
              )}
            </div>
          </div>
        </div>

        <div className="flex justify-end gap-3 pt-4">
          <button
            type="button"
            onClick={onBack}
            className="px-6 py-2.5 bg-white border border-gray-200 text-gray-700 rounded-lg hover:bg-gray-50 transition-all text-sm font-medium"
          >
            Batal
          </button>
          <button
            type="submit"
            disabled={loading || uploading}
            className="flex items-center gap-2 px-8 py-2.5 bg-[#006E62] text-white rounded-lg hover:bg-[#005a50] transition-all text-sm font-medium shadow-md disabled:opacity-50"
          >
            {loading ? (
              <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></div>
            ) : (
              <Save size={18} />
            )}
            Kirim Pengajuan
          </button>
        </div>
      </form>
    </div>
  );
};

export default ReimbursementForm;
