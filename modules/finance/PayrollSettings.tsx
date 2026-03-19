import React, { useState, useEffect } from 'react';
import { ArrowLeft, Save, Building, MapPin, Phone, Mail, Globe, Image as ImageIcon, Info } from 'lucide-react';
import { financeService } from '../../services/financeService';
import { PayrollSettings as PayrollSettingsType } from '../../types';
import Swal from 'sweetalert2';

interface PayrollSettingsProps {
  onBack: () => void;
}

const PayrollSettings: React.FC<PayrollSettingsProps> = ({ onBack }) => {
  const [loading, setLoading] = useState(false);
  const [settings, setSettings] = useState<PayrollSettingsType | null>(null);
  const [formData, setFormData] = useState<Partial<PayrollSettingsType>>({
    company_name: '',
    company_address: '',
    company_phone: '',
    company_email: '',
    company_website: '',
    company_logo_url: ''
  });

  useEffect(() => {
    const fetchSettings = async () => {
      try {
        const data = await financeService.getPayrollSettings();
        setSettings(data);
        setFormData(data);
      } catch (error) {
        console.error('Error fetching settings:', error);
      }
    };
    fetchSettings();
  }, []);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    try {
      await financeService.updatePayrollSettings({ ...formData, id: settings?.id });
      Swal.fire('Berhasil!', 'Pengaturan kop payslip berhasil diperbarui.', 'success');
    } catch (error) {
      console.error('Error updating settings:', error);
      Swal.fire('Gagal!', 'Terjadi kesalahan saat menyimpan data.', 'error');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="max-w-4xl mx-auto space-y-6">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-4">
          <button
            onClick={onBack}
            className="p-2 hover:bg-gray-100 rounded-full transition-colors text-gray-500"
          >
            <ArrowLeft size={20} />
          </button>
          <div>
            <h2 className="text-2xl font-bold text-gray-800">Pengaturan Kop Payslip</h2>
            <p className="text-sm text-gray-500">Kustomisasi identitas perusahaan yang akan muncul di slip gaji.</p>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <div className="md:col-span-2">
          <form onSubmit={handleSubmit} className="bg-white p-8 rounded-2xl shadow-sm border border-gray-100 space-y-8">
            <div className="flex items-center gap-2 text-[#006E62] mb-2">
              <Building size={18} />
              <h3 className="font-bold text-sm uppercase tracking-wider">Informasi Perusahaan</h3>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="md:col-span-2 space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Nama Perusahaan</label>
                <input
                  type="text"
                  value={formData.company_name}
                  onChange={(e) => setFormData({ ...formData, company_name: e.target.value })}
                  className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium"
                  required
                />
              </div>

              <div className="md:col-span-2 space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
                  <MapPin size={12} />
                  Alamat Perusahaan
                </label>
                <textarea
                  value={formData.company_address}
                  onChange={(e) => setFormData({ ...formData, company_address: e.target.value })}
                  className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium min-h-[80px]"
                />
              </div>

              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
                  <Phone size={12} />
                  Telepon
                </label>
                <input
                  type="text"
                  value={formData.company_phone}
                  onChange={(e) => setFormData({ ...formData, company_phone: e.target.value })}
                  className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium"
                />
              </div>

              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
                  <Mail size={12} />
                  Email
                </label>
                <input
                  type="email"
                  value={formData.company_email}
                  onChange={(e) => setFormData({ ...formData, company_email: e.target.value })}
                  className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium"
                />
              </div>

              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
                  <Globe size={12} />
                  Website
                </label>
                <input
                  type="text"
                  value={formData.company_website}
                  onChange={(e) => setFormData({ ...formData, company_website: e.target.value })}
                  className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium"
                />
              </div>

              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
                  <ImageIcon size={12} />
                  URL Logo Perusahaan
                </label>
                <input
                  type="text"
                  value={formData.company_logo_url}
                  onChange={(e) => setFormData({ ...formData, company_logo_url: e.target.value })}
                  placeholder="https://example.com/logo.png"
                  className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium"
                />
              </div>
            </div>

            <div className="flex justify-end gap-3 pt-4">
              <button
                type="button"
                onClick={onBack}
                className="px-6 py-2.5 border border-gray-200 text-gray-600 rounded-lg hover:bg-gray-50 transition-all text-sm font-bold uppercase tracking-wider"
              >
                Batal
              </button>
              <button
                type="submit"
                disabled={loading}
                className="flex items-center gap-2 px-8 py-2.5 bg-[#006E62] text-white rounded-lg hover:bg-[#005a50] transition-all text-sm font-bold uppercase tracking-wider shadow-lg disabled:opacity-50"
              >
                {loading ? (
                  <div className="w-5 h-5 border-2 border-white border-t-transparent rounded-full animate-spin"></div>
                ) : (
                  <Save size={18} />
                )}
                Simpan Pengaturan
              </button>
            </div>
          </form>
        </div>

        <div className="space-y-6">
          <div className="bg-white p-6 rounded-2xl shadow-sm border border-gray-100">
            <div className="flex items-center gap-2 text-[#006E62] mb-4">
              <ImageIcon size={18} />
              <h3 className="font-bold text-sm uppercase tracking-wider">Preview Logo</h3>
            </div>
            <div className="aspect-square bg-gray-50 rounded-xl border-2 border-dashed border-gray-200 flex items-center justify-center overflow-hidden">
              {formData.company_logo_url ? (
                <img 
                  src={formData.company_logo_url} 
                  alt="Logo Preview" 
                  className="max-w-full max-h-full object-contain p-4"
                  onError={(e) => (e.currentTarget.src = 'https://via.placeholder.com/150?text=Logo+Error')}
                  referrerPolicy="no-referrer"
                />
              ) : (
                <div className="text-center p-6">
                  <ImageIcon size={48} className="mx-auto text-gray-300 mb-2" strokeWidth={1} />
                  <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Belum Ada Logo</p>
                </div>
              )}
            </div>
          </div>

          <div className="bg-emerald-50 p-6 rounded-2xl border border-emerald-100">
            <div className="flex items-start gap-3">
              <Info className="text-[#006E62] shrink-0 mt-0.5" size={18} />
              <div className="space-y-2">
                <h4 className="text-sm font-bold text-[#006E62]">Tips Kop Payslip</h4>
                <p className="text-xs text-emerald-700 leading-relaxed">
                  Gunakan logo dengan latar belakang transparan (PNG) untuk hasil terbaik. Data ini akan ditampilkan di bagian atas payslip PDF yang diunduh oleh karyawan.
                </p>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default PayrollSettings;
