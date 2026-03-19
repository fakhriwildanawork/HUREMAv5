import React, { useState } from 'react';
import { ArrowLeft, Save, Receipt, Clock, AlertTriangle, Info, DollarSign, Briefcase, MapPin, Plus } from 'lucide-react';
import { financeService } from '../../services/financeService';
import { SalaryScheme, SalarySchemeInput } from '../../types';
import Swal from 'sweetalert2';

interface SalarySchemeFormProps {
  scheme?: SalaryScheme | null;
  onBack: () => void;
}

const formatNumber = (num: number) => {
  return num.toLocaleString('id-ID');
};

const parseNumber = (str: string) => {
  if (typeof str !== 'string') return 0;
  return Number(str.replace(/\./g, '')) || 0;
};

const InputGroup = ({ 
  label, 
  name, 
  icon: Icon, 
  prefix = "Rp", 
  value, 
  onChange 
}: { 
  label: string, 
  name: string, 
  icon: any, 
  prefix?: string,
  value: number,
  onChange: (val: number) => void
}) => (
  <div className="space-y-1.5">
    <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
      <Icon size={12} />
      {label}
    </label>
    <div className="relative">
      {prefix && (
        <div className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400 text-sm font-medium">
          {prefix}
        </div>
      )}
      <input
        type="text"
        value={formatNumber(value)}
        onChange={(e) => {
          const val = parseNumber(e.target.value);
          onChange(val);
        }}
        className={`w-full ${prefix ? 'pl-10' : 'px-4'} pr-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium`}
        required
      />
    </div>
  </div>
);

const SalarySchemeForm: React.FC<SalarySchemeFormProps> = ({ scheme, onBack }) => {
  const [loading, setLoading] = useState(false);
  const [formData, setFormData] = useState<SalarySchemeInput>({
    name: scheme?.name || '',
    description: scheme?.description || '',
    type: scheme?.type || 'Bulanan',
    basic_salary: scheme?.basic_salary || 0,
    position_allowance: scheme?.position_allowance || 0,
    placement_allowance: scheme?.placement_allowance || 0,
    other_allowance: scheme?.other_allowance || 0,
    overtime_rate_per_hour: scheme?.overtime_rate_per_hour || 0,
    late_deduction_per_minute: scheme?.late_deduction_per_minute || 0,
    early_leave_deduction_per_minute: scheme?.early_leave_deduction_per_minute || 0,
    no_clock_out_deduction_per_day: scheme?.no_clock_out_deduction_per_day || 0,
    absent_deduction_per_day: scheme?.absent_deduction_per_day || 0,
  });

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    try {
      const dataToSave = { ...formData };
      if (formData.type === 'Harian') {
        dataToSave.position_allowance = 0;
        dataToSave.placement_allowance = 0;
        dataToSave.other_allowance = 0;
        dataToSave.absent_deduction_per_day = 0;
      }

      if (scheme) {
        await financeService.updateScheme(scheme.id, dataToSave);
        Swal.fire('Berhasil!', 'Skema gaji berhasil diperbarui.', 'success');
      } else {
        await financeService.createScheme(dataToSave);
        Swal.fire('Berhasil!', 'Skema gaji baru berhasil dibuat.', 'success');
      }
      onBack();
    } catch (error) {
      console.error('Error saving scheme:', error);
      Swal.fire('Gagal!', 'Terjadi kesalahan saat menyimpan data.', 'error');
    } finally {
      setLoading(false);
    }
  };

  const handleInputChange = (name: keyof SalarySchemeInput, val: number) => {
    setFormData(prev => ({ ...prev, [name]: val }));
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
            <h2 className="text-2xl font-bold text-gray-800">{scheme ? 'Edit Skema Gaji' : 'Buat Skema Gaji Baru'}</h2>
            <p className="text-sm text-gray-500">Tentukan komponen pendapatan dan potongan untuk skema ini.</p>
          </div>
        </div>
      </div>

      <form onSubmit={handleSubmit} className="space-y-6">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          {/* Main Info */}
          <div className="md:col-span-3 bg-white p-6 rounded-xl shadow-sm border border-gray-100 space-y-6">
            <div className="flex items-center gap-2 text-[#006E62] mb-2">
              <Info size={18} />
              <h3 className="font-bold text-sm uppercase tracking-wider">Informasi Dasar</h3>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Nama Skema</label>
                <input
                  type="text"
                  value={formData.name}
                  onChange={(e) => setFormData({ ...formData, name: e.target.value })}
                  placeholder="Contoh: Skema Staff Bulanan"
                  className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium"
                  required
                />
              </div>
              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tipe Gaji</label>
                <select
                  value={formData.type}
                  onChange={(e) => setFormData({ ...formData, type: e.target.value as 'Harian' | 'Bulanan' })}
                  className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium"
                >
                  <option value="Bulanan">Bulanan</option>
                  <option value="Harian">Harian</option>
                </select>
              </div>
              <div className="md:col-span-2 space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Keterangan</label>
                <textarea
                  value={formData.description || ''}
                  onChange={(e) => setFormData({ ...formData, description: e.target.value })}
                  placeholder="Berikan penjelasan singkat mengenai skema ini..."
                  className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium min-h-[80px]"
                />
              </div>
            </div>
          </div>

          {/* Income Components */}
          <div className="md:col-span-2 bg-white p-6 rounded-xl shadow-sm border border-gray-100 space-y-6">
            <div className="flex items-center gap-2 text-emerald-600 mb-2">
              <Receipt size={18} />
              <h3 className="font-bold text-sm uppercase tracking-wider">Komponen Pendapatan</h3>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <InputGroup 
                label={`Gaji Pokok (${formData.type})`} 
                name="basic_salary" 
                icon={DollarSign} 
                value={formData.basic_salary}
                onChange={(val) => handleInputChange('basic_salary', val)}
              />
              {formData.type === 'Bulanan' && (
                <>
                  <InputGroup 
                    label="Tunjangan Jabatan" 
                    name="position_allowance" 
                    icon={Briefcase} 
                    value={formData.position_allowance}
                    onChange={(val) => handleInputChange('position_allowance', val)}
                  />
                  <InputGroup 
                    label="Tunjangan Penempatan" 
                    name="placement_allowance" 
                    icon={MapPin} 
                    value={formData.placement_allowance}
                    onChange={(val) => handleInputChange('placement_allowance', val)}
                  />
                  <InputGroup 
                    label="Tunjangan Lainnya" 
                    name="other_allowance" 
                    icon={Plus} 
                    value={formData.other_allowance}
                    onChange={(val) => handleInputChange('other_allowance', val)}
                  />
                  <InputGroup 
                    label="Upah Lembur / Jam" 
                    name="overtime_rate_per_hour" 
                    icon={Clock} 
                    value={formData.overtime_rate_per_hour}
                    onChange={(val) => handleInputChange('overtime_rate_per_hour', val)}
                  />
                </>
              )}
            </div>
            <div className="p-4 bg-emerald-50 rounded-lg flex gap-3">
              <Info className="text-emerald-600 shrink-0" size={18} />
              <p className="text-xs text-emerald-800 leading-relaxed">
                Tunjangan akan dihitung setiap kali periode penggajian {formData.type.toLowerCase()} dilakukan.
              </p>
            </div>
          </div>

          {/* Deduction Components */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 space-y-6">
            <div className="flex items-center gap-2 text-rose-600 mb-2">
              <AlertTriangle size={18} />
              <h3 className="font-bold text-sm uppercase tracking-wider">Komponen Potongan</h3>
            </div>
            <div className="space-y-6">
              <InputGroup 
                label="Keterlambatan / Menit" 
                name="late_deduction_per_minute" 
                icon={Clock} 
                value={formData.late_deduction_per_minute}
                onChange={(val) => handleInputChange('late_deduction_per_minute', val)}
              />
              <InputGroup 
                label="Pulang Awal / Menit" 
                name="early_leave_deduction_per_minute" 
                icon={Clock} 
                value={formData.early_leave_deduction_per_minute}
                onChange={(val) => handleInputChange('early_leave_deduction_per_minute', val)}
              />
              <InputGroup 
                label="Tanpa Absen Pulang / Hari" 
                name="no_clock_out_deduction_per_day" 
                icon={Clock} 
                value={formData.no_clock_out_deduction_per_day}
                onChange={(val) => handleInputChange('no_clock_out_deduction_per_day', val)}
              />
              {formData.type === 'Bulanan' && (
                <InputGroup 
                  label="Absen (Mangkir) / Hari" 
                  name="absent_deduction_per_day" 
                  icon={AlertTriangle} 
                  value={formData.absent_deduction_per_day}
                  onChange={(val) => handleInputChange('absent_deduction_per_day', val)}
                />
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
            disabled={loading}
            className="flex items-center gap-2 px-8 py-2.5 bg-[#006E62] text-white rounded-lg hover:bg-[#005a50] transition-all text-sm font-medium shadow-md disabled:opacity-50"
          >
            {loading ? (
              <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></div>
            ) : (
              <Save size={18} />
            )}
            Simpan Skema
          </button>
        </div>
      </form>
    </div>
  );
};

export default SalarySchemeForm;
