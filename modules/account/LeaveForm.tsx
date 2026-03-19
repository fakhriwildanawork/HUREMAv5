import React, { useState, useEffect } from 'react';
import { X, Save, User } from 'lucide-react';
import { LeaveRequestInput, Account } from '../../types';
import { accountService } from '../../services/accountService';

interface LeaveFormProps {
  accountId: string;
  isAdmin?: boolean;
  onClose: () => void;
  onSubmit: (data: LeaveRequestInput) => void;
}

const LeaveForm: React.FC<LeaveFormProps> = ({ accountId, isAdmin = false, onClose, onSubmit }) => {
  const [formData, setFormData] = useState<LeaveRequestInput>({
    account_id: accountId,
    start_date: new Date().toISOString().split('T')[0],
    end_date: new Date().toISOString().split('T')[0],
    description: ''
  });

  const [accounts, setAccounts] = useState<Account[]>([]);
  const [loadingAccounts, setLoadingAccounts] = useState(false);

  useEffect(() => {
    if (isAdmin) {
      setLoadingAccounts(true);
      accountService.getAll()
        .then(data => {
          // Filter only active accounts
          const today = new Date().toISOString().split('T')[0];
          const active = (data as Account[]).filter(acc => !acc.end_date || acc.end_date > today);
          setAccounts(active);
        })
        .finally(() => setLoadingAccounts(false));
    }
  }, [isAdmin]);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => {
    const { name, value } = e.target;
    setFormData(prev => ({ ...prev, [name]: value }));
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    onSubmit(formData);
  };

  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black bg-opacity-40 p-4">
      <div className="bg-white rounded-md shadow-2xl w-full max-w-md overflow-hidden animate-in zoom-in duration-200">
        <div className="px-5 py-3 border-b border-gray-100 flex items-center justify-between">
          <div>
            <h3 className="text-sm font-bold text-[#006E62]">Pengajuan Libur Mandiri</h3>
            <p className="text-[9px] text-gray-400 font-bold uppercase tracking-widest">Khusus Jadwal Fleksibel / Dinamis</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600">
            <X size={18} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-5 space-y-4">
          {isAdmin && (
            <div className="space-y-1">
              <label htmlFor="account_id" className="text-[9px] font-bold text-gray-500 uppercase">Pilih Karyawan</label>
              <div className="relative">
                <select 
                  id="account_id"
                  required 
                  name="account_id" 
                  value={formData.account_id} 
                  onChange={handleChange} 
                  className="w-full px-2 py-1.5 text-xs border border-gray-200 rounded outline-none focus:ring-1 focus:ring-[#006E62] bg-white"
                >
                  <option value="">-- Pilih Karyawan --</option>
                  {accounts.map(acc => (
                    <option key={acc.id} value={acc.id}>{acc.full_name} ({acc.internal_nik})</option>
                  ))}
                </select>
                {loadingAccounts && <div className="absolute right-8 top-1/2 -translate-y-1/2"><div className="w-3 h-3 border-2 border-[#006E62] border-t-transparent rounded-full animate-spin"></div></div>}
              </div>
            </div>
          )}

          <div className="grid grid-cols-2 gap-4">
            <div className="space-y-1">
              <label htmlFor="start_date" className="text-[9px] font-bold text-gray-500 uppercase">Tanggal Awal</label>
              <div className="relative">
                <input 
                  id="start_date"
                  type="date"
                  required 
                  name="start_date" 
                  value={formData.start_date} 
                  onChange={handleChange} 
                  className="w-full px-2 py-1.5 text-xs border border-gray-200 rounded outline-none focus:ring-1 focus:ring-[#006E62] bg-gray-50" 
                />
              </div>
            </div>
            <div className="space-y-1">
              <label htmlFor="end_date" className="text-[9px] font-bold text-gray-500 uppercase">Tanggal Akhir</label>
              <div className="relative">
                <input 
                  id="end_date"
                  type="date"
                  required 
                  name="end_date" 
                  value={formData.end_date} 
                  onChange={handleChange} 
                  className="w-full px-2 py-1.5 text-xs border border-gray-200 rounded outline-none focus:ring-1 focus:ring-[#006E62] bg-gray-50" 
                />
              </div>
            </div>
          </div>

          <div className="space-y-1">
            <label htmlFor="description" className="text-[9px] font-bold text-gray-500 uppercase">Keterangan Libur</label>
            <textarea 
              id="description"
              required
              name="description" 
              value={formData.description} 
              onChange={handleChange} 
              rows={3} 
              placeholder="Alasan pengajuan libur..."
              className="w-full px-2 py-1.5 text-xs border border-gray-200 rounded outline-none focus:ring-1 focus:ring-[#006E62] resize-none" 
            />
          </div>

          <div className="flex justify-end gap-3 pt-2">
            <button type="button" onClick={onClose} className="text-[10px] font-bold text-gray-400 uppercase">Batal</button>
            <button 
              type="submit" 
              className="flex items-center gap-2 bg-[#006E62] text-white px-5 py-1.5 rounded text-[10px] font-bold uppercase shadow-md hover:bg-[#005a50]"
            >
              <Save size={12} /> Kirim Pengajuan
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};

export default LeaveForm;
