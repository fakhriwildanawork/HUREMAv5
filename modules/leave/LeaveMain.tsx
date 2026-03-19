import React, { useState, useEffect } from 'react';
import { Plane, Plus, Trash2, CheckCircle, XCircle, Clock, Search, Filter, Calendar } from 'lucide-react';
import Swal from 'sweetalert2';
import { leaveService } from '../../services/leaveService';
import { authService } from '../../services/authService';
import { LeaveRequestExtended } from '../../types';
import LeaveForm from '../account/LeaveForm';
import LoadingSpinner from '../../components/Common/LoadingSpinner';

const LeaveMain: React.FC = () => {
  const [leaves, setLeaves] = useState<LeaveRequestExtended[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState<'all' | 'pending' | 'approved' | 'rejected'>('all');
  
  const user = authService.getCurrentUser();
  const isAdmin = user?.role === 'admin' || user?.is_hr_admin || user?.is_performance_admin || user?.is_finance_admin;

  useEffect(() => {
    fetchData();
  }, []);

  const fetchData = async () => {
    try {
      setIsLoading(true);
      let data: LeaveRequestExtended[] = [];
      if (isAdmin) {
        data = await leaveService.getAll();
      } else if (user?.id) {
        data = await leaveService.getByAccountId(user.id);
      }
      setLeaves(data);
    } catch (error) {
      console.error(error);
      Swal.fire('Gagal', 'Gagal memuat data pengajuan libur', 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleCreateLeave = async (formData: any) => {
    setIsSaving(true);
    try {
      const created = await leaveService.create(formData);
      // Refresh data to get account details if admin, or just prepend
      if (isAdmin) {
        fetchData();
      } else {
        setLeaves(prev => [created as any, ...prev]);
      }
      setShowForm(false);
      
      const isAuto = created.status === 'approved';
      Swal.fire({ 
        title: 'Berhasil!', 
        text: isAuto ? 'Libur mandiri Anda langsung disetujui (Otomatis).' : 'Pengajuan libur telah dikirim dan menunggu verifikasi admin.', 
        icon: 'success' 
      });
    } catch (error) {
      Swal.fire('Gagal', 'Gagal mengirim pengajuan libur', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleUpdateStatus = async (id: string, status: 'approved' | 'rejected') => {
    const action = status === 'approved' ? 'Setujui' : 'Tolak';
    const result = await Swal.fire({
      title: `${action} pengajuan ini?`,
      icon: 'question',
      showCancelButton: true,
      confirmButtonColor: status === 'approved' ? '#006E62' : '#ef4444',
      confirmButtonText: `Ya, ${action}!`
    });

    if (result.isConfirmed) {
      setIsSaving(true);
      try {
        await leaveService.updateStatus(id, status);
        setLeaves(prev => prev.map(l => l.id === id ? { ...l, status } : l));
        Swal.fire('Berhasil!', `Pengajuan telah ${status === 'approved' ? 'disetujui' : 'ditolak'}.`, 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Gagal memperbarui status pengajuan', 'error');
      } finally {
        setIsSaving(false);
      }
    }
  };

  const handleDelete = async (id: string) => {
    const result = await Swal.fire({
      title: 'Batalkan pengajuan?',
      text: "Data ini akan dihapus permanen.",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#ef4444',
      confirmButtonText: 'Ya, Batalkan!'
    });

    if (result.isConfirmed) {
      setIsSaving(true);
      try {
        await leaveService.delete(id);
        setLeaves(prev => prev.filter(l => l.id !== id));
        Swal.fire('Dibatalkan!', 'Pengajuan telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Gagal menghapus pengajuan', 'error');
      } finally {
        setIsSaving(false);
      }
    }
  };

  const filteredLeaves = leaves.filter(l => {
    const matchesSearch = isAdmin 
      ? (l.account?.full_name?.toLowerCase().includes(searchTerm.toLowerCase()) || l.description.toLowerCase().includes(searchTerm.toLowerCase()))
      : l.description.toLowerCase().includes(searchTerm.toLowerCase());
    
    const matchesStatus = statusFilter === 'all' || l.status === statusFilter;
    
    return matchesSearch && matchesStatus;
  });

  const formatDate = (dateStr: string) => {
    return new Date(dateStr).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' });
  };

  const canRequestLeave = user?.schedule_type === 'Fleksibel' || user?.schedule_type === 'Shift Dinamis' || isAdmin;

  return (
    <div className="space-y-6">
      {isSaving && <LoadingSpinner />}
      
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-2xl font-bold text-gray-800 tracking-tight">Libur Mandiri</h2>
          <p className="text-sm text-gray-500">Kelola pengajuan libur untuk jadwal fleksibel & dinamis</p>
        </div>
        {canRequestLeave && (
          <button 
            onClick={() => setShowForm(true)}
            className="flex items-center justify-center gap-2 bg-[#006E62] text-white px-6 py-2.5 rounded-md font-bold text-xs uppercase shadow-lg hover:bg-[#005a50] transition-all"
          >
            <Plus size={16} /> Ajukan Libur
          </button>
        )}
      </div>

      <div className="bg-white p-4 rounded-md border border-gray-100 shadow-sm space-y-4">
        <div className="flex flex-col md:flex-row gap-4">
          <div className="flex-1 relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
            <input 
              type="text" 
              placeholder={isAdmin ? "Cari nama karyawan atau keterangan..." : "Cari keterangan..."}
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full pl-10 pr-4 py-2 text-sm border border-gray-200 rounded-md outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62]"
            />
          </div>
          <div className="flex items-center gap-2">
            <Filter size={16} className="text-gray-400" />
            <select 
              value={statusFilter}
              onChange={(e) => setStatusFilter(e.target.value as any)}
              className="text-sm border border-gray-200 rounded-md px-3 py-2 outline-none focus:ring-2 focus:ring-[#006E62]/20"
            >
              <option value="all">Semua Status</option>
              <option value="pending">Menunggu</option>
              <option value="approved">Disetujui</option>
              <option value="rejected">Ditolak</option>
            </select>
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="border-b border-gray-100">
                {isAdmin && <th className="py-3 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Karyawan</th>}
                <th className="py-3 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Periode Libur</th>
                <th className="py-3 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Keterangan</th>
                <th className="py-3 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Status</th>
                <th className="py-3 px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">Aksi</th>
              </tr>
            </thead>
            <tbody>
              {isLoading ? (
                <tr>
                  <td colSpan={isAdmin ? 5 : 4} className="py-10 text-center">
                    <div className="flex justify-center"><div className="w-6 h-6 border-2 border-[#006E62] border-t-transparent rounded-full animate-spin"></div></div>
                  </td>
                </tr>
              ) : filteredLeaves.length === 0 ? (
                <tr>
                  <td colSpan={isAdmin ? 5 : 4} className="py-10 text-center text-gray-400 text-sm italic">
                    Tidak ada data pengajuan libur.
                  </td>
                </tr>
              ) : (
                filteredLeaves.map((l) => (
                  <tr key={l.id} className="border-b border-gray-50 hover:bg-gray-50/50 transition-colors group">
                    {isAdmin && (
                      <td className="py-4 px-4">
                        <div className="flex items-center gap-3">
                          <div className="w-8 h-8 rounded-full bg-[#006E62]/10 flex items-center justify-center text-[#006E62] font-bold text-xs">
                            {l.account?.full_name?.charAt(0)}
                          </div>
                          <div>
                            <p className="text-xs font-bold text-gray-700">{l.account?.full_name}</p>
                            <p className="text-[10px] text-gray-400">{l.account?.internal_nik}</p>
                          </div>
                        </div>
                      </td>
                    )}
                    <td className="py-4 px-4">
                      <div className="flex items-center gap-2 text-xs font-medium text-gray-600">
                        <Calendar size={14} className="text-gray-400" />
                        <span>{formatDate(l.start_date)} - {formatDate(l.end_date)}</span>
                      </div>
                    </td>
                    <td className="py-4 px-4">
                      <p className="text-xs text-gray-600 line-clamp-2 max-w-xs italic">"{l.description}"</p>
                    </td>
                    <td className="py-4 px-4">
                      <span className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[10px] font-bold uppercase ${
                        l.status === 'approved' ? 'bg-emerald-50 text-emerald-600' : 
                        l.status === 'rejected' ? 'bg-red-50 text-red-600' : 
                        'bg-orange-50 text-orange-600'
                      }`}>
                        {l.status === 'approved' ? <CheckCircle size={12} /> : l.status === 'rejected' ? <XCircle size={12} /> : <Clock size={12} />}
                        {l.status === 'approved' ? 'ACC' : l.status === 'rejected' ? 'DITOLAK' : 'PENDING'}
                      </span>
                    </td>
                    <td className="py-4 px-4 text-right">
                      <div className="flex justify-end gap-2">
                        {isAdmin && l.status === 'pending' && (
                          <>
                            <button 
                              onClick={() => handleUpdateStatus(l.id, 'approved')}
                              className="p-1.5 text-emerald-600 hover:bg-emerald-50 rounded transition-colors"
                              title="Setujui"
                            >
                              <CheckCircle size={16} />
                            </button>
                            <button 
                              onClick={() => handleUpdateStatus(l.id, 'rejected')}
                              className="p-1.5 text-red-600 hover:bg-red-50 rounded transition-colors"
                              title="Tolak"
                            >
                              <XCircle size={16} />
                            </button>
                          </>
                        )}
                        {(isAdmin || l.status === 'pending') && (
                          <button 
                            onClick={() => handleDelete(l.id)}
                            className="p-1.5 text-gray-400 hover:text-red-500 hover:bg-red-50 rounded transition-colors"
                            title="Hapus/Batalkan"
                          >
                            <Trash2 size={16} />
                          </button>
                        )}
                      </div>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {showForm && user && (
        <LeaveForm 
          accountId={user.id} 
          isAdmin={isAdmin}
          onClose={() => setShowForm(false)} 
          onSubmit={handleCreateLeave} 
        />
      )}
    </div>
  );
};

export default LeaveMain;
