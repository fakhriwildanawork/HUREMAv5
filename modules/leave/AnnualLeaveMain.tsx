import React, { useState, useEffect } from 'react';
import { Plane, Plus, Search, Calendar, History, MessageSquare, CheckCircle2, XCircle, Clock, ArrowRight, FileText, User, Info } from 'lucide-react';
import Swal from 'sweetalert2';
import { leaveService } from '../../services/leaveService';
import { authService } from '../../services/authService';
import { accountService } from '../../services/accountService';
import { AnnualLeaveRequest, AuthUser, Account } from '../../types';
import AnnualLeaveForm from './AnnualLeaveForm';
import AnnualLeaveDetail from './AnnualLeaveDetail';
import LoadingSpinner from '../../components/Common/LoadingSpinner';
import { CardSkeleton } from '../../components/Common/Skeleton';

const AnnualLeaveMain: React.FC = () => {
  const [requests, setRequests] = useState<AnnualLeaveRequest[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [showForm, setShowForm] = useState(false);
  const [selectedRequest, setSelectedRequest] = useState<AnnualLeaveRequest | null>(null);
  const [user, setUser] = useState<AuthUser | null>(null);
  const [account, setAccount] = useState<Account | null>(null);
  const [usedDays, setUsedDays] = useState(0);

  const isAdmin = user?.role === 'admin' || user?.is_hr_admin || user?.is_performance_admin || user?.is_finance_admin;

  useEffect(() => {
    const currentUser = authService.getCurrentUser();
    setUser(currentUser);
    fetchRequests(currentUser);
  }, []);

  const fetchRequests = async (currentUser: AuthUser | null) => {
    if (!currentUser) return;
    try {
      setIsLoading(true);
      const [leaveData, accountData, used] = await Promise.all([
        isAdmin 
          ? leaveService.getAllAnnual() 
          : leaveService.getAnnualByAccountId(currentUser.id),
        accountService.getById(currentUser.id),
        leaveService.getUsedAnnualLeaveDays(currentUser.id, new Date().getFullYear())
      ]);
      
      setRequests(leaveData);
      setAccount(accountData);
      setUsedDays(used);
    } catch (error) {
      console.error(error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleCreate = async (input: any) => {
    try {
      setIsSaving(true);
      await leaveService.createAnnual(input);
      await fetchRequests(user);
      setShowForm(false);
      Swal.fire({
        title: 'Berhasil!',
        text: 'Pengajuan cuti telah dikirim.',
        icon: 'success',
        timer: 1500,
        showConfirmButton: false
      });
    } catch (error) {
      Swal.fire('Gagal', 'Terjadi kesalahan saat mengirim pengajuan.', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const filteredRequests = requests.filter(r => {
    const searchStr = `${r.account?.full_name || ''} ${r.description}`.toLowerCase();
    return searchStr.includes(searchTerm.toLowerCase());
  });

  const getStatusBadge = (status: string) => {
    switch (status) {
      case 'approved': return <span className="px-2 py-0.5 bg-emerald-50 text-emerald-600 rounded text-[9px] font-bold uppercase">Disetujui</span>;
      case 'rejected': return <span className="px-2 py-0.5 bg-rose-50 text-rose-600 rounded text-[9px] font-bold uppercase">Ditolak</span>;
      case 'negotiating': return <span className="px-2 py-0.5 bg-amber-50 text-amber-600 rounded text-[9px] font-bold uppercase">Negosiasi</span>;
      case 'cancelled': return <span className="px-2 py-0.5 bg-gray-100 text-gray-400 rounded text-[9px] font-bold uppercase">Dibatalkan</span>;
      default: return <span className="px-2 py-0.5 bg-blue-50 text-blue-600 rounded text-[9px] font-bold uppercase">Pending</span>;
    }
  };

  return (
    <div className="space-y-6">
      {isSaving && <LoadingSpinner message="Memproses Data..." />}

      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-xl font-bold text-gray-800 tracking-tight">Cuti Tahunan</h2>
          <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Manajemen & Negosiasi Jadwal Cuti</p>
        </div>
        
        <div className="flex items-center gap-3">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
            <input
              type="text"
              placeholder="Cari pengajuan..."
              className="pl-9 pr-4 py-2 bg-gray-50 border border-gray-100 rounded-md focus:outline-none focus:ring-2 focus:ring-[#006E62] text-xs w-64"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
          {!isAdmin && (
            <button 
              onClick={() => setShowForm(true)}
              className="flex items-center gap-2 bg-[#006E62] text-white px-4 py-2 rounded-md hover:bg-[#005a50] transition-colors shadow-sm text-xs font-bold uppercase"
            >
              <Plus size={16} />
              Ajukan Cuti
            </button>
          )}
        </div>
      </div>

      {!isAdmin && account && (
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
          <div className="bg-white p-4 rounded-xl border border-gray-100 shadow-sm">
            <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Jatah Dasar</p>
            <p className="text-xl font-bold text-gray-800">{account.leave_quota} <span className="text-xs font-medium text-gray-400">Hari</span></p>
          </div>
          <div className="bg-white p-4 rounded-xl border border-gray-100 shadow-sm">
            <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Carry-over</p>
            <p className="text-xl font-bold text-emerald-600">+{account.carry_over_quota} <span className="text-xs font-medium text-gray-400">Hari</span></p>
          </div>
          <div className="bg-white p-4 rounded-xl border border-gray-100 shadow-sm">
            <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Total Jatah</p>
            <p className="text-xl font-bold text-[#006E62]">{account.leave_quota + account.carry_over_quota} <span className="text-xs font-medium text-gray-400">Hari</span></p>
          </div>
          <div className="bg-white p-4 rounded-xl border border-gray-100 shadow-sm">
            <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Sisa Jatah</p>
            <p className="text-xl font-bold text-orange-600">{(account.leave_quota + account.carry_over_quota) - usedDays} <span className="text-xs font-medium text-gray-400">Hari</span></p>
          </div>
        </div>
      )}

      {isLoading ? (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {[1, 2, 3].map(i => <CardSkeleton key={i} />)}
        </div>
      ) : filteredRequests.length === 0 ? (
        <div className="flex flex-col items-center justify-center py-20 text-gray-400 bg-gray-50 rounded-xl border border-dashed border-gray-200">
          <Plane size={48} strokeWidth={1} className="mb-4 opacity-20" />
          <p className="text-sm font-bold uppercase tracking-widest">Belum ada pengajuan cuti</p>
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {filteredRequests.map(request => (
            <div 
              key={request.id}
              onClick={() => setSelectedRequest(request)}
              className="bg-white border border-gray-100 p-5 rounded-xl shadow-sm hover:shadow-md transition-all cursor-pointer border-l-4 border-l-transparent hover:border-l-[#006E62] group"
            >
              <div className="flex justify-between items-start mb-4">
                <div className="flex flex-col">
                  {isAdmin && (
                    <div className="flex items-center gap-1.5 mb-1">
                      <User size={10} className="text-gray-400" />
                      <span className="text-[10px] font-bold text-[#006E62] uppercase">{request.account?.full_name}</span>
                    </div>
                  )}
                  <div className="flex items-center gap-2 text-gray-800 font-bold text-sm">
                    <Calendar size={14} className="text-gray-400" />
                    <span>{new Date(request.start_date).toLocaleDateString('id-ID', { day: 'numeric', month: 'short' })}</span>
                    <ArrowRight size={12} className="text-gray-300" />
                    <span>{new Date(request.end_date).toLocaleDateString('id-ID', { day: 'numeric', month: 'short' })}</span>
                  </div>
                </div>
                {getStatusBadge(request.status)}
              </div>

              <p className="text-[11px] text-gray-500 line-clamp-2 mb-4 italic">"{request.description}"</p>

              <div className="flex items-center justify-between pt-4 border-t border-gray-50">
                <div className="flex items-center gap-3">
                  <div className="flex items-center gap-1 text-[9px] font-bold text-gray-400 uppercase">
                    <History size={12} /> {request.negotiation_data.length} Nego
                  </div>
                  {request.file_id && (
                    <div className="flex items-center gap-1 text-[9px] font-bold text-emerald-500 uppercase">
                      <FileText size={12} /> Dokumen
                    </div>
                  )}
                </div>
                {request.current_negotiator_role === (isAdmin ? 'admin' : 'user') && request.status !== 'approved' && request.status !== 'rejected' && request.status !== 'cancelled' && (
                  <div className="flex items-center gap-1 text-[9px] font-bold text-amber-500 animate-pulse uppercase">
                    <MessageSquare size={12} /> Butuh Respon
                  </div>
                )}
              </div>
            </div>
          ))}
        </div>
      )}

      {showForm && user && account && (
        <AnnualLeaveForm 
          accountId={user.id} 
          leaveQuota={account.leave_quota + account.carry_over_quota}
          usedDays={usedDays}
          onClose={() => setShowForm(false)} 
          onSubmit={handleCreate} 
        />
      )}

      {selectedRequest && user && (
        <AnnualLeaveDetail 
          request={selectedRequest} 
          user={user}
          onClose={() => setSelectedRequest(null)} 
          onUpdate={() => fetchRequests(user)}
        />
      )}
    </div>
  );
};

export default AnnualLeaveMain;
