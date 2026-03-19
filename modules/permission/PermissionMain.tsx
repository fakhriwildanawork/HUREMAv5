import React, { useState, useEffect } from 'react';
import { Plus, Search, Calendar, Filter, Clock, CheckCircle2, XCircle, MessageSquare, ChevronRight, ClipboardList } from 'lucide-react';
import { permissionService } from '../../services/permissionService';
import { authService } from '../../services/authService';
import { PermissionRequest, AuthUser } from '../../types';
import PermissionForm from './PermissionForm';
import PermissionDetail from './PermissionDetail';
import Swal from 'sweetalert2';

const PermissionMain: React.FC = () => {
  const [requests, setRequests] = useState<PermissionRequest[]>([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [selectedRequest, setSelectedRequest] = useState<PermissionRequest | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [filterStatus, setFilterStatus] = useState<string>('all');
  
  const user = authService.getCurrentUser();
  const isAdmin = user?.role === 'admin' || user?.is_hr_admin || user?.is_performance_admin || user?.is_finance_admin;

  useEffect(() => {
    fetchRequests();
  }, []);

  const fetchRequests = async () => {
    try {
      setLoading(true);
      const data = isAdmin 
        ? await permissionService.getAll() 
        : await permissionService.getByAccountId(user?.id || '');
      setRequests(data);
    } catch (error) {
      console.error('Error fetching permissions:', error);
    } finally {
      setLoading(false);
    }
  };

  const handleCreate = async (input: any) => {
    try {
      await permissionService.create(input);
      setShowForm(false);
      fetchRequests();
      Swal.fire({
        title: 'Berhasil!',
        text: 'Pengajuan izin telah dikirim.',
        icon: 'success',
        timer: 1500,
        showConfirmButton: false
      });
    } catch (error: any) {
      Swal.fire('Gagal', error.message || 'Terjadi kesalahan.', 'error');
    }
  };

  const getStatusColor = (status: string) => {
    switch (status) {
      case 'approved': return 'bg-emerald-50 text-emerald-600 border-emerald-100';
      case 'rejected': return 'bg-rose-50 text-rose-600 border-rose-100';
      case 'negotiating': return 'bg-amber-50 text-amber-600 border-amber-100';
      case 'cancelled': return 'bg-gray-50 text-gray-500 border-gray-100';
      default: return 'bg-blue-50 text-blue-600 border-blue-100';
    }
  };

  const getStatusIcon = (status: string) => {
    switch (status) {
      case 'approved': return <CheckCircle2 size={12} />;
      case 'rejected': return <XCircle size={12} />;
      case 'negotiating': return <MessageSquare size={12} />;
      case 'cancelled': return <XCircle size={12} />;
      default: return <Clock size={12} />;
    }
  };

  const filteredRequests = requests.filter(req => {
    const matchesSearch = isAdmin 
      ? req.account?.full_name.toLowerCase().includes(searchTerm.toLowerCase()) || 
        req.account?.internal_nik.includes(searchTerm)
      : req.description.toLowerCase().includes(searchTerm.toLowerCase()) ||
        req.permission_type.toLowerCase().includes(searchTerm.toLowerCase());
    
    const matchesStatus = filterStatus === 'all' || req.status === filterStatus;
    
    return matchesSearch && matchesStatus;
  });

  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      {/* Header Section */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-2xl font-bold text-gray-800 tracking-tight">Modul Izin</h2>
          <p className="text-[10px] font-bold text-gray-400 uppercase tracking-[0.2em] mt-1">
            {isAdmin ? 'Manajemen Pengajuan Izin Karyawan' : 'Daftar Pengajuan Izin Anda'}
          </p>
        </div>
        {!isAdmin && (
          <button 
            onClick={() => setShowForm(true)}
            className="flex items-center justify-center gap-2 bg-[#006E62] text-white px-6 py-3 rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-[#005a50] shadow-lg shadow-[#006E62]/20 transition-all group"
          >
            <Plus size={18} className="group-hover:rotate-90 transition-transform duration-300" />
            Buat Pengajuan Izin
          </button>
        )}
      </div>

      {/* Filters & Search */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        <div className="md:col-span-2 relative">
          <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
          <input 
            type="text"
            placeholder={isAdmin ? "Cari nama atau NIK karyawan..." : "Cari alasan atau jenis izin..."}
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="w-full pl-12 pr-4 py-3 bg-white border border-gray-100 rounded-2xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/10 focus:border-[#006E62] transition-all text-sm font-medium"
          />
        </div>
        <div className="relative">
          <Filter className="absolute left-4 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
          <select 
            value={filterStatus}
            onChange={(e) => setFilterStatus(e.target.value)}
            className="w-full pl-12 pr-4 py-3 bg-white border border-gray-100 rounded-2xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/10 focus:border-[#006E62] transition-all text-sm font-bold appearance-none"
          >
            <option value="all">SEMUA STATUS</option>
            <option value="pending">PENDING</option>
            <option value="negotiating">NEGOSIASI</option>
            <option value="approved">DISETUJUI</option>
            <option value="rejected">DITOLAK</option>
            <option value="cancelled">DIBATALKAN</option>
          </select>
        </div>
      </div>

      {/* Content Section */}
      {loading ? (
        <div className="flex flex-col items-center justify-center py-20">
          <div className="w-12 h-12 border-4 border-[#006E62] border-t-transparent rounded-full animate-spin"></div>
          <p className="mt-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Memuat data...</p>
        </div>
      ) : filteredRequests.length > 0 ? (
        <div className="grid grid-cols-1 gap-4">
          {filteredRequests.map((req) => (
            <div 
              key={req.id}
              onClick={() => setSelectedRequest(req)}
              className="group bg-white border border-gray-100 p-5 rounded-2xl hover:shadow-xl hover:shadow-gray-200/50 transition-all cursor-pointer relative overflow-hidden"
            >
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div className="flex items-start gap-4">
                  <div className={`w-12 h-12 rounded-xl flex items-center justify-center shrink-0 ${
                    req.status === 'approved' ? 'bg-emerald-50 text-emerald-600' : 
                    req.status === 'rejected' ? 'bg-rose-50 text-rose-600' : 'bg-gray-50 text-gray-400'
                  }`}>
                    <ClipboardList size={24} />
                  </div>
                  <div>
                    <div className="flex items-center gap-2 mb-1">
                      <span className="text-sm font-bold text-gray-800">{req.permission_type}</span>
                      <div className={`flex items-center gap-1 px-2 py-0.5 rounded-full border text-[8px] font-bold uppercase tracking-wider ${getStatusColor(req.status)}`}>
                        {getStatusIcon(req.status)}
                        {req.status}
                      </div>
                    </div>
                    {isAdmin && (
                      <div className="text-[10px] font-bold text-[#006E62] uppercase tracking-widest mb-1">
                        {req.account?.full_name} • {req.account?.internal_nik}
                      </div>
                    )}
                    <div className="flex items-center gap-4 text-[10px] text-gray-400 font-medium">
                      <div className="flex items-center gap-1">
                        <Calendar size={12} />
                        {new Date(req.start_date).toLocaleDateString('id-ID')} - {new Date(req.end_date).toLocaleDateString('id-ID')}
                      </div>
                      <div className="flex items-center gap-1">
                        <Clock size={12} />
                        {new Date(req.created_at || '').toLocaleDateString('id-ID')}
                      </div>
                    </div>
                  </div>
                </div>
                
                <div className="flex items-center gap-4">
                  <div className="text-right hidden md:block">
                    <div className="text-[9px] font-bold text-gray-300 uppercase tracking-widest mb-1">Giliran Respons</div>
                    <div className={`text-[10px] font-bold uppercase ${
                      req.current_negotiator_role === (isAdmin ? 'admin' : 'user') ? 'text-amber-500' : 'text-gray-400'
                    }`}>
                      {req.current_negotiator_role === (isAdmin ? 'admin' : 'user') ? 'ANDA' : req.current_negotiator_role}
                    </div>
                  </div>
                  <ChevronRight size={20} className="text-gray-300 group-hover:text-[#006E62] group-hover:translate-x-1 transition-all" />
                </div>
              </div>
              
              {/* Progress bar for negotiation */}
              <div className="absolute bottom-0 left-0 h-1 bg-[#006E62] transition-all duration-500" style={{ 
                width: req.status === 'approved' ? '100%' : req.status === 'rejected' ? '100%' : '30%' 
              }} />
            </div>
          ))}
        </div>
      ) : (
        <div className="flex flex-col items-center justify-center py-20 bg-gray-50 rounded-3xl border border-dashed border-gray-200">
          <ClipboardList size={48} className="text-gray-200 mb-4" />
          <p className="text-sm font-medium text-gray-400">Tidak ada data pengajuan izin.</p>
        </div>
      )}

      {/* Modals */}
      {showForm && user && (
        <PermissionForm 
          accountId={user.id}
          onClose={() => setShowForm(false)}
          onSubmit={handleCreate}
        />
      )}

      {selectedRequest && user && (
        <PermissionDetail 
          request={selectedRequest}
          user={user}
          onClose={() => setSelectedRequest(null)}
          onUpdate={fetchRequests}
        />
      )}
    </div>
  );
};

export default PermissionMain;
