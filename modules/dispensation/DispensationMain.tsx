import React, { useState, useEffect } from 'react';
import { ClipboardList, Plus, Search, Filter, Clock, CheckCircle2, XCircle, AlertCircle, Eye, Edit2, Trash2, FileText, Calendar } from 'lucide-react';
import { dispensationService } from '../../services/dispensationService';
import { DispensationRequest, AuthUser } from '../../types';
import { authService } from '../../services/authService';
import Swal from 'sweetalert2';
import DispensationForm from './DispensationForm';
import DispensationDetail from './components/DispensationDetail';

interface DispensationMainProps {
  user: AuthUser;
}

const DispensationMain: React.FC<DispensationMainProps> = ({ user }) => {
  const [requests, setRequests] = useState<DispensationRequest[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [selectedRequest, setSelectedRequest] = useState<DispensationRequest | null>(null);
  const [showDetail, setShowDetail] = useState(false);
  const [editingRequest, setEditingRequest] = useState<DispensationRequest | null>(null);

  useEffect(() => {
    if (user?.id) {
      fetchRequests();
    }
  }, [user?.id]);

  const fetchRequests = async () => {
    try {
      setIsLoading(true);
      const data = await dispensationService.getByAccountId(user!.id);
      setRequests(data);
    } catch (error) {
      console.error(error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleDelete = async (id: string) => {
    const result = await Swal.fire({
      title: 'Apakah Anda yakin?',
      text: "Data pengajuan akan dihapus permanen!",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
      cancelButtonColor: '#3085d6',
      confirmButtonText: 'Ya, Hapus!',
      cancelButtonText: 'Batal'
    });

    if (result.isConfirmed) {
      try {
        await dispensationService.delete(id);
        setRequests(prev => prev.filter(r => r.id !== id));
        Swal.fire('Terhapus!', 'Pengajuan telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Terjadi kesalahan saat menghapus data.', 'error');
      }
    }
  };

  const getStatusBadge = (status: string) => {
    switch (status) {
      case 'APPROVED':
        return <span className="px-2 py-1 bg-emerald-100 text-emerald-700 rounded-lg text-[10px] font-bold uppercase tracking-wider flex items-center gap-1 w-fit"><CheckCircle2 size={10} /> Disetujui</span>;
      case 'REJECTED':
        return <span className="px-2 py-1 bg-rose-100 text-rose-700 rounded-lg text-[10px] font-bold uppercase tracking-wider flex items-center gap-1 w-fit"><XCircle size={10} /> Ditolak</span>;
      case 'PARTIAL':
        return <span className="px-2 py-1 bg-amber-100 text-amber-700 rounded-lg text-[10px] font-bold uppercase tracking-wider flex items-center gap-1 w-fit"><AlertCircle size={10} /> Sebagian</span>;
      default:
        return <span className="px-2 py-1 bg-blue-100 text-blue-700 rounded-lg text-[10px] font-bold uppercase tracking-wider flex items-center gap-1 w-fit"><Clock size={10} /> Pending</span>;
    }
  };

  if (isLoading) return (
    <div className="flex flex-col items-center justify-center py-20 text-gray-400">
      <div className="w-12 h-12 border-4 border-gray-200 border-t-[#006E62] rounded-full animate-spin mb-4"></div>
      <p className="text-xs font-bold uppercase tracking-widest">Memuat Data Dispensasi...</p>
    </div>
  );

  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 bg-white p-6 rounded-2xl border border-gray-100 shadow-sm">
        <div className="flex items-center gap-4">
          <div className="w-12 h-12 bg-[#006E62]/10 rounded-xl flex items-center justify-center text-[#006E62]">
            <ClipboardList size={28} />
          </div>
          <div>
            <h2 className="text-xl font-bold text-gray-800 tracking-tight">Dispensasi Presensi</h2>
            <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Pengajuan Koreksi Data Kehadiran Bermasalah</p>
          </div>
        </div>
        <button 
          onClick={() => {
            setEditingRequest(null);
            setShowForm(true);
          }}
          className="flex items-center justify-center gap-2 bg-[#006E62] text-white px-6 py-3 rounded-xl font-bold text-xs uppercase tracking-wider hover:bg-[#005c52] transition-all shadow-lg shadow-[#006E62]/20 active:scale-95"
        >
          <Plus size={18} />
          Buat Pengajuan
        </button>
      </div>

      {/* Stats Summary */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        {[
          { label: 'Total', value: requests.length, color: 'bg-gray-500' },
          { label: 'Pending', value: requests.filter(r => r.status === 'PENDING').length, color: 'bg-blue-500' },
          { label: 'Disetujui', value: requests.filter(r => r.status === 'APPROVED' || r.status === 'PARTIAL').length, color: 'bg-emerald-500' },
          { label: 'Ditolak', value: requests.filter(r => r.status === 'REJECTED').length, color: 'bg-rose-500' },
        ].map((stat, i) => (
          <div key={i} className="bg-white p-4 rounded-2xl border border-gray-100 shadow-sm flex items-center justify-between">
            <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">{stat.label}</span>
            <span className={`text-lg font-black ${stat.color.replace('bg-', 'text-')}`}>{stat.value}</span>
          </div>
        ))}
      </div>

      {/* List Table */}
      <div className="bg-white rounded-2xl border border-gray-100 shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-gray-50/50 border-b border-gray-100">
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tanggal</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Masalah</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Alasan</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Status</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {requests.length === 0 ? (
                <tr>
                  <td colSpan={5} className="px-6 py-12 text-center text-gray-400 italic text-xs">Belum ada riwayat pengajuan dispensasi.</td>
                </tr>
              ) : (
                requests.map((req) => (
                  <tr key={req.id} className="hover:bg-gray-50/50 transition-colors group">
                    <td className="px-6 py-4">
                      <div className="flex items-center gap-2">
                        <Calendar size={14} className="text-gray-400" />
                        <span className="text-xs font-bold text-gray-700">{new Date(req.date).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' })}</span>
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <div className="flex flex-wrap gap-1">
                        {req.issues.map((issue, idx) => (
                          <span key={idx} className={`px-1.5 py-0.5 rounded text-[9px] font-bold uppercase ${
                            issue.status === 'APPROVED' ? 'bg-emerald-50 text-emerald-600' :
                            issue.status === 'REJECTED' ? 'bg-rose-50 text-rose-600' :
                            'bg-blue-50 text-blue-600'
                          }`}>
                            {issue.type.replace('_', ' ')}
                          </span>
                        ))}
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <p className="text-xs text-gray-500 line-clamp-1 max-w-[200px]">{req.reason}</p>
                    </td>
                    <td className="px-6 py-4">
                      <div className="flex justify-center">
                        {getStatusBadge(req.status)}
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <div className="flex items-center justify-end gap-2">
                        <button 
                          onClick={() => {
                            setSelectedRequest(req);
                            setShowDetail(true);
                          }}
                          className="p-2 text-gray-400 hover:text-[#006E62] hover:bg-[#006E62]/5 rounded-lg transition-all"
                          title="Detail"
                        >
                          <Eye size={16} />
                        </button>
                        {req.status === 'PENDING' && (
                          <>
                            <button 
                              onClick={() => {
                                setEditingRequest(req);
                                setShowForm(true);
                              }}
                              className="p-2 text-gray-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg transition-all"
                              title="Edit"
                            >
                              <Edit2 size={16} />
                            </button>
                            <button 
                              onClick={() => handleDelete(req.id)}
                              className="p-2 text-gray-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-all"
                              title="Hapus"
                            >
                              <Trash2 size={16} />
                            </button>
                          </>
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

      {/* Modals */}
      {showForm && (
        <DispensationForm 
          onClose={() => {
            setShowForm(false);
            setEditingRequest(null);
          }}
          onSuccess={() => {
            setShowForm(false);
            setEditingRequest(null);
            fetchRequests();
          }}
          editData={editingRequest}
        />
      )}

      {showDetail && selectedRequest && (
        <DispensationDetail 
          request={selectedRequest}
          onClose={() => {
            setShowDetail(false);
            setSelectedRequest(null);
          }}
          isAdmin={false}
        />
      )}
    </div>
  );
};

export default DispensationMain;
