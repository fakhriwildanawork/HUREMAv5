import React, { useState, useEffect } from 'react';
import { Search, Filter, Clock, CheckCircle2, XCircle, AlertCircle, Eye, Loader2, Calendar, User, ArrowRight, LucideIcon } from 'lucide-react';
import { submissionService } from '../../services/submissionService';
import { Submission, AuthUser, SubmissionStatus } from '../../types';
import Swal from 'sweetalert2';
import SubmissionDetail from '../submission/SubmissionDetail';

interface AdminSubmissionModuleProps {
  user: AuthUser;
  type: string;
  title: string;
  subtitle: string;
  icon: LucideIcon;
}

const AdminSubmissionModule: React.FC<AdminSubmissionModuleProps> = ({ user, type, title, subtitle, icon: Icon }) => {
  const [submissions, setSubmissions] = useState<Submission[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState<string>('ALL');
  const [selectedSubmission, setSelectedSubmission] = useState<Submission | null>(null);
  const [showDetail, setShowDetail] = useState(false);

  useEffect(() => {
    fetchSubmissions();
  }, [type]);

  const fetchSubmissions = async () => {
    try {
      setIsLoading(true);
      const data = await submissionService.getByType(type);
      setSubmissions(data);
    } catch (error) {
      console.error(error);
    } finally {
      setIsLoading(false);
    }
  };

  const filteredSubmissions = submissions.filter(sub => {
    const matchesSearch = 
      sub.account?.full_name.toLowerCase().includes(searchTerm.toLowerCase()) ||
      sub.account?.internal_nik.toLowerCase().includes(searchTerm.toLowerCase());
    const matchesStatus = statusFilter === 'ALL' || sub.status === statusFilter;
    return matchesSearch && matchesStatus;
  });

  const getStatusBadge = (status: string) => {
    switch (status) {
      case 'Disetujui':
        return <span className="px-2 py-1 bg-emerald-100 text-emerald-700 rounded-lg text-[10px] font-bold uppercase tracking-wider flex items-center gap-1 w-fit"><CheckCircle2 size={10} /> Disetujui</span>;
      case 'Ditolak':
        return <span className="px-2 py-1 bg-rose-100 text-rose-700 rounded-lg text-[10px] font-bold uppercase tracking-wider flex items-center gap-1 w-fit"><XCircle size={10} /> Ditolak</span>;
      default:
        return <span className="px-2 py-1 bg-blue-100 text-blue-700 rounded-lg text-[10px] font-bold uppercase tracking-wider flex items-center gap-1 w-fit"><Clock size={10} /> Pending</span>;
    }
  };

  const handleVerify = async (id: string, status: SubmissionStatus, notes?: string) => {
    try {
      const result = await Swal.fire({
        title: `Konfirmasi ${status}`,
        text: `Apakah Anda yakin ingin ${status.toLowerCase()} pengajuan ini?`,
        icon: 'question',
        showCancelButton: true,
        confirmButtonColor: status === 'Disetujui' ? '#006E62' : '#ef4444',
        cancelButtonColor: '#6b7280',
        confirmButtonText: 'Ya, Lanjutkan',
        cancelButtonText: 'Batal'
      });

      if (result.isConfirmed) {
        await submissionService.verify(id, status, user.id, notes);
        await Swal.fire({
          title: 'Berhasil!',
          text: `Pengajuan telah ${status.toLowerCase()}.`,
          icon: 'success',
          timer: 1500,
          showConfirmButton: false
        });
        setShowDetail(false);
        setSelectedSubmission(null);
        fetchSubmissions();
      }
    } catch (error) {
      console.error(error);
      Swal.fire('Error', 'Gagal memproses verifikasi', 'error');
    }
  };

  if (isLoading) return (
    <div className="flex flex-col items-center justify-center py-20 text-gray-400">
      <div className="w-12 h-12 border-4 border-gray-200 border-t-[#006E62] rounded-full animate-spin mb-4"></div>
      <p className="text-xs font-bold uppercase tracking-widest">Memuat Data {title}...</p>
    </div>
  );

  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      {/* Header */}
      <div className="bg-white p-6 rounded-2xl border border-gray-100 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-center gap-4">
          <div className="w-12 h-12 bg-[#006E62]/10 rounded-xl flex items-center justify-center text-[#006E62]">
            <Icon size={28} />
          </div>
          <div>
            <h2 className="text-xl font-bold text-gray-800 tracking-tight">{title}</h2>
            <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">{subtitle}</p>
          </div>
        </div>

        <div className="flex flex-col sm:flex-row gap-3">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
            <input 
              type="text"
              placeholder="Cari nama atau NIK..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="pl-10 pr-4 py-2.5 bg-gray-50 border border-gray-100 rounded-xl text-xs focus:ring-2 focus:ring-[#006E62] outline-none w-full sm:w-64 transition-all"
            />
          </div>
          <select 
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value)}
            className="px-4 py-2.5 bg-gray-50 border border-gray-100 rounded-xl text-xs font-bold text-gray-600 outline-none focus:ring-2 focus:ring-[#006E62] transition-all"
          >
            <option value="ALL">SEMUA STATUS</option>
            <option value="Pending">PENDING</option>
            <option value="Disetujui">DISETUJUI</option>
            <option value="Ditolak">DITOLAK</option>
          </select>
        </div>
      </div>

      {/* Table */}
      <div className="bg-white rounded-2xl border border-gray-100 shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-gray-50/50 border-b border-gray-100">
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Pegawai</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tanggal Pengajuan</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Keterangan</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-center">Status</th>
                <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {filteredSubmissions.length === 0 ? (
                <tr>
                  <td colSpan={5} className="px-6 py-12 text-center text-gray-400 italic text-xs">Tidak ada data pengajuan yang ditemukan.</td>
                </tr>
              ) : (
                filteredSubmissions.map((sub) => (
                  <tr key={sub.id} className="hover:bg-gray-50/50 transition-colors group">
                    <td className="px-6 py-4">
                      <div className="flex items-center gap-3">
                        <div className="w-8 h-8 bg-gray-100 rounded-lg flex items-center justify-center text-gray-400">
                          <User size={16} />
                        </div>
                        <div>
                          <p className="text-xs font-bold text-gray-800">{sub.account?.full_name}</p>
                          <p className="text-[10px] text-gray-400 font-medium">{sub.account?.internal_nik}</p>
                        </div>
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <div className="flex items-center gap-2">
                        <Calendar size={14} className="text-gray-400" />
                        <span className="text-xs font-bold text-gray-700">{new Date(sub.created_at).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' })}</span>
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <p className="text-xs text-gray-600 line-clamp-1">{sub.submission_data.reason || sub.submission_data.notes || '-'}</p>
                    </td>
                    <td className="px-6 py-4">
                      <div className="flex justify-center">
                        {getStatusBadge(sub.status)}
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <div className="flex items-center justify-end">
                        <button 
                          onClick={() => {
                            setSelectedSubmission(sub);
                            setShowDetail(true);
                          }}
                          className="flex items-center gap-2 bg-white border border-gray-200 text-gray-600 px-4 py-2 rounded-xl text-[10px] font-bold uppercase tracking-wider hover:bg-gray-50 hover:border-gray-300 transition-all shadow-sm active:scale-95"
                        >
                          <Eye size={14} />
                          Detail & Verifikasi
                        </button>
                      </div>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {showDetail && selectedSubmission && (
        <SubmissionDetail 
          submission={selectedSubmission}
          onClose={() => {
            setShowDetail(false);
            setSelectedSubmission(null);
          }}
          onVerify={handleVerify}
          canVerify={selectedSubmission.status === 'Pending'}
        />
      )}
    </div>
  );
};

export default AdminSubmissionModule;
