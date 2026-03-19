import React, { useState, useEffect } from 'react';
import { ClipboardCheck, Search, Filter, Plus, Clock, CheckCircle2, XCircle, ListFilter, LayoutGrid } from 'lucide-react';
import Swal from 'sweetalert2';
import { submissionService } from '../../services/submissionService';
import { authService } from '../../services/authService';
import { Submission } from '../../types';
import SubmissionForm from './SubmissionForm';
import SubmissionDetail from './SubmissionDetail';
import LoadingSpinner from '../../components/Common/LoadingSpinner';
import { CardSkeleton } from '../../components/Common/Skeleton';

const SubmissionMain: React.FC = () => {
  const [submissions, setSubmissions] = useState<Submission[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [activeTab, setActiveTab] = useState<'verification' | 'monitoring' | 'history'>('verification');
  const [searchTerm, setSearchTerm] = useState('');
  const [showForm, setShowForm] = useState(false);
  const [selectedSubmission, setSelectedSubmission] = useState<Submission | null>(null);


  const currentUser = authService.getCurrentUser();

  useEffect(() => {
    fetchSubmissions();
  }, []);

  const fetchSubmissions = async () => {
    try {
      setIsLoading(true);
      const data = await submissionService.getAll();
      setSubmissions(data);
    } catch (error) {
      console.error(error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleVerify = async (id: string, status: 'Disetujui' | 'Ditolak', notes?: string) => {
    if (!currentUser) return;
    
    const confirm = await Swal.fire({
      title: `Konfirmasi ${status}?`,
      text: notes ? `Catatan: ${notes}` : "Aksi ini akan memperbarui status pengajuan karyawan.",
      icon: status === 'Disetujui' ? 'success' : 'warning',
      showCancelButton: true,
      confirmButtonColor: status === 'Disetujui' ? '#006E62' : '#ef4444',
      confirmButtonText: 'Ya, Proses'
    });

    if (confirm.isConfirmed) {
      try {
        setIsSaving(true);
        await submissionService.verify(id, status, currentUser.id, notes);
        await fetchSubmissions();
        setSelectedSubmission(null);
        Swal.fire('Berhasil', `Pengajuan telah ${status.toLowerCase()}.`, 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Terjadi kesalahan saat memproses verifikasi.', 'error');
      } finally {
        setIsSaving(false);
      }
    }
  };

  const filteredSubmissions = submissions.filter(s => {
    const searchStr = `${s.account?.full_name} ${s.type} ${s.description}`.toLowerCase();
    const matchesSearch = searchStr.includes(searchTerm.toLowerCase());
    
    if (activeTab === 'verification') return matchesSearch && s.status === 'Pending';
    if (activeTab === 'history') return matchesSearch && (s.status === 'Disetujui' || s.status === 'Ditolak');
    return matchesSearch; // Monitoring Tab
  });

  const TabButton = ({ id, label, icon: Icon, count }: { id: any, label: string, icon: any, count?: number }) => (
    <button
      onClick={() => setActiveTab(id)}
      className={`flex items-center gap-2 px-6 py-4 text-[11px] font-bold uppercase tracking-widest border-b-2 transition-all ${
        activeTab === id ? 'border-[#006E62] text-[#006E62] bg-emerald-50/30' : 'border-transparent text-gray-400 hover:text-gray-600'
      }`}
    >
      <Icon size={14} />
      {label}
      {count !== undefined && count > 0 && (
        <span className={`ml-1 px-1.5 py-0.5 rounded-full text-[9px] ${activeTab === id ? 'bg-[#006E62] text-white' : 'bg-gray-100 text-gray-400'}`}>
          {count}
        </span>
      )}
    </button>
  );

  return (
    <div className="space-y-6">
      {isSaving && <LoadingSpinner message="Memproses Verifikasi..." />}

      <div className="flex border-b border-gray-100 bg-white -mt-4 mb-6 sticky top-16 z-20 overflow-x-auto scrollbar-none">
        <TabButton 
          id="verification" 
          label="Verifikasi" 
          icon={Clock} 
          count={submissions.filter(s => s.status === 'Pending').length} 
        />
        <TabButton id="monitoring" label="Daftar Masuk" icon={ListFilter} />
        <TabButton id="history" label="Riwayat" icon={CheckCircle2} />
      </div>

      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="relative flex-1 max-w-md">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
          <input
            type="text"
            placeholder="Cari pengajuan (Nama, Jenis, Alasan)..."
            className="w-full pl-10 pr-4 py-2 bg-gray-50 border border-gray-200 rounded-md focus:outline-none focus:ring-2 focus:ring-[#006E62] text-sm"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </div>
        
        <button 
          onClick={() => setShowForm(true)}
          className="flex items-center gap-2 bg-[#006E62] text-white px-4 py-2 rounded-md hover:bg-[#005a50] transition-colors shadow-sm text-sm font-medium"
        >
          <Plus size={18} />
          <span>Buat Pengajuan</span>
        </button>
      </div>

      {isLoading ? (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {[1, 2, 3].map(i => <CardSkeleton key={i} />)}
        </div>
      ) : filteredSubmissions.length === 0 ? (
        <div className="flex flex-col items-center justify-center py-20 text-gray-400 bg-gray-50 rounded-xl border border-dashed border-gray-200">
          <ClipboardCheck size={48} strokeWidth={1} className="mb-4" />
          <p className="text-lg font-medium">Tidak ada data pengajuan.</p>
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {filteredSubmissions.map(submission => (
            <div 
              key={submission.id}
              onClick={() => setSelectedSubmission(submission)}
              className="bg-white border border-gray-100 p-5 rounded-xl shadow-sm hover:shadow-md transition-all cursor-pointer border-l-4 border-l-transparent hover:border-l-[#006E62] group relative"
            >
              <div className="flex justify-between items-start mb-3">
                <div className="flex flex-col">
                  <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-0.5">{submission.type}</span>
                  <h3 className="font-bold text-gray-800 text-sm group-hover:text-[#006E62] line-clamp-1">{submission.account?.full_name}</h3>
                </div>
                <span className={`text-[9px] font-bold px-2 py-0.5 rounded uppercase ${
                  submission.status === 'Pending' ? 'bg-orange-50 text-orange-600' :
                  submission.status === 'Disetujui' ? 'bg-emerald-50 text-emerald-600' : 'bg-red-50 text-red-600'
                }`}>
                  {submission.status}
                </span>
              </div>
              
              <p className="text-[11px] text-gray-500 line-clamp-2 mb-4 h-8">"{submission.description}"</p>
              
              <div className="flex items-center justify-between pt-4 border-t border-gray-50 text-[10px] font-bold text-gray-400 uppercase">
                 <div className="flex items-center gap-1.5">
                   <Clock size={12} /> {new Date(submission.created_at!).toLocaleDateString('id-ID')}
                 </div>
                 {submission.submission_data.duration_days && (
                   <span className="text-[#00FFE4]">{submission.submission_data.duration_days} Hari</span>
                 )}
              </div>
            </div>
          ))}
        </div>
      )}

      {showForm && (
        <SubmissionForm 
          onClose={() => setShowForm(false)} 
          onSuccess={() => { setShowForm(false); fetchSubmissions(); }} 
        />
      )}

      {selectedSubmission && (
        <SubmissionDetail 
          submission={selectedSubmission} 
          onClose={() => setSelectedSubmission(null)} 
          onVerify={handleVerify}
          canVerify={currentUser?.id !== selectedSubmission.account_id && selectedSubmission.status === 'Pending'}
        />
      )}
    </div>
  );
};

export default SubmissionMain;