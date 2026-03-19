import React, { useState, useEffect } from 'react';
import { MessageSquare, Plus, Search, Calendar, User, Shield, CheckCircle2, Clock, Eye, Trash2, Filter, AlertCircle } from 'lucide-react';
import Swal from 'sweetalert2';
import { feedbackService } from '../../services/feedbackService';
import { authService } from '../../services/authService';
import { Feedback, AuthUser } from '../../types';
import FeedbackForm from './FeedbackForm';
import FeedbackDetail from './FeedbackDetail';
import LoadingSpinner from '../../components/Common/LoadingSpinner';
import { CardSkeleton } from '../../components/Common/Skeleton';

const FeedbackMain: React.FC = () => {
  const [feedbacks, setFeedbacks] = useState<Feedback[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [categoryFilter, setCategoryFilter] = useState('All');
  const [showForm, setShowForm] = useState(false);
  const [selectedFeedback, setSelectedFeedback] = useState<Feedback | null>(null);
  const [showDetail, setShowDetail] = useState(false);
  const [user, setUser] = useState<AuthUser | null>(null);

  const isAdmin = user?.role === 'admin' || user?.is_hr_admin || user?.is_performance_admin || user?.is_finance_admin;

  useEffect(() => {
    const currentUser = authService.getCurrentUser();
    setUser(currentUser);
    fetchData(currentUser);
  }, []);

  const fetchData = async (currentUser: AuthUser | null) => {
    if (!currentUser) return;
    const isUserAdmin = currentUser.role === 'admin' || currentUser.is_hr_admin || currentUser.is_performance_admin || currentUser.is_finance_admin;
    try {
      setIsLoading(true);
      const data = isUserAdmin 
        ? await feedbackService.getAll() 
        : await feedbackService.getByAccountId(currentUser.id);
      setFeedbacks(data);
    } catch (error) {
      console.error(error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleSubmitFeedback = async (input: any) => {
    if (!user) return;
    try {
      setIsSaving(true);
      await feedbackService.submitFeedback({
        ...input,
        account_id: user.id
      });
      await fetchData(user);
      setShowForm(false);
      Swal.fire({
        title: 'Berhasil!',
        text: 'Feedback Anda telah dikirim ke manajemen.',
        icon: 'success',
        timer: 2000,
        showConfirmButton: false
      });
    } catch (error) {
      Swal.fire('Gagal', 'Terjadi kesalahan saat mengirim feedback.', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleDelete = async (id: string) => {
    const result = await Swal.fire({
      title: 'Hapus Feedback?',
      text: "Data feedback akan dihapus permanen.",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
      confirmButtonText: 'Ya, Hapus!'
    });

    if (result.isConfirmed) {
      try {
        setIsSaving(true);
        await feedbackService.delete(id);
        await fetchData(user);
        Swal.fire('Terhapus!', 'Feedback telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Terjadi kesalahan saat menghapus feedback.', 'error');
      } finally {
        setIsSaving(false);
      }
    }
  };

  const handleViewDetail = async (feedback: Feedback) => {
    setSelectedFeedback(feedback);
    setShowDetail(true);
    
    // Auto mark as read for admin
    if (isAdmin && feedback.status === 'Unread') {
      try {
        await feedbackService.markAsRead(feedback.id);
        // Update local state
        setFeedbacks(prev => prev.map(f => f.id === feedback.id ? { ...f, status: 'Read' } : f));
      } catch (error) {
        console.error('Failed to mark as read:', error);
      }
    }
  };

  const filteredFeedbacks = feedbacks.filter(f => {
    const matchesSearch = f.description.toLowerCase().includes(searchTerm.toLowerCase()) ||
                         f.category.toLowerCase().includes(searchTerm.toLowerCase()) ||
                         (!f.is_anonymous && f.account?.full_name?.toLowerCase().includes(searchTerm.toLowerCase()));
    const matchesCategory = categoryFilter === 'All' || f.category === categoryFilter;
    return matchesSearch && matchesCategory;
  });

  const categories = ['All', 'Gaji', 'Fasilitas', 'Hubungan Kerja', 'Lainnya'];

  return (
    <div className="space-y-6">
      {isSaving && <LoadingSpinner message="Memproses Data..." />}

      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-xl font-bold text-gray-800 tracking-tight">Feedback Pegawai</h2>
          <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Suara Pegawai untuk Perusahaan Lebih Baik</p>
        </div>
        
        <div className="flex flex-wrap items-center gap-3">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
            <input
              type="text"
              placeholder="Cari feedback..."
              className="pl-9 pr-4 py-2 bg-gray-50 border border-gray-100 rounded-md focus:outline-none focus:ring-2 focus:ring-[#006E62] text-xs w-48 md:w-64"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
          
          <select 
            value={categoryFilter}
            onChange={(e) => setCategoryFilter(e.target.value)}
            className="px-3 py-2 bg-gray-50 border border-gray-100 rounded-md focus:outline-none focus:ring-2 focus:ring-[#006E62] text-xs font-bold text-gray-600"
          >
            {categories.map(cat => <option key={cat} value={cat}>{cat}</option>)}
          </select>

          {!isAdmin && (
            <button 
              onClick={() => setShowForm(true)}
              className="flex items-center gap-2 bg-[#006E62] text-white px-4 py-2 rounded-md hover:bg-[#005a50] transition-colors shadow-sm text-xs font-bold uppercase"
            >
              <Plus size={16} />
              Kirim Feedback
            </button>
          )}
        </div>
      </div>

      {isLoading ? (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {[1, 2, 3].map(i => <CardSkeleton key={i} />)}
        </div>
      ) : (
        <div className="space-y-4">
          {isAdmin ? (
            <div className="bg-white rounded-2xl border border-gray-100 shadow-sm overflow-hidden">
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-gray-50/50 border-b border-gray-100">
                      <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Status</th>
                      <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Kategori</th>
                      <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Pengirim</th>
                      <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Prioritas</th>
                      <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tanggal</th>
                      <th className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest text-right">Aksi</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-50">
                    {filteredFeedbacks.map(feedback => (
                      <tr key={feedback.id} className={`hover:bg-gray-50/50 transition-colors ${feedback.status === 'Unread' ? 'bg-emerald-50/20' : ''}`}>
                        <td className="px-6 py-4">
                          {feedback.status === 'Unread' ? (
                            <span className="flex items-center gap-1.5 text-[10px] font-bold text-emerald-600 uppercase">
                              <div className="w-1.5 h-1.5 bg-emerald-600 rounded-full animate-pulse"></div>
                              Baru
                            </span>
                          ) : (
                            <span className="flex items-center gap-1.5 text-[10px] font-bold text-gray-400 uppercase">
                              <CheckCircle2 size={12} />
                              Dibaca
                            </span>
                          )}
                        </td>
                        <td className="px-6 py-4">
                          <span className="px-2 py-1 bg-gray-100 text-gray-600 rounded text-[9px] font-bold uppercase">
                            {feedback.category}
                          </span>
                        </td>
                        <td className="px-6 py-4">
                          {feedback.is_anonymous ? (
                            <span className="flex items-center gap-1.5 text-xs font-bold text-gray-400 italic">
                              <Shield size={12} />
                              Anonim
                            </span>
                          ) : (
                            <div className="flex items-center gap-2">
                              <div className="w-6 h-6 rounded-full bg-gray-100 flex items-center justify-center text-[10px] font-bold text-gray-500">
                                {feedback.account?.full_name?.charAt(0)}
                              </div>
                              <span className="text-xs font-bold text-gray-700">{feedback.account?.full_name}</span>
                            </div>
                          )}
                        </td>
                        <td className="px-6 py-4">
                          <span className={`text-[10px] font-bold uppercase ${
                            feedback.priority === 'High' ? 'text-rose-500' : 
                            feedback.priority === 'Medium' ? 'text-amber-500' : 'text-emerald-500'
                          }`}>
                            {feedback.priority}
                          </span>
                        </td>
                        <td className="px-6 py-4 text-[10px] font-bold text-gray-400 uppercase">
                          {new Date(feedback.created_at).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' })}
                        </td>
                        <td className="px-6 py-4 text-right">
                          <div className="flex items-center justify-end gap-2">
                            <button 
                              onClick={() => handleViewDetail(feedback)}
                              className="p-2 text-gray-400 hover:text-[#006E62] hover:bg-emerald-50 rounded-lg transition-all"
                              title="Lihat Detail"
                            >
                              <Eye size={16} />
                            </button>
                            <button 
                              onClick={() => handleDelete(feedback.id)}
                              className="p-2 text-gray-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-all"
                              title="Hapus"
                            >
                              <Trash2 size={16} />
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          ) : (
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
              {filteredFeedbacks.map(feedback => (
                <div key={feedback.id} className="bg-white border border-gray-100 p-5 rounded-2xl shadow-sm hover:shadow-md transition-all border-l-4 border-l-transparent hover:border-l-[#006E62] group flex flex-col h-full">
                  <div className="flex justify-between items-start mb-4">
                    <span className={`px-2 py-1 rounded text-[9px] font-bold uppercase ${
                      feedback.priority === 'High' ? 'bg-rose-50 text-rose-600' : 
                      feedback.priority === 'Medium' ? 'bg-amber-50 text-amber-600' : 'bg-emerald-50 text-emerald-600'
                    }`}>
                      {feedback.priority} Priority
                    </span>
                    <div className="flex items-center gap-1">
                      <button 
                        onClick={() => handleViewDetail(feedback)}
                        className="p-1.5 text-gray-400 hover:text-[#006E62] hover:bg-emerald-50 rounded-lg transition-all"
                      >
                        <Eye size={14} />
                      </button>
                    </div>
                  </div>
                  
                  <div className="mb-3">
                    <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">{feedback.category}</p>
                    <p className="text-xs text-gray-600 line-clamp-3 italic">"{feedback.description}"</p>
                  </div>

                  <div className="mt-auto pt-4 border-t border-gray-50 flex items-center justify-between">
                    <div className="flex items-center gap-2 text-[10px] font-bold text-gray-400 uppercase">
                      <Calendar size={12} />
                      <span>{new Date(feedback.created_at).toLocaleDateString('id-ID', { day: 'numeric', month: 'short' })}</span>
                    </div>
                    <div className="flex items-center gap-1.5">
                      {feedback.status === 'Read' ? (
                        <span className="text-[9px] font-bold text-emerald-600 uppercase flex items-center gap-1">
                          <CheckCircle2 size={10} />
                          Ditinjau
                        </span>
                      ) : (
                        <span className="text-[9px] font-bold text-amber-500 uppercase flex items-center gap-1">
                          <Clock size={10} />
                          Pending
                        </span>
                      )}
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}

          {filteredFeedbacks.length === 0 && (
            <div className="flex flex-col items-center justify-center py-20 text-gray-400 bg-gray-50 rounded-xl border border-dashed border-gray-200">
              <MessageSquare size={48} strokeWidth={1} className="mb-4 opacity-20" />
              <p className="text-sm font-bold uppercase tracking-widest">Belum ada feedback ditemukan</p>
            </div>
          )}
        </div>
      )}

      {showForm && (
        <FeedbackForm 
          onClose={() => setShowForm(false)} 
          onSubmit={handleSubmitFeedback}
        />
      )}

      {showDetail && selectedFeedback && (
        <FeedbackDetail 
          feedback={selectedFeedback}
          onClose={() => { setShowDetail(false); setSelectedFeedback(null); }}
          isAdmin={isAdmin}
        />
      )}
    </div>
  );
};

export default FeedbackMain;
