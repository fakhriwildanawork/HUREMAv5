
import React, { useState, useEffect } from 'react';
import { Trophy, Star, Calendar, Users, Plus, Trash2, Award, Quote, ChevronRight, User, UserCircle, History } from 'lucide-react';
import { awardService } from '../../../services/awardService';
import { authService } from '../../../services/authService';
import { googleDriveService } from '../../../services/googleDriveService';
import { EmployeeOfThePeriod } from '../../../types';
import EmployeeOfThePeriodForm from './EmployeeOfThePeriodForm';
import Swal from 'sweetalert2';

const EmployeeOfThePeriodMain: React.FC = () => {
  const [awards, setAwards] = useState<EmployeeOfThePeriod[]>([]);
  const [loading, setLoading] = useState(true);
  const [isFormOpen, setIsFormOpen] = useState(false);
  
  const user = authService.getCurrentUser();
  const isAdmin = user?.role === 'admin' || user?.is_performance_admin;

  const fetchAwards = async () => {
    setLoading(true);
    try {
      const data = await awardService.getEmployeeOfThePeriodAll();
      setAwards(data);
    } catch (error) {
      console.error('Error fetching awards:', error);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchAwards();
  }, []);

  const handleDelete = async (id: string) => {
    const result = await Swal.fire({
      title: 'Hapus Data?',
      text: "Data penghargaan ini akan dihapus permanen.",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#ef4444',
      confirmButtonText: 'Ya, Hapus'
    });

    if (result.isConfirmed) {
      try {
        await awardService.deleteEmployeeOfThePeriod(id);
        Swal.fire('Berhasil', 'Data telah dihapus.', 'success');
        fetchAwards();
      } catch (error) {
        Swal.fire('Gagal', 'Terjadi kesalahan.', 'error');
      }
    }
  };

  const latestAward = awards[0];
  const historyAwards = awards.slice(1);

  if (loading) {
    return (
      <div className="flex flex-col items-center justify-center h-64">
        <div className="w-10 h-10 border-4 border-amber-500 border-t-transparent rounded-full animate-spin"></div>
        <p className="mt-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Memuat Hall of Fame...</p>
      </div>
    );
  }

  return (
    <div className="space-y-8 pb-12">
      {/* Header Section */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-2xl font-bold text-gray-800 flex items-center gap-2">
            <Trophy className="text-amber-500" size={28} />
            Employee of The Period
          </h2>
          <p className="text-sm text-gray-500 italic">Apresiasi bagi bintang-bintang HUREMA yang bersinar terang.</p>
        </div>
        {isAdmin && (
          <button
            onClick={() => setIsFormOpen(true)}
            className="flex items-center gap-2 px-4 py-2 bg-amber-500 text-white rounded-lg hover:bg-amber-600 transition-all text-sm font-bold shadow-md uppercase tracking-wider"
          >
            <Plus size={18} />
            Tambah Pemenang
          </button>
        )}
      </div>

      {/* Hero Section - Latest Winner */}
      {latestAward ? (
        <div className="relative overflow-hidden rounded-2xl bg-gradient-to-br from-amber-500 via-amber-600 to-orange-600 p-1 shadow-xl">
          <div className="bg-white rounded-[15px] overflow-hidden">
            <div className="grid md:grid-cols-12 gap-0">
              <div className="md:col-span-5 relative h-64 md:h-auto bg-gray-100">
                {latestAward.accounts && latestAward.accounts.length > 0 ? (
                  <div className="flex h-full">
                    {latestAward.accounts.map((acc, idx) => (
                      <div key={acc.id} className="flex-1 relative overflow-hidden group">
                        {acc.photo_google_id ? (
                          <img 
                            src={googleDriveService.getFileUrl(acc.photo_google_id)} 
                            alt={acc.full_name}
                            className="w-full h-full object-cover transition-transform duration-500 group-hover:scale-110"
                            referrerPolicy="no-referrer"
                          />
                        ) : (
                          <div className="w-full h-full flex items-center justify-center bg-gray-100 text-gray-400">
                            <UserCircle size={80} strokeWidth={1} />
                          </div>
                        )}
                        <div className="absolute inset-0 bg-gradient-to-t from-black/60 via-transparent to-transparent opacity-0 group-hover:opacity-100 transition-opacity flex flex-col justify-end p-4">
                          <p className="text-white text-xs font-bold uppercase tracking-widest">{acc.full_name}</p>
                          <p className="text-white/80 text-[10px] uppercase">{acc.position}</p>
                        </div>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="w-full h-full flex items-center justify-center bg-amber-50 text-amber-200">
                    <User size={80} strokeWidth={1} />
                  </div>
                )}
                <div className="absolute top-4 left-4 bg-amber-500 text-white px-3 py-1 rounded-full text-[10px] font-bold uppercase tracking-widest shadow-lg flex items-center gap-1.5">
                  <Star size={12} fill="white" />
                  Pemenang Utama
                </div>
              </div>
              
              <div className="md:col-span-7 p-8 md:p-12 flex flex-col justify-center relative">
                <div className="absolute top-0 right-0 p-8 opacity-5">
                  <Award size={160} />
                </div>
                
                <div className="space-y-6 relative z-10">
                  <div>
                    <div className="flex items-center gap-2 text-amber-600 mb-2">
                      <Calendar size={18} />
                      <span className="text-sm font-bold uppercase tracking-widest">
                        Periode {new Date(0, latestAward.month - 1).toLocaleString('id-ID', { month: 'long' })} {latestAward.year}
                      </span>
                    </div>
                    <h3 className="text-3xl font-black text-gray-900 leading-tight">
                      {latestAward.accounts?.map(a => a.full_name).join(' & ')}
                    </h3>
                  </div>

                  {latestAward.reason && (
                    <div className="relative">
                      <Quote className="absolute -left-6 -top-2 text-amber-200" size={32} />
                      <p className="text-gray-600 italic leading-relaxed text-lg pl-2">
                        {latestAward.reason}
                      </p>
                    </div>
                  )}

                  <div className="pt-4 flex items-center gap-4">
                    <div className="flex -space-x-3">
                      {latestAward.accounts?.map((acc) => (
                        <div key={acc.id} className="w-10 h-10 rounded-full border-2 border-white bg-gray-200 overflow-hidden shadow-sm">
                          {acc.photo_google_id ? (
                            <img src={googleDriveService.getFileUrl(acc.photo_google_id)} alt="" className="w-full h-full object-cover" referrerPolicy="no-referrer" />
                          ) : (
                            <div className="w-full h-full flex items-center justify-center text-gray-400"><User size={16} /></div>
                          )}
                        </div>
                      ))}
                    </div>
                    <div className="h-8 w-px bg-gray-100"></div>
                    <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Terpilih Berdasarkan Performa Terbaik</p>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      ) : (
        <div className="bg-gray-50 rounded-2xl border-2 border-dashed border-gray-200 p-12 text-center">
          <div className="w-20 h-20 bg-white rounded-full flex items-center justify-center mx-auto mb-4 shadow-sm">
            <Trophy size={40} className="text-gray-200" />
          </div>
          <p className="text-gray-500 font-medium">Belum ada data Employee of The Period yang diumumkan.</p>
        </div>
      )}

      {/* History Section */}
      {historyAwards.length > 0 && (
        <div className="space-y-6">
          <div className="flex items-center justify-between">
            <h3 className="text-lg font-bold text-gray-800 uppercase tracking-widest flex items-center gap-2">
              <History className="text-gray-400" size={20} />
              Riwayat Penghargaan
            </h3>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-6">
            {historyAwards.map((award) => (
              <div key={award.id} className="group bg-white rounded-xl border border-gray-100 overflow-hidden shadow-sm hover:shadow-md transition-all duration-300">
                <div className="aspect-video relative bg-gray-100 overflow-hidden">
                  <div className="flex h-full">
                    {award.accounts?.map((acc) => (
                      <div key={acc.id} className="flex-1 relative">
                        {acc.photo_google_id ? (
                          <img 
                            src={googleDriveService.getFileUrl(acc.photo_google_id)} 
                            alt="" 
                            className="w-full h-full object-cover group-hover:scale-110 transition-transform duration-500"
                            referrerPolicy="no-referrer"
                          />
                        ) : (
                          <div className="w-full h-full flex items-center justify-center bg-gray-50 text-gray-200"><User size={32} /></div>
                        )}
                      </div>
                    ))}
                  </div>
                  <div className="absolute inset-0 bg-gradient-to-t from-black/80 via-black/20 to-transparent"></div>
                  <div className="absolute bottom-4 left-4 right-4">
                    <p className="text-amber-400 text-[10px] font-bold uppercase tracking-widest mb-1">
                      {new Date(0, award.month - 1).toLocaleString('id-ID', { month: 'long' })} {award.year}
                    </p>
                    <h4 className="text-white font-bold truncate">
                      {award.accounts?.map(a => a.full_name).join(' & ')}
                    </h4>
                  </div>
                  {isAdmin && (
                    <button 
                      onClick={() => handleDelete(award.id)}
                      className="absolute top-2 right-2 p-2 bg-white/10 hover:bg-red-500 text-white rounded-lg backdrop-blur-md transition-all opacity-0 group-hover:opacity-100"
                    >
                      <Trash2 size={14} />
                    </button>
                  )}
                </div>
                <div className="p-4">
                  <p className="text-xs text-gray-500 line-clamp-2 italic">"{award.reason || 'Tanpa keterangan.'}"</p>
                  <div className="mt-4 pt-4 border-t border-gray-50 flex items-center justify-between">
                    <div className="flex -space-x-2">
                      {award.accounts?.map((acc) => (
                        <div key={acc.id} className="w-6 h-6 rounded-full border border-white bg-gray-200 overflow-hidden shadow-sm">
                          {acc.photo_google_id ? (
                            <img src={googleDriveService.getFileUrl(acc.photo_google_id)} alt="" className="w-full h-full object-cover" referrerPolicy="no-referrer" />
                          ) : (
                            <div className="w-full h-full flex items-center justify-center text-gray-400 text-[8px]"><User size={10} /></div>
                          )}
                        </div>
                      ))}
                    </div>
                    <span className="text-[10px] font-bold text-gray-300 uppercase tracking-tighter">Hall of Fame</span>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {isFormOpen && (
        <EmployeeOfThePeriodForm 
          onClose={() => setIsFormOpen(false)}
          onSuccess={() => {
            setIsFormOpen(false);
            fetchAwards();
          }}
        />
      )}
    </div>
  );
};

export default EmployeeOfThePeriodMain;
