import React, { useState, useEffect, useRef } from 'react';
import { Timer, Clock, MapPin, History, AlertCircle, Map as MapIcon, Camera, UserX, ShieldCheck, Info } from 'lucide-react';
import Swal from 'sweetalert2';
import { overtimeService } from '../../services/overtimeService';
import { presenceService } from '../../services/presenceService';
import { accountService } from '../../services/accountService';
import { authService } from '../../services/authService';
import { googleDriveService } from '../../services/googleDriveService';
import { settingsService } from '../../services/settingsService';
import { submissionService } from '../../services/submissionService';
import { Account, Overtime } from '../../types';
import PresenceCamera from '../presence/PresenceCamera';
import PresenceMap from '../presence/PresenceMap';
import OvertimeHistory from './OvertimeHistory';
import LoadingSpinner from '../../components/Common/LoadingSpinner';

const OvertimeMain: React.FC = () => {
  const [activeTab, setActiveTab] = useState<'capture' | 'history'>('capture');
  const [account, setAccount] = useState<Account | null>(null);
  const [todayOT, setTodayOT] = useState<Overtime | null>(null);
  const [recentLogs, setRecentLogs] = useState<Overtime[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isCapturing, setIsCapturing] = useState(false);
  const [isCameraActive, setIsCameraActive] = useState(false);
  const [isRegularActive, setIsRegularActive] = useState(false);
  const [serverTime, setServerTime] = useState<Date>(new Date());
  const [coords, setCoords] = useState<{lat: number, lng: number} | null>(null);
  const [distance, setDistance] = useState<number | null>(null);
  const watchId = useRef<number | null>(null);

  const currentUser = authService.getCurrentUser();
  const currentAccountId = currentUser?.id;

  useEffect(() => {
    if (!currentAccountId) return;
    
    fetchInitialData();
    const timeInterval = setInterval(() => {
      setServerTime(prev => new Date(prev.getTime() + 1000));
    }, 1000);
    
    startWatchingLocation();

    return () => {
      clearInterval(timeInterval);
      if (watchId.current !== null) navigator.geolocation.clearWatch(watchId.current);
    };
  }, [currentAccountId]);

  const fetchInitialData = async () => {
    if (!currentAccountId) return;
    try {
      setIsLoading(true);
      const [acc, ot, history, isRegActive, sTime] = await Promise.all([
        accountService.getById(currentAccountId),
        overtimeService.getTodayOvertime(currentAccountId),
        overtimeService.getRecentHistory(currentAccountId),
        presenceService.isRegularSessionActive(currentAccountId),
        presenceService.getServerTime()
      ]);
      setAccount(acc as any);
      setTodayOT(ot);
      setRecentLogs(history);
      setIsRegularActive(isRegActive);
      setServerTime(sTime);
    } catch (error) {
      console.error(error);
    } finally {
      setIsLoading(false);
    }
  };

  const startWatchingLocation = () => {
    if (navigator.geolocation) {
      watchId.current = navigator.geolocation.watchPosition(
        (pos) => {
          if (pos.coords.accuracy < 1000) {
            setCoords({ lat: pos.coords.latitude, lng: pos.coords.longitude });
          }
        },
        (err) => console.error("GPS Error:", err),
        { enableHighAccuracy: true, maximumAge: 0, timeout: 20000 }
      );
    }
  };

  useEffect(() => {
    if (coords && account?.location) {
      const d = presenceService.calculateDistance(
        coords.lat, coords.lng,
        account.location.latitude, account.location.longitude
      );
      setDistance(d);
    }
  }, [coords, account]);

  const handleOvertime = async (photoBlob: Blob) => {
    if (isRegularActive) {
      return Swal.fire('Akses Ditolak', 'Anda tidak dapat memulai lembur saat sesi kerja reguler sedang aktif. Harap Check-Out reguler terlebih dahulu.', 'warning');
    }

    if (!account) return;

    if (!coords || distance === null) {
      return Swal.fire('Lokasi Belum Siap', 'Sinyal GPS sedang dioptimalkan...', 'warning');
    }

    // Geofencing Check
    const isCheckOut = !!todayOT && !todayOT.check_out;
    // FIX: Ensure strict boolean check
    const isLimited = isCheckOut 
      ? account.is_presence_limited_ot_out === true 
      : account.is_presence_limited_ot_in === true;
    const locationRadius = account.location?.radius || 100;

    if (isLimited && distance > locationRadius) {
       setIsCameraActive(false);
       return Swal.fire('Diluar Radius', `Anda berada diluar radius penempatan untuk lembur.`, 'error');
    }

    // Alasan Lembur Mandatory Prompt
    const { value: otReason, isConfirmed } = await Swal.fire({
      title: 'Konfirmasi Lembur',
      input: 'textarea',
      inputLabel: `Sebutkan alasan/kegiatan lembur (${isCheckOut ? 'Check-Out' : 'Check-In'}):`,
      inputPlaceholder: 'Contoh: Menyelesaikan laporan bulanan...',
      showCancelButton: true,
      confirmButtonColor: '#d97706',
      cancelButtonColor: '#6b7280',
      confirmButtonText: 'Kirim Presensi',
      cancelButtonText: 'Batal',
      inputValidator: (value) => {
        if (!value) return 'Alasan lembur wajib diisi!';
        return null;
      }
    });

    if (!isConfirmed) return;

    try {
      setIsCapturing(true);
      
      // Capture current state to avoid race conditions
      const currentOT = todayOT;
      const isCurrentlyCheckingOut = !!currentOT && !currentOT.check_out;

      const [address, photoId, otPolicy] = await Promise.all([
        presenceService.getReverseGeocode(coords.lat, coords.lng),
        googleDriveService.uploadFile(new File([photoBlob], `OT_${isCurrentlyCheckingOut ? 'OUT' : 'IN'}_${Date.now()}.jpg`)),
        settingsService.getSetting('ot_approval_policy', 'manual')
      ]);

      const currentTimeStr = serverTime.toISOString();
      
      if (!isCurrentlyCheckingOut) {
        await overtimeService.checkIn({
          account_id: account.id,
          check_in: currentTimeStr,
          in_latitude: coords.lat,
          in_longitude: coords.lng,
          in_photo_id: photoId,
          in_address: address,
          reason: otReason
        });
      } else {
        if (!currentOT?.id) {
          throw new Error("ID referensi lembur tidak ditemukan. Harap muat ulang halaman.");
        }

        const start = new Date(currentOT.check_in!);
        const diffMs = serverTime.getTime() - start.getTime();
        const diffMins = Math.max(0, Math.floor(diffMs / 60000));
        
        const h = Math.floor(diffMs / 3600000);
        const m = Math.floor((diffMs % 3600000) / 60000);
        const s = Math.floor((diffMs % 60000) / 1000);
        const durationFormatted = `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;

        await overtimeService.checkOut(currentOT.id, {
          check_out: currentTimeStr,
          out_latitude: coords.lat,
          out_longitude: coords.lng,
          out_photo_id: photoId,
          out_address: address,
          duration_minutes: diffMins,
          work_duration: durationFormatted,
          reason: otReason
        });

        // WORKFLOW INTEGRATION: Masuk ke Pengajuan jika policy MANUAL
        if (otPolicy === 'manual') {
          await submissionService.create({
            account_id: account.id,
            type: 'Lembur',
            description: `Lembur pada ${new Date(currentOT.check_in!).toLocaleDateString('id-ID')}. Kegiatan: ${otReason}`,
            file_id: photoId,
            submission_data: {
              overtime_id: currentOT.id,
              date: currentOT.check_in!.split('T')[0],
              duration: durationFormatted,
              minutes: diffMins,
              check_in: currentOT.check_in,
              check_out: currentTimeStr
            }
          });
        }
      }

      setIsCameraActive(false); 
      
      let successMsg = `Presensi Lembur dicatat.`;
      if (isCurrentlyCheckingOut && otPolicy === 'manual') {
        successMsg = `Presensi Lembur tersimpan. Pengajuan verifikasi telah dikirim ke Admin.`;
      } else if (isCurrentlyCheckingOut && otPolicy === 'auto') {
        successMsg = `Presensi Lembur tersimpan dan disetujui otomatis oleh sistem.`;
      }

      Swal.fire({ title: 'Berhasil!', text: successMsg, icon: 'success', timer: 3000, showConfirmButton: false });
      await fetchInitialData();
    } catch (error: any) {
      console.error("Overtime Process Error:", error);
      Swal.fire({
        title: 'Gagal',
        text: error.message || 'Terjadi kesalahan sistem saat memproses lembur.',
        icon: 'error',
        confirmButtonColor: '#d97706'
      });
    } finally {
      setIsCapturing(false);
    }
  };

  const getOTDurationLive = () => {
    if (!todayOT?.check_in || todayOT.check_out) return null;
    const start = new Date(todayOT.check_in);
    const diffMs = serverTime.getTime() - start.getTime();
    if (diffMs < 0) return "00 : 00 : 00";
    const hours = Math.floor(diffMs / 3600000);
    const minutes = Math.floor((diffMs % 3600000) / 60000);
    const seconds = Math.floor((diffMs % 60000) / 1000);
    return `${String(hours).padStart(2, '0')} : ${String(minutes).padStart(2, '0')} : ${String(seconds).padStart(2, '0')}`;
  };

  if (isLoading) return <LoadingSpinner message="Menghubungkan Modul Lembur..." />;
  if (!account) return <div className="p-10 text-center">Akun tidak valid.</div>;

  const isCheckOut = !!todayOT && !todayOT.check_out;
  const isLimited = isCheckOut 
    ? account.is_presence_limited_ot_out === true 
    : account.is_presence_limited_ot_in === true;
  const isWithinRadius = distance !== null && distance <= (account?.location?.radius || 100);
  const isBlockedByLocation = isLimited && !isWithinRadius;

  return (
    <div className="max-w-5xl mx-auto space-y-6">
      <div className="bg-white rounded-2xl border border-gray-100 p-6 shadow-sm flex flex-col md:flex-row justify-between items-center gap-6">
        <div className="flex items-center gap-4">
          <div className="w-12 h-12 bg-amber-50 rounded-xl flex items-center justify-center text-amber-600">
            <Timer size={28} />
          </div>
          <div>
            <h2 className="text-xl font-bold text-gray-800 tracking-tight">Presensi Lembur</h2>
            <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">{account?.full_name} • Modul OT</p>
          </div>
        </div>
        <div className="flex gap-2 p-1 bg-gray-50 rounded-xl border border-gray-100">
          <button 
            onClick={() => setActiveTab('capture')}
            className={`flex items-center gap-2 px-6 py-2.5 rounded-lg text-xs font-bold uppercase transition-all ${activeTab === 'capture' ? 'bg-white text-amber-600 shadow-sm' : 'text-gray-400 hover:text-gray-600'}`}
          >
            <Camera size={16} /> Capturing
          </button>
          <button 
            onClick={() => setActiveTab('history')}
            className={`flex items-center gap-2 px-6 py-2.5 rounded-lg text-xs font-bold uppercase transition-all ${activeTab === 'history' ? 'bg-white text-amber-600 shadow-sm' : 'text-gray-400 hover:text-gray-600'}`}
          >
            <History size={16} /> Log OT
          </button>
        </div>
      </div>

      {activeTab === 'capture' ? (
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start animate-in fade-in duration-500">
          <div className="lg:col-span-7">
            {isCameraActive ? (
              <PresenceCamera 
                onCapture={handleOvertime}
                onClose={() => setIsCameraActive(false)}
                isProcessing={isCapturing}
              />
            ) : isRegularActive ? (
              <div className="bg-white rounded-2xl border border-gray-100 p-12 flex flex-col items-center justify-center shadow-sm text-center">
                <div className="w-24 h-24 rounded-full bg-rose-50 text-rose-500 flex items-center justify-center mb-8 shadow-xl">
                   <AlertCircle size={48} />
                </div>
                <h3 className="text-2xl font-bold text-gray-800">Sesi Reguler Aktif</h3>
                <p className="text-sm text-gray-400 mt-4 max-w-xs leading-relaxed">Sistem mendeteksi Anda masih berada dalam jam kerja reguler (Belum Check-Out). Lembur hanya bisa dimulai di luar jam sesi kerja reguler.</p>
              </div>
            ) : (!todayOT || !todayOT.check_out) ? (
              <div className="bg-white rounded-2xl border border-gray-100 p-12 flex flex-col items-center justify-center shadow-sm text-center">
                <div className={`w-24 h-24 rounded-full flex items-center justify-center mb-8 shadow-xl transition-all duration-500 ${!isBlockedByLocation ? 'bg-amber-50 text-amber-600' : 'bg-rose-50 text-rose-500'}`}>
                   <Timer size={48} />
                </div>
                <h3 className="text-2xl font-bold text-gray-800">
                   {!!todayOT && !todayOT.check_out ? 'Selesaikan Lembur?' : 'Mulai Lembur Sekarang?'}
                </h3>
                <p className="text-sm text-gray-400 mt-2 max-w-xs">
                  {!isBlockedByLocation 
                    ? 'Verifikasi identitas diperlukan untuk mencatat waktu tambahan kerja Anda hari ini.'
                    : 'Akses presensi lembur terkunci. Anda harus berada di area lokasi penempatan.'}
                </p>
                <button 
                  disabled={isBlockedByLocation || isCapturing}
                  onClick={() => setIsCameraActive(true)}
                  className={`mt-10 flex items-center gap-3 px-12 py-4 rounded-2xl font-bold uppercase text-xs tracking-widest shadow-lg transition-all ${
                    !isBlockedByLocation && !isCapturing 
                    ? 'bg-amber-600 text-white hover:bg-amber-700 hover:scale-105' 
                    : 'bg-gray-100 text-gray-300 cursor-not-allowed shadow-none'
                  }`}
                >
                  <Camera size={18} />
                  {isCapturing ? 'MEMPROSES...' : 'MULAI VERIFIKASI WAJAH'}
                </button>
              </div>
            ) : (
              <div className="bg-white rounded-2xl border border-gray-100 p-20 flex flex-col items-center justify-center shadow-sm text-center">
                <div className="w-20 h-20 bg-emerald-100 rounded-full flex items-center justify-center text-[#006E62] mb-6">
                  <ShieldCheck size={48} />
                </div>
                <h3 className="text-2xl font-bold text-gray-800">Lembur Selesai!</h3>
                <p className="text-sm text-gray-500 mt-2 max-w-xs">Data lembur hari ini telah tersimpan dengan aman.</p>
              </div>
            )}
          </div>

          <div className="lg:col-span-5 space-y-6">
            <div className="bg-white rounded-2xl border border-gray-100 p-6 shadow-sm">
               <div className="flex items-center gap-2 mb-6">
                  <Clock size={16} className="text-amber-600" />
                  <h4 className="text-[11px] font-bold uppercase tracking-widest text-gray-500">Waktu Terverifikasi</h4>
               </div>
               <div className="text-center py-4">
                  <div className="text-5xl font-mono font-bold text-gray-800 tracking-tighter">
                    {serverTime.toLocaleTimeString('id-ID', { hour12: false })}
                  </div>
                  <div className="text-[11px] font-bold text-amber-500 uppercase tracking-widest mt-2">
                    {serverTime.toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long' })}
                  </div>

                  {todayOT?.check_in && !todayOT.check_out && (
                    <div className="mt-8 p-3 bg-amber-50 rounded-2xl border border-amber-100 animate-pulse">
                      <p className="text-[9px] font-black text-amber-600 uppercase tracking-[0.2em] mb-1">Durasi Lembur Berjalan</p>
                      <div className="text-2xl font-mono font-black text-amber-700 tracking-widest">
                        {getOTDurationLive()}
                      </div>
                    </div>
                  )}
               </div>

               <div className="mt-8 p-4 bg-gray-50 rounded-xl border border-gray-100 flex gap-3 items-start">
                  <Info size={16} className="text-blue-500 shrink-0 mt-0.5" />
                  <p className="text-[10px] text-gray-500 leading-relaxed font-medium italic">Sistem lembur tidak mengacu pada jadwal kerja reguler. Durasi dihitung murni dari waktu kehadiran awal hingga selesai.</p>
               </div>
            </div>

            <div className="bg-white rounded-2xl border border-gray-100 p-6 shadow-sm">
               <div className="flex items-center justify-between mb-4">
                  <div className="flex items-center gap-2">
                    <MapPin size={16} className="text-amber-500" />
                    <h4 className="text-[11px] font-bold uppercase tracking-widest text-gray-500">Status Geotag</h4>
                  </div>
                  <div className={`flex items-center gap-1.5 px-2 py-1 rounded-full text-[9px] font-bold uppercase ${!isLimited ? 'bg-blue-50 text-blue-600' : (isWithinRadius ? 'bg-emerald-50 text-emerald-600' : 'bg-rose-50 text-rose-600')}`}>
                    <div className={`w-1.5 h-1.5 rounded-full ${!isLimited ? 'bg-blue-500' : (isWithinRadius ? 'bg-emerald-500' : 'bg-rose-500')}`}></div>
                    {!isLimited ? 'Bebas Lokasi' : (isWithinRadius ? 'Area Kerja' : 'Diluar Area')}
                  </div>
               </div>
               
               {account?.location && coords ? (
                 <div className="space-y-4">
                    <PresenceMap 
                      userLat={coords.lat} 
                      userLng={coords.lng} 
                      officeLat={account.location.latitude} 
                      officeLng={account.location.longitude}
                      radius={account.location.radius}
                    />
                    <div className="grid grid-cols-2 gap-4 text-[10px] pt-2">
                       <div className="p-3 bg-gray-50 rounded-xl border border-gray-100">
                          <span className="text-gray-400 font-bold uppercase block mb-1">Jarak</span>
                          <span className="text-sm font-bold text-gray-700">{distance !== null ? `${Math.round(distance)}m` : '...'}</span>
                       </div>
                       <div className="p-3 bg-gray-50 rounded-xl border border-gray-100">
                          <span className="text-gray-400 font-bold uppercase block mb-1">Status OT</span>
                          <span className="text-sm font-bold text-amber-600">Terbatas</span>
                       </div>
                    </div>
                    {isBlockedByLocation && (
                      <div className="p-3 bg-rose-50 border border-rose-100 rounded-xl flex gap-3 items-start animate-pulse">
                         <AlertCircle size={16} className="text-rose-500 shrink-0 mt-0.5" />
                         <p className="text-[10px] text-rose-600 font-bold leading-tight uppercase tracking-tight">Presensi Lembur dikunci. Anda berada diluar zona yang diizinkan.</p>
                      </div>
                    )}
                 </div>
               ) : (
                 <div className="flex flex-col items-center justify-center py-14 text-gray-300">
                    <MapIcon size={40} className="animate-bounce" />
                    <p className="text-[10px] font-bold uppercase tracking-widest mt-4">Mengunci Sinyal GPS...</p>
                 </div>
               )}
            </div>
          </div>
        </div>
      ) : (
        <div className="animate-in fade-in slide-in-from-bottom-4 duration-500">
          <OvertimeHistory logs={recentLogs} isLoading={isLoading} />
        </div>
      )}
    </div>
  );
};

export default OvertimeMain;