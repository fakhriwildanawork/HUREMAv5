
import React, { useState, useEffect, useRef } from 'react';
/* Added ShieldCheck to lucide-react imports */
import { Fingerprint, Clock, MapPin, History, AlertCircle, Map as MapIcon, Camera, Search, UserX, CalendarClock, MessageSquare, ShieldCheck, Umbrella, RefreshCw, Check } from 'lucide-react';
import Swal from 'sweetalert2';
import { presenceService } from '../../services/presenceService';
import { accountService } from '../../services/accountService';
import { scheduleService } from '../../services/scheduleService';
import { authService } from '../../services/authService';
import { googleDriveService } from '../../services/googleDriveService';
import { Account, Attendance, Schedule, ScheduleRule } from '../../types';
import PresenceCamera from './PresenceCamera';
import PresenceMap from './PresenceMap';
import PresenceHistory from './PresenceHistory';
import LoadingSpinner from '../../components/Common/LoadingSpinner';

const PresenceMain: React.FC = () => {
  const [activeTab, setActiveTab] = useState<'capture' | 'history'>('capture');
  const [account, setAccount] = useState<Account | null>(null);
  const [todayAttendance, setTodayAttendance] = useState<Attendance | null>(null);
  const [recentLogs, setRecentLogs] = useState<Attendance[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isCapturing, setIsCapturing] = useState(false);
  const [isCameraActive, setIsCameraActive] = useState(false);
  const [serverTime, setServerTime] = useState<Date>(new Date());
  const [coords, setCoords] = useState<{lat: number, lng: number} | null>(null);
  const [gpsAccuracy, setGpsAccuracy] = useState<number | null>(null);
  const [distance, setDistance] = useState<number | null>(null);
  const [activeHoliday, setActiveHoliday] = useState<any>(null);
  
  // State khusus Shift Dinamis
  const [dynamicShifts, setDynamicShifts] = useState<Schedule[]>([]);
  const [selectedShift, setSelectedShift] = useState<Schedule | null>(null);
  const [isFetchingShifts, setIsFetchingShifts] = useState(false);

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
      const [acc, attendance, history, sTime] = await Promise.all([
        accountService.getById(currentAccountId),
        presenceService.getTodayAttendance(currentAccountId),
        presenceService.getRecentHistory(currentAccountId),
        presenceService.getServerTime()
      ]);
      setAccount(acc as any);
      setTodayAttendance(attendance);
      setRecentLogs(history);
      setServerTime(sTime);

      // Tarik daftar shift jika akun bertipe DINAMIS (Cek via schedule_type)
      if (acc && acc.schedule_type === 'Shift Dinamis' && acc.location_id) {
        setIsFetchingShifts(true);
        const shifts = await scheduleService.getByLocation(acc.location_id);
        // MANDATORY: Hanya bertipe Shift Kerja (Type 2)
        setDynamicShifts(shifts.filter((s: any) => s.type === 2));
        setIsFetchingShifts(false);
      }

      // Cek apakah hari ini Libur Khusus (Type 3) di lokasi user
      if (acc && acc.location_id) {
        const holiday = await presenceService.checkHolidayStatus(currentAccountId, acc.location_id, sTime);
        setActiveHoliday(holiday);
      }
    } catch (error) {
      console.error(error);
    } finally {
      setIsLoading(false);
    }
  };

  const getLiveWorkDuration = () => {
    if (!todayAttendance?.check_in || todayAttendance.check_out) return null;
    const start = new Date(todayAttendance.check_in);
    const diffMs = serverTime.getTime() - start.getTime();
    if (diffMs < 0) return "00 : 00 : 00";
    const hours = Math.floor(diffMs / 3600000);
    const minutes = Math.floor((diffMs % 3600000) / 60000);
    const seconds = Math.floor((diffMs % 60000) / 1000);
    return `${String(hours).padStart(2, '0')} : ${String(minutes).padStart(2, '0')} : ${String(seconds).padStart(2, '0')}`;
  };

  const startWatchingLocation = () => {
    if (navigator.geolocation) {
      watchId.current = navigator.geolocation.watchPosition(
        (pos) => {
          setGpsAccuracy(pos.coords.accuracy);
          // Mengizinkan koordinat masuk jika akurasi masih dalam batas toleransi logis
          if (pos.coords.accuracy < 1000) {
            setCoords({ lat: pos.coords.latitude, lng: pos.coords.longitude });
          }
        },
        (err) => {
          console.error("GPS Error:", err);
        },
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

  const handleAttendance = async (photoBlob: Blob) => {
    const todayDay = serverTime.getDay();
    
    // Resolve Target Schedule (Statis vs Dinamis vs Fleksibel)
    let targetSchedule: Schedule | null = null;
    if (account?.schedule_type === 'Fleksibel') {
      targetSchedule = { id: 'FLEKSIBEL', name: 'Fleksibel' } as any;
    } else if (account?.schedule_type === 'Shift Dinamis') {
      if (!selectedShift) return Swal.fire('Peringatan', 'Pilih shift terlebih dahulu.', 'warning');
      targetSchedule = selectedShift;
    } else {
      targetSchedule = account?.schedule || null;
    }

    const scheduleRule = targetSchedule?.rules?.find(r => r.day_of_week === todayDay);
    const isHolidayToday = !!activeHoliday || !!scheduleRule?.is_holiday;

    // Blokir jika hari libur (kecuali Fleksibel)
    if (isHolidayToday && account?.schedule_type !== 'Fleksibel') {
      const holidayLabel = activeHoliday ? `khusus (${activeHoliday.name})` : 'terjadwal (Off Day)';
      return Swal.fire('Akses Ditolak', `Hari ini adalah hari libur ${holidayLabel}. Presensi dinonaktifkan.`, 'info');
    }

    if (!account) {
      return Swal.fire('Gagal', 'Data akun tidak ditemukan.', 'error');
    }
    
    if (!targetSchedule) {
      return Swal.fire('Gagal', 'Jadwal kerja Anda belum ditentukan.', 'warning');
    }

    if (!coords || distance === null) {
      return Swal.fire({
        title: 'Lokasi Belum Siap',
        text: 'Sinyal GPS sedang dioptimalkan. Mohon tunggu beberapa detik hingga indikator jarak muncul.',
        icon: 'warning',
        confirmButtonColor: '#006E62'
      });
    }

    const locationRadius = account.location?.radius || 100;
    const isCheckOut = !!todayAttendance && !todayAttendance.check_out;
    // FIX: Ensure strict boolean check to prevent string "false" or null issues
    const isLimited = isCheckOut 
      ? account.is_presence_limited_checkout === true 
      : account.is_presence_limited_checkin === true;

    if (isLimited && distance > locationRadius) {
       setIsCameraActive(false);
       return Swal.fire({
         title: 'Diluar Radius',
         text: `Anda berada diluar radius penempatan (${Math.round(distance)}m > ${locationRadius}m). Presensi dibatasi hanya di area lokasi.`,
         icon: 'error',
         confirmButtonColor: '#006E62'
       });
    }

    if (isCapturing) return;

    const scheduleResult = presenceService.calculateStatus(serverTime, targetSchedule, isCheckOut ? 'OUT' : 'IN');
    
    let reason = null;
    if (scheduleResult.status !== 'Tepat Waktu') {
      const { value: text, isConfirmed } = await Swal.fire({
        title: `Konfirmasi ${scheduleResult.status}`,
        input: 'textarea',
        inputLabel: `Anda terdeteksi ${scheduleResult.status.toLowerCase()} ${scheduleResult.minutes} menit. Harap berikan alasan:`,
        inputPlaceholder: 'Tuliskan alasan Anda di sini...',
        inputAttributes: { 'aria-label': 'Tuliskan alasan Anda di sini' },
        showCancelButton: true,
        confirmButtonColor: '#006E62',
        cancelButtonColor: '#d33',
        confirmButtonText: 'Kirim Presensi',
        cancelButtonText: 'Batal',
        inputValidator: (value) => {
          if (!value) return 'Alasan wajib diisi!';
          return null;
        }
      });

      if (!isConfirmed) return;
      reason = text;
    }

    try {
      setIsCapturing(true);
      
      // Capture current attendance state to avoid race conditions
      const currentAttendance = todayAttendance;
      const isCurrentlyCheckingOut = !!currentAttendance && !currentAttendance.check_out;

      // Tahap 1: Jalankan geocoding dan upload secara paralel untuk efisiensi
      const [address, photoId] = await Promise.all([
        presenceService.getReverseGeocode(coords.lat, coords.lng),
        googleDriveService.uploadFile(new File([photoBlob], `Presence_${isCurrentlyCheckingOut ? 'OUT' : 'IN'}_${Date.now()}.jpg`))
      ]);

      const currentTimeStr = serverTime.toISOString();
      
      // Tahap 2: Simpan ke database Supabase
      if (!isCurrentlyCheckingOut) {
        const payload: any = {
          account_id: account.id,
          check_in: currentTimeStr,
          in_latitude: coords.lat,
          in_longitude: coords.lng,
          in_photo_id: photoId,
          in_address: address,
          status_in: scheduleResult.status,
          late_minutes: scheduleResult.minutes,
          late_reason: reason
        };
        await presenceService.checkIn(payload);
      } else {
        if (!currentAttendance?.id) {
          throw new Error("ID referensi presensi tidak ditemukan. Harap muat ulang halaman.");
        }
        const payload: any = {
          check_out: currentTimeStr,
          out_latitude: coords.lat,
          out_longitude: coords.lng,
          out_photo_id: photoId,
          out_address: address,
          status_out: scheduleResult.status,
          early_departure_minutes: scheduleResult.minutes,
          early_departure_reason: reason
        };
        await presenceService.checkOut(currentAttendance.id, payload);
      }

      // Tahap 3: Finalisasi UI (hanya setelah simpan DB sukses)
      setIsCameraActive(false); 
      
      await Swal.fire({ 
        title: 'Berhasil!', 
        text: `Presensi ${isCurrentlyCheckingOut ? 'Pulang' : 'Masuk'} berhasil dicatat.`, 
        icon: 'success', 
        timer: 2000,
        showConfirmButton: false
      });
      
      // Refresh data tampilan
      await fetchInitialData();
    } catch (error: any) {
      console.error("Attendance Process Error:", error);
      Swal.fire({
        title: 'Gagal Menyimpan',
        text: error.message || 'Terjadi kesalahan sistem saat menyimpan data. Harap periksa koneksi internet Anda.',
        icon: 'error',
        confirmButtonColor: '#006E62'
      });
    } finally {
      setIsCapturing(false);
    }
  };

  if (isLoading) return <LoadingSpinner message="Sinkronisasi Data Satelit..." />;

  if (!account) {
    return (
      <div className="max-w-md mx-auto mt-20 p-8 bg-white rounded-3xl border border-gray-100 shadow-xl text-center">
        <div className="w-20 h-20 bg-rose-50 rounded-full flex items-center justify-center text-rose-500 mx-auto mb-6">
          <UserX size={40} />
        </div>
        <h3 className="text-xl font-bold text-gray-800">Akun Tidak Dikenali</h3>
        <p className="text-sm text-gray-500 mt-2">ID karyawan tidak terdaftar atau telah dinonaktifkan. Harap hubungi Admin HR.</p>
        <button 
          onClick={() => window.location.reload()}
          className="mt-8 w-full py-3 bg-[#006E62] text-white rounded-xl font-bold text-sm uppercase tracking-widest shadow-lg hover:bg-[#005a50] transition-all"
        >
          Coba Muat Ulang
        </button>
      </div>
    );
  }

  const isCheckOut = !!todayAttendance && !todayAttendance.check_out;
  const isLimited = isCheckOut 
    ? account.is_presence_limited_checkout === true 
    : account.is_presence_limited_checkin === true;
  const isWithinRadius = distance !== null && distance <= (account?.location?.radius || 100);
  const isBlockedByLocation = isLimited && !isWithinRadius;
  const todayDay = serverTime.getDay();
  
  // Resolve Rule untuk UI Tampilan Jadwal
  let displaySchedule = account.schedule;
  if (account.schedule_type === 'Shift Dinamis') displaySchedule = selectedShift || undefined;
  
  const scheduleRule = displaySchedule?.rules?.find(r => r.day_of_week === todayDay);
  const isHolidayToday = account.schedule_type !== 'Fleksibel' && (!!activeHoliday || !!scheduleRule?.is_holiday);

  return (
    <div className="max-w-5xl mx-auto space-y-6">
      <div className="bg-white rounded-2xl border border-gray-100 p-6 shadow-sm flex flex-col md:flex-row justify-between items-center gap-6">
        <div className="flex items-center gap-4">
          <div className="w-12 h-12 bg-[#006E62]/10 rounded-xl flex items-center justify-center text-[#006E62]">
            <Fingerprint size={28} />
          </div>
          <div>
            <h2 className="text-xl font-bold text-gray-800 tracking-tight">Presensi Reguler</h2>
            <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">{account?.full_name} • {account?.internal_nik}</p>
          </div>
        </div>
        <div className="flex gap-2 p-1 bg-gray-50 rounded-xl border border-gray-100">
          <button 
            onClick={() => setActiveTab('capture')}
            className={`flex items-center gap-2 px-6 py-2.5 rounded-lg text-xs font-bold uppercase transition-all ${activeTab === 'capture' ? 'bg-white text-[#006E62] shadow-sm' : 'text-gray-400 hover:text-gray-600'}`}
          >
            <Camera size={16} /> Presensi
          </button>
          <button 
            onClick={() => setActiveTab('history')}
            className={`flex items-center gap-2 px-6 py-2.5 rounded-lg text-xs font-bold uppercase transition-all ${activeTab === 'history' ? 'bg-white text-[#006E62] shadow-sm' : 'text-gray-400 hover:text-gray-600'}`}
          >
            <History size={16} /> Riwayat
          </button>
        </div>
      </div>

      {activeTab === 'capture' ? (
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start animate-in fade-in duration-500">
          <div className="lg:col-span-7">
            {isCameraActive ? (
              <PresenceCamera 
                onCapture={handleAttendance}
                onClose={() => setIsCameraActive(false)}
                isProcessing={isCapturing}
              />
            ) : isHolidayToday ? (
              <div className="bg-white rounded-2xl border border-gray-100 p-12 flex flex-col items-center justify-center shadow-sm text-center">
                <div className="w-24 h-24 rounded-full bg-rose-50 text-rose-500 flex items-center justify-center mb-8 shadow-xl">
                   <Umbrella size={48} />
                </div>
                <h3 className="text-2xl font-bold text-gray-800">
                  {activeHoliday ? 'Hari Libur Khusus' : 'Hari Libur Terjadwal'}
                </h3>
                <p className="text-sm text-rose-600 font-bold mt-2 max-w-xs uppercase tracking-tight">
                  {activeHoliday ? `"${activeHoliday.name}"` : '"Off Day / Hari Libur"'}
                </p>
                <div className="mt-8 p-4 bg-gray-50 rounded-xl border border-gray-100">
                  <p className="text-xs text-gray-500 leading-relaxed font-medium">Sesuai kebijakan manajemen, sistem presensi dinonaktifkan selama periode libur ini. Nikmati waktu istirahat Anda.</p>
                </div>
              </div>
            ) : (!todayAttendance || !todayAttendance.check_out) ? (
              <div className="bg-white rounded-2xl border border-gray-100 p-8 flex flex-col items-center justify-center shadow-sm text-center">
                <div className={`w-20 h-20 rounded-full flex items-center justify-center mb-6 shadow-xl transition-all duration-500 ${!isBlockedByLocation ? 'bg-emerald-50 text-[#006E62]' : 'bg-rose-50 text-rose-500'}`}>
                   {account.schedule_type === 'Shift Dinamis' ? <RefreshCw size={40} className={isFetchingShifts ? 'animate-spin' : ''} /> : <Fingerprint size={40} />}
                </div>
                <h3 className="text-xl font-bold text-gray-800">
                   {!!todayAttendance && !todayAttendance.check_out ? 'Waktunya Pulang?' : 'Siap Bekerja Hari Ini?'}
                </h3>
                
                {/* LOGIKA KHUSUS SHIFT DINAMIS: PILEH JADWAL (Cek via schedule_type) */}
                {account.schedule_type === 'Shift Dinamis' && !todayAttendance && (
                  <div className="mt-6 w-full max-w-sm space-y-3">
                    <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest text-left mb-2">Pilih Shift Kerja Hari Ini:</p>
                    {isFetchingShifts ? (
                      <div className="py-4 text-xs text-gray-400 italic">Mengambil daftar shift...</div>
                    ) : dynamicShifts.length === 0 ? (
                      <div className="p-4 bg-rose-50 border border-rose-100 rounded-xl text-[10px] text-rose-600 font-bold">Tidak ada jadwal Shift Kerja (Type 2) tersedia di lokasi Anda.</div>
                    ) : (
                      <div className="grid grid-cols-1 gap-2 max-h-60 overflow-y-auto pr-1">
                        {dynamicShifts.map(s => (
                          <button 
                            key={s.id}
                            onClick={() => setSelectedShift(s)}
                            className={`flex items-center justify-between p-3 rounded-xl border transition-all text-left group ${selectedShift?.id === s.id ? 'border-[#006E62] bg-emerald-50' : 'border-gray-100 bg-white hover:border-[#006E62]'}`}
                          >
                            <div className="flex-1">
                              <p className={`text-xs font-bold ${selectedShift?.id === s.id ? 'text-[#006E62]' : 'text-gray-700'}`}>{s.name}</p>
                              <p className="text-[9px] text-gray-400 uppercase font-bold">
                                {s.rules?.find(r => r.day_of_week === todayDay)?.check_in_time?.slice(0,5)} - {s.rules?.find(r => r.day_of_week === todayDay)?.check_out_time?.slice(0,5)}
                              </p>
                            </div>
                            <div className={`w-5 h-5 rounded-full border flex items-center justify-center transition-colors ${selectedShift?.id === s.id ? 'bg-[#006E62] border-[#006E62] text-white' : 'border-gray-200 group-hover:border-[#006E62]'}`}>
                               {selectedShift?.id === s.id && <Check size={12} />}
                            </div>
                          </button>
                        ))}
                      </div>
                    )}
                  </div>
                )}

                <p className="text-xs text-gray-400 mt-6 max-w-xs leading-tight">
                  {!isBlockedByLocation 
                    ? (account.schedule_type === 'Shift Dinamis' && !todayAttendance && !selectedShift 
                        ? 'Silahkan pilih salah satu shift di atas untuk mengaktifkan tombol verifikasi.'
                        : 'Klik tombol di bawah untuk verifikasi identitas AI dan mencatat lokasi Anda.')
                    : 'Akses presensi terkunci. Anda harus berada di area lokasi penempatan.'}
                </p>
                
                <button 
                  disabled={isBlockedByLocation || isCapturing || (account.schedule_type === 'Shift Dinamis' && !todayAttendance && !selectedShift)}
                  onClick={() => setIsCameraActive(true)}
                  className={`mt-8 flex items-center gap-3 px-12 py-4 rounded-2xl font-bold uppercase text-xs tracking-widest shadow-lg transition-all ${
                    !isBlockedByLocation && !isCapturing && (account.schedule_type !== 'Shift Dinamis' || !!todayAttendance || !!selectedShift)
                    ? 'bg-[#006E62] text-white hover:bg-[#005a50] hover:scale-105 active:scale-95' 
                    : 'bg-gray-100 text-gray-300 cursor-not-allowed shadow-none'
                  }`}
                >
                  <Camera size={18} />
                  {isCapturing ? 'MEMPROSES...' : 'MULAI VERIFIKASI WAJAH'}
                </button>
              </div>
            ) : (
              <div className="bg-white rounded-2xl border border-gray-100 p-20 flex flex-col items-center justify-center shadow-sm text-center">
                <div className="w-20 h-20 bg-emerald-100 rounded-full flex items-center justify-center text-[#006E62] mb-6 animate-pulse">
                  <ShieldCheck size={48} />
                </div>
                <h3 className="text-2xl font-bold text-gray-800">Tugas Selesai!</h3>
                <p className="text-sm text-gray-500 mt-2 max-w-xs">Terima kasih, Anda telah menyelesaikan presensi masuk dan pulang untuk hari ini.</p>
              </div>
            )}
          </div>

          <div className="lg:col-span-5 space-y-6">
            <div className="bg-white rounded-2xl border border-gray-100 p-6 shadow-sm overflow-hidden relative">
               <div className="absolute top-0 right-0 p-4 opacity-5">
                  <Clock size={120} />
               </div>
               <div className="flex items-center justify-between mb-6">
                  <div className="flex items-center gap-2">
                    <Clock size={16} className="text-[#006E62]" />
                    <h4 className="text-[11px] font-bold uppercase tracking-widest text-gray-500">Waktu Terverifikasi</h4>
                  </div>
                  <span className="text-[10px] font-bold text-[#00FFE4] bg-[#006E62]/5 px-2 py-0.5 rounded uppercase tracking-tighter">
                    {displaySchedule?.name || (account.schedule_type === 'Fleksibel' ? 'Fleksibel' : 'Reguler')}
                  </span>
               </div>
               <div className="text-center py-4 relative z-10">
                  <div className="text-5xl font-mono font-bold text-gray-800 tracking-tighter">
                    {serverTime.toLocaleTimeString('id-ID', { hour12: false })}
                  </div>
                  <div className="text-[11px] font-bold text-[#00FFE4] uppercase tracking-widest mt-2">
                    {serverTime.toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long' })}
                  </div>

                  {todayAttendance?.check_in && !todayAttendance.check_out && (
                    <div className="mt-8 p-3 bg-[#006E62]/5 rounded-2xl border border-emerald-100/50 animate-in fade-in zoom-in duration-700">
                      <p className="text-[9px] font-black text-gray-400 uppercase tracking-[0.2em] mb-1">Durasi Kerja Berjalan</p>
                      <div className="text-2xl font-mono font-black text-[#006E62] tracking-widest">
                        {getLiveWorkDuration()}
                      </div>
                    </div>
                  )}
               </div>

               {account.schedule_type !== 'Fleksibel' && (
                 <div className="mt-8 p-4 bg-emerald-50/50 rounded-xl border border-emerald-100/50 space-y-3">
                    <div className="flex items-center gap-2 text-[10px] font-bold text-[#006E62] uppercase tracking-wider mb-2">
                      <CalendarClock size={14} /> {account.schedule_type === 'Shift Dinamis' ? 'Shift Terpilih' : 'Jadwal Hari Ini'}
                    </div>
                    <div className="grid grid-cols-2 gap-4">
                       <div>
                          <p className="text-[9px] text-gray-400 font-bold uppercase">Jam Masuk</p>
                          <p className="text-xs font-bold text-gray-700">{scheduleRule?.check_in_time ? scheduleRule.check_in_time.slice(0, 5) : '--:--'}</p>
                       </div>
                       <div>
                          <p className="text-[9px] text-gray-400 font-bold uppercase">Jam Pulang</p>
                          <p className="text-xs font-bold text-gray-700">{scheduleRule?.check_out_time ? scheduleRule.check_out_time.slice(0, 5) : '--:--'}</p>
                       </div>
                    </div>
                    <div className="pt-2 border-t border-emerald-100/50">
                      <p className="text-[9px] text-gray-400 font-bold uppercase">Toleransi</p>
                      <p className="text-[10px] font-medium text-[#006E62]">
                        Masuk: {displaySchedule?.tolerance_checkin_minutes || 0}m • Pulang: {displaySchedule?.tolerance_minutes || 0}m
                      </p>
                    </div>
                 </div>
               )}

               <div className="grid grid-cols-2 gap-4 mt-6 pt-6 border-t border-gray-50">
                  <div className="text-center">
                    <p className="text-[9px] font-bold text-gray-400 uppercase mb-1">Status Masuk</p>
                    <p className={`text-xs font-bold ${todayAttendance?.status_in === 'Terlambat' ? 'text-rose-500' : 'text-[#006E62]'}`}>
                      {todayAttendance?.check_in ? todayAttendance.status_in : '--:--'}
                    </p>
                  </div>
                  <div className="text-center border-l border-gray-50">
                    <p className="text-[9px] font-bold text-gray-400 uppercase mb-1">Status Pulang</p>
                    <p className={`text-xs font-bold ${todayAttendance?.status_out === 'Pulang Cepat' ? 'text-rose-500' : 'text-blue-500'}`}>
                      {todayAttendance?.check_out ? todayAttendance.status_out : '--:--'}
                    </p>
                  </div>
               </div>
            </div>

            <div className="bg-white rounded-2xl border border-gray-100 p-6 shadow-sm">
               <div className="flex items-center justify-between mb-4">
                  <div className="flex items-center gap-2">
                    <MapPin size={16} className="text-[#00FFE4]" />
                    <h4 className="text-[11px] font-bold uppercase tracking-widest text-gray-500">Status Geotag</h4>
                  </div>
                  <div className={`flex items-center gap-1.5 px-2 py-1 rounded-full text-[9px] font-bold uppercase ${!isLimited ? 'bg-blue-50 text-blue-600' : (isWithinRadius ? 'bg-emerald-50 text-emerald-600' : 'bg-rose-50 text-rose-600')}`}>
                    <div className={`w-1.5 h-1.5 rounded-full ${!isLimited ? 'bg-blue-500' : (isWithinRadius ? 'bg-emerald-500' : 'bg-rose-500')}`}></div>
                    {!isLimited ? 'Bebas Lokasi' : (isWithinRadius ? 'Dalam Radius' : 'Luar Radius')}
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
                          <span className={`text-sm font-bold ${isWithinRadius ? 'text-[#006E62]' : 'text-rose-500'}`}>{distance !== null ? `${Math.round(distance)}m` : 'Calculating...'}</span>
                       </div>
                       <div className="p-3 bg-gray-50 rounded-xl border border-gray-100">
                          <span className="text-gray-400 font-bold uppercase block mb-1">Maks Radius</span>
                          <span className="text-sm font-bold text-gray-700">{account.location.radius}m</span>
                       </div>
                    </div>
                    {isBlockedByLocation && (
                      <div className="p-3 bg-rose-50 border border-rose-100 rounded-xl flex gap-3 items-start animate-pulse">
                         <AlertCircle size={16} className="text-rose-500 shrink-0 mt-0.5" />
                         <p className="text-[10px] text-rose-600 font-bold leading-tight uppercase tracking-tight">Presensi dikunci. Anda berada diluar zona yang diizinkan.</p>
                      </div>
                    )}
                 </div>
               ) : (
                 <div className="flex flex-col items-center justify-center py-14 text-gray-300">
                    <MapIcon size={40} strokeWidth={1} className="animate-bounce" />
                    <p className="text-[10px] font-bold uppercase tracking-widest mt-4">Mengunci Sinyal GPS...</p>
                 </div>
               )}
            </div>
          </div>
        </div>
      ) : (
        <div className="animate-in fade-in slide-in-from-bottom-4 duration-500">
          <PresenceHistory logs={recentLogs} isLoading={isLoading} />
        </div>
      )}
    </div>
  );
};

export default PresenceMain;
