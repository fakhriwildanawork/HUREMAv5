
import { supabase } from '../lib/supabase';
import { Attendance, AttendanceInput, Account, Schedule, ScheduleRule } from '../types';

const sanitizePayload = (payload: any) => {
  const sanitized = { ...payload };
  Object.keys(sanitized).forEach(key => {
    if (sanitized[key] === '' || sanitized[key] === undefined) {
      sanitized[key] = null;
    }
  });
  return sanitized;
};

export const presenceService = {
  /**
   * Mengambil waktu server (Anti-Fake Time)
   */
  async getServerTime(): Promise<Date> {
    const { data, error } = await supabase.rpc('get_server_time');
    if (error) {
      try {
        const res = await fetch('https://worldtimeapi.org/api/timezone/Asia/Jakarta');
        const json = await res.json();
        return new Date(json.datetime);
      } catch (e) {
        return new Date();
      }
    }
    return new Date(data);
  },

  /**
   * Mendapatkan alamat dari koordinat (Nominatim API) dengan timeout
   */
  async getReverseGeocode(lat: number, lng: number): Promise<string> {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 3500); // Timeout 3.5 detik

    try {
      const response = await fetch(`https://nominatim.openstreetmap.org/reverse?format=json&lat=${lat}&lon=${lng}&zoom=18&addressdetails=1`, {
        signal: controller.signal
      });
      clearTimeout(timeoutId);
      const data = await response.json();
      return data.display_name || `${lat.toFixed(6)}, ${lng.toFixed(6)}`;
    } catch (err) {
      clearTimeout(timeoutId);
      console.error("Reverse Geotag Error (Using Fallback):", err);
      return `Koordinat: ${lat.toFixed(6)}, ${lng.toFixed(6)}`;
    }
  },

  /**
   * Menghitung keterlambatan atau pulang cepat berdasarkan jadwal
   */
  calculateStatus(currentTime: Date, schedule: Schedule, type: 'IN' | 'OUT'): { status: string, minutes: number } {
    // Mode Fleksibel Bypass
    if (schedule.id === 'FLEKSIBEL') {
      return { status: 'Tepat Waktu', minutes: 0 };
    }

    const dayOfWeek = currentTime.getDay();
    const rule = schedule.rules?.find(r => r.day_of_week === dayOfWeek);

    if (!rule || rule.is_holiday) return { status: 'Tepat Waktu', minutes: 0 };

    const [targetH, targetM] = (type === 'IN' ? rule.check_in_time : rule.check_out_time || '00:00:00').split(':').map(Number);
    const targetDate = new Date(currentTime);
    targetDate.setHours(targetH, targetM, 0, 0);

    const diffMs = currentTime.getTime() - targetDate.getTime();
    const diffMins = Math.floor(diffMs / 60000);

    if (type === 'IN') {
      const tolerance = schedule.tolerance_checkin_minutes || 0;
      if (diffMins > tolerance) {
        return { status: 'Terlambat', minutes: diffMins };
      }
    } else {
      const tolerance = schedule.tolerance_minutes || 0;
      // Jika pulang sebelum waktu seharusnya (diffMins negatif) diluar toleransi
      if (diffMins < -tolerance) {
        return { status: 'Pulang Cepat', minutes: Math.abs(diffMins) };
      }
    }

    return { status: 'Tepat Waktu', minutes: 0 };
  },

  /**
   * Menghitung jarak antara 2 koordinat (Haversine Formula)
   */
  calculateDistance(lat1: number, lon1: number, lat2: number, lon2: number): number {
    const R = 6371e3; // Earth radius in meters
    const φ1 = lat1 * Math.PI / 180;
    const φ2 = lat2 * Math.PI / 180;
    const Δφ = (lat2 - lat1) * Math.PI / 180;
    const Δλ = (lon2 - lon1) * Math.PI / 180;

    const a = Math.sin(Δφ / 2) * Math.sin(Δφ / 2) +
              Math.cos(φ1) * Math.cos(φ2) *
              Math.sin(Δλ / 2) * Math.sin(Δλ / 2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));

    return R * c; // Distance in meters
  },

  async getTodayAttendance(accountId: string) {
    const today = new Date().toISOString().split('T')[0];
    const { data, error } = await supabase
      .from('attendances')
      .select('*')
      .eq('account_id', accountId)
      .gte('created_at', `${today}T00:00:00Z`)
      .order('created_at', { ascending: false })
      .limit(1)
      .maybeSingle();
    
    if (error) throw error;
    return data as Attendance | null;
  },

  /**
   * Memastikan user tidak sedang dalam sesi kerja reguler (Mutual Exclusion)
   */
  async isRegularSessionActive(accountId: string): Promise<boolean> {
    const attendance = await this.getTodayAttendance(accountId);
    return !!(attendance && attendance.check_in && !attendance.check_out);
  },

  /**
   * Cek apakah hari ini Libur Khusus (Tipe 3) di lokasi user
   */
  async checkHolidayStatus(accountId: string, locationId: string, checkDate: Date) {
    const dateStr = checkDate.toISOString().split('T')[0];
    
    const { data, error } = await supabase
      .from('schedules')
      .select('*, schedule_locations!inner(location_id)')
      .eq('type', 3) // Libur Khusus
      .eq('schedule_locations.location_id', locationId)
      .lte('start_date', dateStr)
      .gte('end_date', dateStr);

    if (error) return null;
    
    // Filter out if user is excluded
    const activeHoliday = data?.find(s => !s.excluded_account_ids?.includes(accountId));
    return activeHoliday || null;
  },

  async checkIn(input: Partial<AttendanceInput>) {
    const sanitized = sanitizePayload(input);
    const { data, error } = await supabase
      .from('attendances')
      .insert([sanitized])
      .select()
      .single();
    
    if (error) throw error;
    return data as Attendance;
  },

  async checkOut(id: string, input: Partial<AttendanceInput>) {
    if (!id) throw new Error("ID presensi tidak valid untuk proses check-out.");
    const sanitized = sanitizePayload(input);
    const { data, error } = await supabase
      .from('attendances')
      .update(sanitized)
      .eq('id', id)
      .select()
      .single();
    
    if (error) throw error;
    return data as Attendance;
  },

  async getRecentHistory(accountId: string, limit = 31) {
    const { data, error } = await supabase
      .from('attendances')
      .select('*')
      .eq('account_id', accountId)
      .order('created_at', { ascending: false })
      .limit(limit);
    
    if (error) throw error;
    return data as Attendance[];
  },

  async getAttendanceByRange(startDate: string, endDate: string, accountId?: string) {
    let query = supabase
      .from('attendances')
      .select('*')
      .gte('created_at', `${startDate}T00:00:00Z`)
      .lte('created_at', `${endDate}T23:59:59Z`)
      .order('created_at', { ascending: true });
    
    if (accountId) {
      query = query.eq('account_id', accountId);
    }

    const { data, error } = await query;
    
    if (error) throw error;
    return data as Attendance[];
  }
};
