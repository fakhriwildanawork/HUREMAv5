
import { supabase } from '../lib/supabase';
import { Schedule, ScheduleInput, ScheduleRule } from '../types';

/**
 * Fungsi pembantu untuk membersihkan data sebelum dikirim ke Supabase.
 * Mengubah string kosong ('') menjadi null agar tidak error saat masuk ke kolom DATE atau TIME.
 */
const sanitizePayload = (payload: any) => {
  const sanitized = { ...payload };
  Object.keys(sanitized).forEach(key => {
    if (sanitized[key] === '' || sanitized[key] === undefined) {
      sanitized[key] = null;
    }
  });
  return sanitized;
};

export const scheduleService = {
  async getAll() {
    const { data, error } = await supabase
      .from('schedules')
      .select(`
        *,
        schedule_locations(location_id),
        schedule_rules(*)
      `)
      .order('created_at', { ascending: false });
    
    if (error) throw error;
    return data.map(item => ({
      ...item,
      location_ids: item.schedule_locations.map((sl: any) => sl.location_id),
      rules: item.schedule_rules
    })) as Schedule[];
  },

  async getByLocation(locationId: string) {
    if (!locationId) return [];
    const { data, error } = await supabase
      .from('schedule_locations')
      .select('schedule_id, schedules(*)')
      .eq('location_id', locationId);
    
    if (error) throw error;
    return (data as any[]).map(item => item.schedules) as unknown as Schedule[];
  },

  async getById(id: string) {
    const { data, error } = await supabase
      .from('schedules')
      .select(`
        *,
        schedule_locations(location_id),
        schedule_rules(*)
      `)
      .eq('id', id)
      .single();
    
    if (error) throw error;
    return {
      ...data,
      location_ids: data.schedule_locations.map((sl: any) => sl.location_id),
      rules: data.schedule_rules
    } as Schedule;
  },

  async create(input: ScheduleInput) {
    let { rules, location_ids, ...scheduleData } = input;
    
    // Sanitasi Tipe 3 (Libur Khusus): Tidak boleh ada aturan jam dan toleransi
    if (scheduleData.type === 3) {
      scheduleData.tolerance_minutes = 0;
      scheduleData.tolerance_checkin_minutes = 0;
      rules = [];
    }

    // 1. Insert schedule with sanitization
    const sanitizedSchedule = sanitizePayload(scheduleData);
    const { data: schedule, error: sError } = await supabase
      .from('schedules')
      .insert([sanitizedSchedule])
      .select()
      .single();
    
    if (sError) throw sError;

    // 2. Insert rules with sanitization
    if (rules && rules.length > 0) {
      const rulesToInsert = rules.map(r => sanitizePayload({ ...r, schedule_id: schedule.id }));
      const { error: rError } = await supabase
        .from('schedule_rules')
        .insert(rulesToInsert);
      
      if (rError) throw rError;
    }

    // 3. Insert locations
    if (location_ids && location_ids.length > 0) {
      const locationsToInsert = location_ids.map(lid => ({ schedule_id: schedule.id, location_id: lid }));
      const { error: lError } = await supabase
        .from('schedule_locations')
        .insert(locationsToInsert);
      
      if (lError) throw lError;
    }

    return schedule as Schedule;
  },

  async update(id: string, input: Partial<ScheduleInput>) {
    let { rules, location_ids, ...scheduleData } = input;

    // Sanitasi Tipe 3 (Libur Khusus)
    if (scheduleData.type === 3) {
      scheduleData.tolerance_minutes = 0;
      scheduleData.tolerance_checkin_minutes = 0;
      rules = [];
    }

    // 1. Update master with sanitization
    const sanitizedSchedule = sanitizePayload(scheduleData);
    const { data: schedule, error: sError } = await supabase
      .from('schedules')
      .update(sanitizedSchedule)
      .eq('id', id)
      .select()
      .single();
    
    if (sError) throw sError;

    // 2. Update rules (delete and re-insert for simplicity)
    if (rules !== undefined) {
      await supabase.from('schedule_rules').delete().eq('schedule_id', id);
      if (rules.length > 0) {
        const rulesToInsert = rules.map(r => sanitizePayload({ ...r, schedule_id: id }));
        await supabase.from('schedule_rules').insert(rulesToInsert);
      }
    }

    // 3. Update locations
    if (location_ids) {
      await supabase.from('schedule_locations').delete().eq('schedule_id', id);
      if (location_ids.length > 0) {
        const locationsToInsert = location_ids.map(lid => ({ schedule_id: id, location_id: lid }));
        await supabase.from('schedule_locations').insert(locationsToInsert);
      }
    }

    return schedule as Schedule;
  },

  async delete(id: string) {
    const { error } = await supabase.from('schedules').delete().eq('id', id);
    if (error) throw error;
    return true;
  },

  /**
   * Mendapatkan Libur Khusus (Tipe 3) dalam rentang tanggal untuk lokasi tertentu
   */
  async getSpecialHolidaysByRange(locationId: string, startDate: string, endDate: string) {
    const { data, error } = await supabase
      .from('schedules')
      .select('*, schedule_locations!inner(location_id)')
      .eq('type', 3) // Libur Khusus
      .eq('schedule_locations.location_id', locationId)
      .lte('start_date', endDate)
      .gte('end_date', startDate);

    if (error) throw error;
    return data as Schedule[];
  }
};
