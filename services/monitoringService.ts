
import { supabase } from '../lib/supabase';
import { accountService } from './accountService';
import { scheduleService } from './scheduleService';

export const monitoringService = {
  async getDailyMonitoringData(date: Date = new Date()) {
    // Force Jakarta Timezone for calculations
    const jakartaDateFormatter = new Intl.DateTimeFormat('en-CA', {
      timeZone: 'Asia/Jakarta',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit'
    });
    
    const jakartaDayFormatter = new Intl.DateTimeFormat('en-US', {
      timeZone: 'Asia/Jakarta',
      weekday: 'short'
    });

    const dateStr = jakartaDateFormatter.format(date);
    const dayStr = jakartaDayFormatter.format(date);
    const daysMap: { [key: string]: number } = { 'Sun': 0, 'Mon': 1, 'Tue': 2, 'Wed': 3, 'Thu': 4, 'Fri': 5, 'Sat': 6 };
    const dayOfWeek = daysMap[dayStr];

    // 1. Fetch all active accounts
    const accounts = await accountService.getAll();

    // 2. Fetch all schedules and their rules
    const schedules = await scheduleService.getAll();

    // 3. Fetch today's attendances (using Jakarta date range)
    // We use gte/lte on created_at which is UTC, so we need to be careful.
    // However, the app usually stores date-only fields or we can query by date string if available.
    // Given the previous code used created_at with T00:00:00Z, let's stick to a more robust approach:
    // Querying by the date part if possible, or calculating the UTC range for Jakarta's day.
    const startOfDay = new Date(`${dateStr}T00:00:00+07:00`).toISOString();
    const endOfDay = new Date(`${dateStr}T23:59:59+07:00`).toISOString();

    const { data: attendances } = await supabase
      .from('attendances')
      .select('*')
      .gte('created_at', startOfDay)
      .lte('created_at', endOfDay);

    // 4. Fetch today's overtimes
    const { data: overtimes } = await supabase
      .from('overtimes')
      .select('*')
      .gte('created_at', startOfDay)
      .lte('created_at', endOfDay);

    // 5. Fetch approved leaves (Cuti Tahunan)
    const { data: annualLeaves } = await supabase
      .from('account_annual_leaves')
      .select('*')
      .eq('status', 'approved')
      .lte('start_date', dateStr)
      .gte('end_date', dateStr);

    // 6. Fetch approved Libur Mandiri
    const { data: leaveRequests } = await supabase
      .from('account_leave_requests')
      .select('*')
      .eq('status', 'approved')
      .lte('start_date', dateStr)
      .gte('end_date', dateStr);

    // 7. Fetch approved Permissions
    const { data: permissions } = await supabase
      .from('account_permission_requests')
      .select('*')
      .eq('status', 'approved')
      .lte('start_date', dateStr)
      .gte('end_date', dateStr);

    // 8. Fetch approved Maternity Leaves
    const { data: maternityLeaves } = await supabase
      .from('account_maternity_leaves')
      .select('*')
      .eq('status', 'approved')
      .lte('start_date', dateStr)
      .gte('end_date', dateStr);

    // 9. Fetch Libur Khusus (Holiday) from schedules
    const holidaySchedules = schedules.filter(s => s.type === 3 && dateStr >= (s.start_date || '') && dateStr <= (s.end_date || ''));

    // Map data for easy lookup
    const attendanceMap = new Map(attendances?.map(a => [a.account_id, a]));
    const overtimeMap = new Map(overtimes?.map(o => [o.account_id, o]));
    const annualLeaveMap = new Map(annualLeaves?.map(l => [l.account_id, l]));
    const leaveRequestMap = new Map(leaveRequests?.map(l => [l.account_id, l]));
    const permissionMap = new Map(permissions?.map(p => [p.account_id, p]));
    const maternityLeaveMap = new Map(maternityLeaves?.map(m => [m.account_id, m]));
    const scheduleMap = new Map(schedules.map(s => [s.id, s]));

    const results = {
      present: [] as any[],
      notPresentYet: [] as any[],
      onHoliday: [] as any[],
      onLeaveMandiri: [] as any[],
      onOvertime: [] as any[],
      onAnnualLeave: [] as any[],
      onMaternityLeave: [] as any[],
      onPermission: [] as any[]
    };

    accounts.forEach((acc: any) => {
      const attendance = attendanceMap.get(acc.id);
      const overtime = overtimeMap.get(acc.id);
      const annualLeave = annualLeaveMap.get(acc.id);
      const leaveRequest = leaveRequestMap.get(acc.id);
      const permission = permissionMap.get(acc.id);
      const maternityLeave = maternityLeaveMap.get(acc.id);

      // Check if it's a holiday for this user
      const schedule = scheduleMap.get(acc.schedule_id);
      const rule = schedule?.rules?.find(r => r.day_of_week === dayOfWeek);
      
      // FIX: For Fleksibel/Dinamis, if no rule is found, it's NOT a holiday by default.
      // It's only a holiday if explicitly marked as is_holiday.
      let isWeekendOrHolidayRule = false;
      if (acc.schedule_type === 'Tetap') {
        isWeekendOrHolidayRule = !rule || rule.is_holiday;
      } else {
        // Fleksibel/Dinamis: Only holiday if rule exists and is_holiday is true
        isWeekendOrHolidayRule = rule?.is_holiday || false;
      }
      
      const isSpecialHoliday = holidaySchedules.some(hs => 
        hs.location_ids?.includes(acc.location_id) && 
        !(hs.excluded_account_ids || []).includes(acc.id)
      );

      const isHoliday = isWeekendOrHolidayRule || isSpecialHoliday;

      // Attach schedule info for UI
      const accWithInfo = { 
        ...acc, 
        schedule_name: schedule?.name || acc.schedule_type,
        today_rule: rule 
      };

      // PRIORITY HIERARCHY: One user = One primary category
      // 1. Present (Check-in)
      if (attendance) {
        results.present.push({ ...accWithInfo, attendance });
      } 
      // 2. On Leave / Permission / Maternity (Approved)
      else if (annualLeave) {
        results.onAnnualLeave.push({ ...accWithInfo, annualLeave });
      } else if (maternityLeave) {
        results.onMaternityLeave.push({ ...accWithInfo, maternityLeave });
      } else if (permission) {
        results.onPermission.push({ ...accWithInfo, permission });
      } 
      // 3. Holiday (Mandiri / Special / Weekend)
      else if (leaveRequest) {
        results.onLeaveMandiri.push({ ...accWithInfo, leaveRequest });
      } else if (isHoliday) {
        results.onHoliday.push(accWithInfo);
      } 
      // 4. Not Present Yet (Must work but no check-in)
      else {
        results.notPresentYet.push(accWithInfo);
      }

      // Overtime is a secondary status (can be present AND on overtime)
      if (overtime) {
        results.onOvertime.push({ ...accWithInfo, overtime });
      }
    });

    return results;
  }
};
