import { supabase } from '../lib/supabase';
import { EmployeeReportData, AttendanceSummary, LeaveSummary, OvertimeSummary, PayrollSummary } from '../types';

export const reportService = {
  async getEmployeeReportData(): Promise<EmployeeReportData> {
    // Fetch all accounts
    const { data: accounts, error: accountsError } = await supabase
      .from('accounts')
      .select('*, location:locations(name)');

    if (accountsError) throw new Error(accountsError.message);

    // Fetch terminations for exit employees (e.g., in the last 30 days)
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
    
    const { data: terminations } = await supabase
      .from('termination_logs')
      .select('*')
      .gte('termination_date', thirtyDaysAgo.toISOString());

    // Fetch warnings for discipline summary
    const { data: warnings } = await supabase
      .from('warning_logs')
      .select('*');

    const totalEmployees = accounts.length;
    
    // New employees in last 30 days
    const newEmployees = accounts.filter(a => {
      if (!a.start_date) return false;
      return new Date(a.start_date) >= thirtyDaysAgo;
    }).length;

    const exitEmployees = terminations?.length || 0;

    // Gender Ratio
    const genderMap = accounts.reduce((acc: any, curr) => {
      const g = curr.gender || 'Lainnya';
      acc[g] = (acc[g] || 0) + 1;
      return acc;
    }, {});
    const genderRatio = Object.entries(genderMap).map(([name, value]) => ({ name, value: value as number }));

    // Age Distribution
    const ageDistribution = [
      { name: '< 20', value: 0 },
      { name: '20-30', value: 0 },
      { name: '31-40', value: 0 },
      { name: '41-50', value: 0 },
      { name: '> 50', value: 0 },
    ];
    accounts.forEach(a => {
      if (!a.dob) return;
      const age = new Date().getFullYear() - new Date(a.dob).getFullYear();
      if (age < 20) ageDistribution[0].value++;
      else if (age <= 30) ageDistribution[1].value++;
      else if (age <= 40) ageDistribution[2].value++;
      else if (age <= 50) ageDistribution[3].value++;
      else ageDistribution[4].value++;
    });

    // Education Distribution
    const eduMap = accounts.reduce((acc: any, curr) => {
      const edu = curr.last_education || 'Tidak Diketahui';
      acc[edu] = (acc[edu] || 0) + 1;
      return acc;
    }, {});
    const educationDistribution = Object.entries(eduMap).map(([name, value]) => ({ name, value: value as number }));

    // Location Distribution
    const locMap = accounts.reduce((acc: any, curr) => {
      const loc = curr.location?.name || 'Tidak Diketahui';
      acc[loc] = (acc[loc] || 0) + 1;
      return acc;
    }, {});
    const locationDistribution = Object.entries(locMap).map(([name, value]) => ({ name, value: value as number }));

    // Position Distribution
    const posMap = accounts.reduce((acc: any, curr) => {
      const pos = curr.position || 'Tidak Diketahui';
      acc[pos] = (acc[pos] || 0) + 1;
      return acc;
    }, {});
    const positionDistribution = Object.entries(posMap).map(([name, value]) => ({ name, value: value as number }));

    // Contract Type Distribution
    const contractMap = accounts.reduce((acc: any, curr) => {
      const type = curr.employee_type || 'Tidak Diketahui';
      acc[type] = (acc[type] || 0) + 1;
      return acc;
    }, {});
    const contractTypeDistribution = Object.entries(contractMap).map(([name, value]) => ({ name, value: value as number }));

    // Tenure Distribution
    const tenureDistribution = [
      { name: '< 1 Thn', value: 0 },
      { name: '1-3 Thn', value: 0 },
      { name: '3-5 Thn', value: 0 },
      { name: '> 5 Thn', value: 0 },
    ];
    accounts.forEach(a => {
      if (!a.start_date) return;
      const years = (new Date().getTime() - new Date(a.start_date).getTime()) / (1000 * 60 * 60 * 24 * 365);
      if (years < 1) tenureDistribution[0].value++;
      else if (years <= 3) tenureDistribution[1].value++;
      else if (years <= 5) tenureDistribution[2].value++;
      else tenureDistribution[3].value++;
    });

    // Health Risk Profile
    const healthMap = accounts.reduce((acc: any, curr) => {
      const risk = curr.health_risk || 'Tidak Diketahui';
      acc[risk] = (acc[risk] || 0) + 1;
      return acc;
    }, {});
    const healthRiskProfile = Object.entries(healthMap).map(([name, value]) => ({ name, value: value as number }));

    // Discipline Summary
    const warningMap = (warnings || []).reduce((acc: any, curr) => {
      acc[curr.warning_type] = (acc[curr.warning_type] || 0) + 1;
      return acc;
    }, {});
    const disciplineSummary = Object.entries(warningMap).map(([name, value]) => ({ name, value: value as number }));

    return {
      totalEmployees,
      newEmployees,
      exitEmployees,
      genderRatio,
      ageDistribution,
      educationDistribution,
      locationDistribution,
      positionDistribution,
      contractTypeDistribution,
      tenureDistribution,
      healthRiskProfile,
      disciplineSummary,
    };
  },

  async getAttendanceReportSummary(startDate: string, endDate: string): Promise<AttendanceSummary[]> {
    const [
      { data: accounts },
      { data: attendances },
      { data: leaves },
      { data: annualLeaves },
      { data: permissions },
      { data: maternityLeaves },
      { data: dispensations },
      { data: holidays }
    ] = await Promise.all([
      supabase.from('accounts').select('id, full_name, internal_nik'),
      supabase.from('attendances').select('*').gte('check_in', startDate).lte('check_in', endDate),
      supabase.from('leave_requests').select('*').eq('status', 'approved').gte('start_date', startDate).lte('end_date', endDate),
      supabase.from('annual_leave_requests').select('*').eq('status', 'approved').gte('start_date', startDate).lte('end_date', endDate),
      supabase.from('permission_requests').select('*').eq('status', 'approved').gte('start_date', startDate).lte('end_date', endDate),
      supabase.from('maternity_leave_requests').select('*').eq('status', 'approved').gte('start_date', startDate).lte('end_date', endDate),
      supabase.from('dispensation_requests').select('*').eq('status', 'APPROVED').gte('date', startDate).lte('date', endDate),
      supabase.from('holidays').select('*').gte('date', startDate).lte('date', endDate)
    ]);

    if (!accounts) return [];

    const start = new Date(startDate);
    const end = new Date(endDate);
    const totalDays = Math.ceil((end.getTime() - start.getTime()) / (1000 * 60 * 60 * 24)) + 1;

    return accounts.map(acc => {
      const accAttendances = (attendances || []).filter(a => a.account_id === acc.id);
      const accLeaves = (leaves || []).filter(l => l.account_id === acc.id);
      const accAnnualLeaves = (annualLeaves || []).filter(l => l.account_id === acc.id);
      const accPermissions = (permissions || []).filter(p => p.account_id === acc.id);
      const accMaternity = (maternityLeaves || []).filter(m => m.account_id === acc.id);
      const accDispensations = (dispensations || []).filter(d => d.account_id === acc.id);

      const present = accAttendances.length;
      const late = accAttendances.filter(a => a.late_minutes > 0).length;
      const lateMinutes = accAttendances.reduce((sum, a) => sum + (a.late_minutes || 0), 0);
      const earlyDeparture = accAttendances.filter(a => a.early_leave_minutes > 0).length;
      const earlyDepartureMinutes = accAttendances.reduce((sum, a) => sum + (a.early_leave_minutes || 0), 0);
      const noClockOut = accAttendances.filter(a => !a.check_out).length;
      
      const leave = accLeaves.length + accAnnualLeaves.length;
      const maternityLeave = accMaternity.length;
      const permission = accPermissions.length;
      const dispensationCount = accDispensations.length;

      // Simple absent calculation: total days - (present + leave + maternity + permission + holidays)
      // Note: This is a simplification and might need refinement based on actual work days
      const holidayCount = (holidays || []).length;
      const absent = Math.max(0, totalDays - (present + leave + maternityLeave + permission + holidayCount));

      const attendanceRate = totalDays > 0 ? ((present + leave + maternityLeave + permission) / totalDays) * 100 : 0;

      return {
        accountId: acc.id,
        fullName: acc.full_name,
        nik: acc.internal_nik,
        totalDays,
        present,
        late,
        lateMinutes,
        earlyDeparture,
        earlyDepartureMinutes,
        absent,
        leave,
        maternityLeave,
        permission,
        holiday: holidayCount,
        specialHoliday: 0,
        noClockOut,
        dispensationCount,
        attendanceRate,
        dailyDetails: [] // This would need a day-by-day map if required by heatmap
      };
    });
  },

  async getAttendanceReport(startDate?: string, endDate?: string) {
    let accQuery = supabase.from('accounts').select('*, location:locations(name)');
    let attQuery = supabase.from('attendances').select('*');
    let otQuery = supabase.from('overtimes').select('*');
    let lQuery = supabase.from('leave_requests').select('*').eq('status', 'approved');
    let alQuery = supabase.from('annual_leave_requests').select('*').eq('status', 'approved');
    let pQuery = supabase.from('permission_requests').select('*').eq('status', 'approved');
    let mlQuery = supabase.from('maternity_leave_requests').select('*').eq('status', 'approved');

    if (startDate) {
      attQuery = attQuery.gte('check_in', startDate);
      otQuery = otQuery.gte('check_in', startDate);
      lQuery = lQuery.gte('start_date', startDate);
      alQuery = alQuery.gte('start_date', startDate);
      pQuery = pQuery.gte('start_date', startDate);
      mlQuery = mlQuery.gte('start_date', startDate);
    }
    if (endDate) {
      attQuery = attQuery.lte('check_in', endDate);
      otQuery = otQuery.lte('check_in', endDate);
      lQuery = lQuery.lte('end_date', endDate);
      alQuery = alQuery.lte('end_date', endDate);
      pQuery = pQuery.lte('end_date', endDate);
      mlQuery = mlQuery.lte('end_date', endDate);
    }

    const [accounts, attendances, overtimes, leaves, annualLeaves, permissions, maternityLeaves] = await Promise.all([
      accQuery,
      attQuery,
      otQuery,
      lQuery,
      alQuery,
      pQuery,
      mlQuery,
    ]);

    return {
      accounts: accounts.data || [],
      attendances: attendances.data || [],
      overtimes: overtimes.data || [],
      leaves: leaves.data || [],
      annualLeaves: annualLeaves.data || [],
      permissions: permissions.data || [],
      maternityLeaves: maternityLeaves.data || []
    };
  },

  async getFinanceReport(startDate?: string, endDate?: string) {
    let pQuery = supabase.from('payroll_items').select('*, account:accounts(full_name, internal_nik)');
    let otQuery = supabase.from('overtimes').select('*, account:accounts(full_name, internal_nik)');
    let rQuery = supabase.from('reimbursements').select('*, account:accounts(full_name, internal_nik)');
    let cQuery = supabase.from('compensations').select('*, account:accounts(full_name, internal_nik)');

    if (startDate) {
      pQuery = pQuery.gte('created_at', startDate);
      otQuery = otQuery.gte('check_in', startDate);
      rQuery = rQuery.gte('created_at', startDate);
      cQuery = cQuery.gte('created_at', startDate);
    }
    if (endDate) {
      pQuery = pQuery.lte('created_at', endDate);
      otQuery = otQuery.lte('check_in', endDate);
      rQuery = rQuery.lte('created_at', endDate);
      cQuery = cQuery.lte('created_at', endDate);
    }

    const [payrollItems, overtimes, reimbursements, compensations] = await Promise.all([
      pQuery,
      otQuery,
      rQuery,
      cQuery,
    ]);

    return {
      payrollItems: payrollItems.data || [],
      overtimes: overtimes.data || [],
      reimbursements: reimbursements.data || [],
      compensations: compensations.data || []
    };
  },

  async getLeaveReport(startDate?: string, endDate?: string): Promise<LeaveSummary[]> {
    const [
      { data: accounts },
      { data: leaves },
      { data: annualLeaves },
      { data: maternityLeaves },
      { data: permissions }
    ] = await Promise.all([
      supabase.from('accounts').select('id, full_name, internal_nik, leave_quota'),
      supabase.from('leave_requests').select('*').eq('status', 'approved'),
      supabase.from('annual_leave_requests').select('*').eq('status', 'approved'),
      supabase.from('maternity_leave_requests').select('*').eq('status', 'approved'),
      supabase.from('permission_requests').select('*').eq('status', 'approved')
    ]);

    if (!accounts) return [];
    
    return accounts.map(acc => {
      const usedQuota = (leaves || []).filter(l => l.account_id === acc.id).length + 
                        (annualLeaves || []).filter(l => l.account_id === acc.id).length;
      const maternityUsed = (maternityLeaves || []).filter(m => m.account_id === acc.id).length;
      const permissionCount = (permissions || []).filter(p => p.account_id === acc.id).length;

      return {
        accountId: acc.id,
        fullName: acc.full_name,
        nik: acc.internal_nik,
        totalQuota: acc.leave_quota || 12,
        usedQuota,
        remainingQuota: (acc.leave_quota || 12) - usedQuota,
        carryOverQuota: 0,
        maternityQuota: 90,
        maternityUsed,
        permissionCount
      };
    });
  },

  async getOvertimeReport(startDate?: string, endDate?: string): Promise<OvertimeSummary[]> {
    const [
      { data: accounts },
      { data: overtimes }
    ] = await Promise.all([
      supabase.from('accounts').select('id, full_name, internal_nik'),
      supabase.from('overtimes').select('*').gte('check_in', startDate || '2000-01-01').lte('check_in', endDate || '2100-01-01')
    ]);

    if (!accounts) return [];

    return accounts.map(acc => {
      const accOvertimes = (overtimes || []).filter(o => o.account_id === acc.id);
      const totalMinutes = accOvertimes.reduce((sum, o) => {
        if (!o.check_in || !o.check_out) return sum;
        const diff = new Date(o.check_out).getTime() - new Date(o.check_in).getTime();
        return sum + Math.floor(diff / (1000 * 60));
      }, 0);

      return {
        accountId: acc.id,
        fullName: acc.full_name,
        nik: acc.internal_nik,
        totalOvertimeMinutes: totalMinutes,
        totalOvertimeHours: Number((totalMinutes / 60).toFixed(2)),
        overtimeCount: accOvertimes.length,
        estimatedCost: 0 // Cost calculation would need payroll settings
      };
    });
  }
};
