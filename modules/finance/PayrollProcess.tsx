import React, { useState, useEffect, useMemo } from 'react';
import { ArrowLeft, Save, Calculator, Search, Check, AlertCircle, Info, UserCheck, DollarSign, Calendar, Clock, Trash2, Users } from 'lucide-react';
import { financeService } from '../../services/financeService';
import { accountService } from '../../services/accountService';
import { presenceService } from '../../services/presenceService';
import { overtimeService } from '../../services/overtimeService';
import { submissionService } from '../../services/submissionService';
import { scheduleService } from '../../services/scheduleService';
import { authService } from '../../services/authService';
import { Account, SalaryScheme, SalaryAssignmentExtended, SalaryAdjustment, Payroll, PayrollItem, PayrollSettings } from '../../types';
import Swal from 'sweetalert2';

interface PayrollProcessProps {
  payroll?: Payroll | null;
  onBack: () => void;
}

const PayrollProcess: React.FC<PayrollProcessProps> = ({ payroll, onBack }) => {
  const [loading, setLoading] = useState(false);
  const [calculating, setCalculating] = useState(false);
  const [accounts, setAccounts] = useState<Account[]>([]);
  const [assignments, setAssignments] = useState<SalaryAssignmentExtended[]>([]);
  const [verifiers, setVerifiers] = useState<Account[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedAccountIds, setSelectedAccountIds] = useState<string[]>([]);
  
  const [config, setConfig] = useState({
    month: payroll?.month || new Date().getMonth() + 1,
    year: payroll?.year || new Date().getFullYear(),
    start_date: payroll?.start_date || new Date(new Date().getFullYear(), new Date().getMonth() - 1, 21).toISOString().split('T')[0],
    end_date: payroll?.end_date || new Date(new Date().getFullYear(), new Date().getMonth(), 20).toISOString().split('T')[0],
    verifier_id: payroll?.verifier_id || ''
  });

  const [payrollItems, setPayrollItems] = useState<Partial<PayrollItem>[]>([]);

  useEffect(() => {
    const fetchData = async () => {
      setLoading(true);
      try {
        const [accs, asgs, vfrs] = await Promise.all([
          accountService.getAll(),
          financeService.getAssignments(),
          accountService.getAll()
        ]);
        setAccounts(accs);
        setAssignments(asgs);
        setVerifiers(vfrs);

        if (payroll) {
          const items = await financeService.getPayrollItems(payroll.id);
          setPayrollItems(items);
          setSelectedAccountIds(items.map(i => i.account_id));
        }
      } catch (error) {
        console.error('Error fetching data:', error);
        Swal.fire('Gagal!', 'Terjadi kesalahan saat memuat data payroll.', 'error');
      } finally {
        setLoading(false);
      }
    };
    fetchData();
  }, [payroll]);

  const handleCalculate = async () => {
    if (selectedAccountIds.length === 0) {
      Swal.fire('Peringatan', 'Silakan pilih minimal satu karyawan.', 'warning');
      return;
    }

    setCalculating(true);
    try {
      const [attendances, adjustments, earlySalaryRequests, overtimes, submissions, allSchedules] = await Promise.all([
        presenceService.getAttendanceByRange(config.start_date, config.end_date),
        financeService.getSalaryAdjustments({ month: config.month, year: config.year }),
        financeService.getEarlySalaryRequests({ month: config.month, year: config.year }),
        overtimeService.getOvertimeByRange(config.start_date, config.end_date),
        submissionService.getSubmissionsByRange(config.start_date, config.end_date),
        scheduleService.getAll()
      ]);

      const newItems: Partial<PayrollItem>[] = selectedAccountIds.map(accountId => {
        const account = accounts.find(a => a.id === accountId);
        const assignment = assignments.find(as => as.account_id === accountId);
        const scheme = assignment?.scheme;
        const schedule = allSchedules.find(s => s.id === account?.schedule_id);
        
        const userAttendances = attendances.filter(at => at.account_id === accountId);
        const userAdjustments = adjustments.filter(ad => ad.account_id === accountId);
        const userEarlySalary = earlySalaryRequests
          .filter(es => es.account_id === accountId && (es.status === 'Approved' || es.status === 'Paid'))
          .reduce((sum, es) => sum + es.amount, 0);
        
        const userEarlySalaryNotes = earlySalaryRequests
          .filter(es => es.account_id === accountId && (es.status === 'Approved' || es.status === 'Paid'))
          .map(es => `Kasbon: Rp ${es.amount.toLocaleString('id-ID')}`)
          .join(', ');

        const userOvertimes = overtimes.filter(ot => ot.account_id === accountId);
        const userSubmissions = submissions.filter(s => s.account_id === accountId && s.status === 'Disetujui');

        // Calculate Overtime Pay
        const totalOvertimeMinutes = userOvertimes.reduce((sum, ot) => sum + (ot.duration_minutes || 0), 0);
        const totalOvertimeHours = totalOvertimeMinutes / 60;
        const overtimePay = totalOvertimeHours * (scheme?.overtime_rate_per_hour || 0);
        const overtimePayNotes = totalOvertimeHours > 0 ? `${totalOvertimeHours.toFixed(2)} Jam x Rp ${(scheme?.overtime_rate_per_hour || 0).toLocaleString('id-ID')}` : '';

        // Calculate Attendance & Absences
        const startDate = new Date(config.start_date);
        const endDate = new Date(config.end_date);
        const totalDaysInRange = Math.ceil((endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24)) + 1;
        
        const attendanceDates = new Set(userAttendances.map(at => new Date(at.created_at || '').toISOString().split('T')[0]));
        
        // Submissions unique dates
        const submissionDates = new Set<string>();
        userSubmissions.forEach(s => {
          if (['Cuti', 'Izin', 'Libur Mandiri', 'Cuti Melahirkan', 'Cuti Tahunan'].includes(s.type)) {
            const sData = s.submission_data;
            if (sData.start_date && sData.end_date) {
              let current = new Date(sData.start_date);
              const end = new Date(sData.end_date);
              while (current <= end) {
                const dateStr = current.toISOString().split('T')[0];
                if (dateStr >= config.start_date && dateStr <= config.end_date) {
                  submissionDates.add(dateStr);
                }
                current.setDate(current.getDate() + 1);
              }
            }
          }
        });
        
        // Determine working days and absences
        let absentDays = 0;
        let presentDays = attendanceDates.size;
        let validAbsenceDays = 0;

        const userLocationId = account?.location_id || (account as any)?.location_id;
        const specialHolidays = allSchedules.filter(s => 
          s.type === 3 && 
          s.location_ids?.includes(userLocationId) &&
          !s.excluded_account_ids?.includes(accountId)
        );

        // Iterate through each day in range to check for absences
        let current = new Date(startDate);
        while (current <= endDate) {
          const dateStr = current.toISOString().split('T')[0];
          const dayOfWeek = current.getDay();
          const rule = schedule?.rules?.find(r => r.day_of_week === dayOfWeek);
          
          // 1. Check Special Holiday (Type 3)
          const isSpecialHoliday = specialHolidays.some(h => 
            h.start_date && h.end_date && dateStr >= h.start_date && dateStr <= h.end_date
          );

          // 2. Check Regular Working Day
          const isFlexible = schedule?.id === 'FLEKSIBEL' || account?.schedule_type === 'Fleksibel' || account?.schedule_type === 'Shift Dinamis';
          
          let isWorkingDay = false;
          if (!isSpecialHoliday) {
            if (isFlexible) {
              isWorkingDay = true;
            } else if (rule) {
              isWorkingDay = !rule.is_holiday;
            }
          }
          
          const isPresent = attendanceDates.has(dateStr);
          
          // 3. Check Submission
          const submission = userSubmissions.find(s => {
            const sData = s.submission_data || {};
            const start = sData.start_date || sData.date;
            const end = sData.end_date || sData.date;
            if (!start) return false;
            return dateStr >= start && dateStr <= end;
          });

          if (isWorkingDay && !isPresent) {
            if (submission) {
              validAbsenceDays++;
              // Izin/Libur Mandiri tetap ngga potong gaji secara default (Manual by Admin)
            } else {
              // Alpha (No presence, no submission, is working day)
              absentDays++;
            }
          }
          current.setDate(current.getDate() + 1);
        }

        // Basic Salary Calculation
        let basicSalary = scheme?.basic_salary || 0;
        let basicSalaryNotes = '';
        if (scheme?.type === 'Harian') {
          basicSalary = (scheme?.basic_salary || 0) * presentDays;
          basicSalaryNotes = `${presentDays} Hari x Rp ${(scheme?.basic_salary || 0).toLocaleString('id-ID')}`;
        }

        // Deductions from Attendance
        let lateMins = 0;
        let earlyMins = 0;
        let noClockOutDays = 0;
        userAttendances.forEach(at => {
          if (at.status_in === 'Terlambat') lateMins += at.late_minutes || 0;
          if (at.status_out === 'Pulang Cepat') earlyMins += at.early_departure_minutes || 0;
          if (at.check_in && !at.check_out) noClockOutDays += 1;
        });

        const lateDeduction = (scheme?.late_deduction_per_minute || 0) * lateMins;
        const earlyDeduction = (scheme?.early_leave_deduction_per_minute || 0) * earlyMins;
        const noClockOutDeduction = (scheme?.no_clock_out_deduction_per_day || 0) * noClockOutDays;
        const absentDeduction = (scheme?.absent_deduction_per_day || 0) * absentDays;

        // Custom Adjustments
        const otherAdditions = userAdjustments
          .filter(ad => ad.type === 'Addition')
          .reduce((sum, ad) => sum + ad.amount, 0);
        const otherAdditionsNotes = userAdjustments
          .filter(ad => ad.type === 'Addition')
          .map(ad => ad.description)
          .join(', ');

        const otherDeductions = userAdjustments
          .filter(ad => ad.type === 'Deduction')
          .reduce((sum, ad) => sum + ad.amount, 0) + userEarlySalary;
        const otherDeductionsNotes = [
          ...userAdjustments.filter(ad => ad.type === 'Deduction').map(ad => ad.description),
          userEarlySalaryNotes
        ].filter(Boolean).join(', ');

        const totalIncome = basicSalary + (scheme?.position_allowance || 0) + (scheme?.placement_allowance || 0) + (scheme?.other_allowance || 0) + otherAdditions + overtimePay;
        const totalDeduction = lateDeduction + earlyDeduction + noClockOutDeduction + absentDeduction + otherDeductions;
        const takeHomePay = Math.max(0, totalIncome - totalDeduction);

        return {
          account_id: accountId,
          salary_type: scheme?.type || 'N/A',
          basic_salary: basicSalary,
          basic_salary_notes: basicSalaryNotes || (!scheme ? 'Skema Belum Diatur' : ''),
          position_allowance: scheme?.position_allowance || 0,
          placement_allowance: scheme?.placement_allowance || 0,
          other_allowance: scheme?.other_allowance || 0,
          overtime_pay: overtimePay,
          overtime_pay_notes: overtimePayNotes,
          other_additions: otherAdditions,
          other_additions_notes: otherAdditionsNotes,
          late_deduction: lateDeduction,
          late_deduction_notes: `${lateMins} Menit`,
          early_leave_deduction: earlyDeduction,
          early_leave_deduction_notes: `${earlyMins} Menit`,
          absent_deduction: absentDeduction + noClockOutDeduction,
          absent_deduction_notes: `${absentDays} Hari Potong (Alpha), ${noClockOutDays} Hari No-Out, ${validAbsenceDays} Hari Total Izin/Cuti`,
          other_deductions: otherDeductions,
          other_deductions_notes: otherDeductionsNotes,
          bpjs_kesehatan: 0,
          bpjs_ketenagakerjaan: 0,
          pph21: 0,
          total_income: totalIncome,
          total_deduction: totalDeduction,
          take_home_pay: takeHomePay,
          account: {
            full_name: account?.full_name || '',
            internal_nik: account?.internal_nik || '',
            position: account?.position || '',
            grade: account?.grade || '',
            department: (account as any)?.department || '-',
            location: account?.location?.name || (account as any)?.location || '-'
          }
        };
      });

      setPayrollItems(newItems);
      Swal.fire('Berhasil!', 'Kalkulasi otomatis selesai. Silakan review dan sesuaikan jika perlu.', 'success');
    } catch (error) {
      console.error('Error calculating payroll:', error);
      Swal.fire('Gagal!', 'Terjadi kesalahan saat kalkulasi.', 'error');
    } finally {
      setCalculating(false);
    }
  };

  const handleManualOverride = (index: number, updates: Partial<PayrollItem>) => {
    setPayrollItems(prev => {
      const updatedItems = [...prev];
      const item = { ...updatedItems[index], ...updates };
      
      // Recalculate totals
      const income = (item.basic_salary || 0) + (item.position_allowance || 0) + (item.placement_allowance || 0) + (item.other_allowance || 0) + (item.overtime_pay || 0) + (item.other_additions || 0);
      const deduction = (item.late_deduction || 0) + (item.early_leave_deduction || 0) + (item.absent_deduction || 0) + (item.other_deductions || 0) + (item.bpjs_kesehatan || 0) + (item.bpjs_ketenagakerjaan || 0) + (item.pph21 || 0);
      
      item.total_income = income;
      item.total_deduction = deduction;
      item.take_home_pay = Math.max(0, income - deduction);
      
      updatedItems[index] = item;
      return updatedItems;
    });
  };

  const handleSave = async (status: Payroll['status'] = 'Draft') => {
    if (payrollItems.length === 0) {
      Swal.fire('Peringatan', 'Belum ada data payroll untuk disimpan.', 'warning');
      return;
    }

    if (status === 'Pending' && !config.verifier_id) {
      Swal.fire('Peringatan', 'Silakan pilih verifikator sebelum mengajukan.', 'warning');
      return;
    }

    setLoading(true);
    try {
      let payrollId = payroll?.id;
      const currentUser = authService.getCurrentUser();
      
      if (!payrollId) {
        const newPayroll = await financeService.createPayroll({
          month: config.month,
          year: config.year,
          start_date: config.start_date,
          end_date: config.end_date,
          status: status,
          verifier_id: config.verifier_id || undefined,
          created_by: currentUser?.id,
          updated_by: currentUser?.id
        });
        payrollId = newPayroll.id;
      } else {
        await financeService.updatePayrollStatus(payrollId, status, config.verifier_id || undefined);
        // Also update updated_by
        await financeService.updatePayroll(payrollId, { updated_by: currentUser?.id });
      }

      const itemsToSave = payrollItems.map(item => ({
        payroll_id: payrollId!,
        account_id: item.account_id!,
        salary_type: item.salary_type!,
        basic_salary: item.basic_salary!,
        basic_salary_notes: item.basic_salary_notes,
        position_allowance: item.position_allowance!,
        position_allowance_notes: item.position_allowance_notes,
        placement_allowance: item.placement_allowance!,
        placement_allowance_notes: item.placement_allowance_notes,
        other_allowance: item.other_allowance!,
        other_allowance_notes: item.other_allowance_notes,
        overtime_pay: item.overtime_pay!,
        overtime_pay_notes: item.overtime_pay_notes,
        other_additions: item.other_additions!,
        other_additions_notes: item.other_additions_notes,
        late_deduction: item.late_deduction!,
        late_deduction_notes: item.late_deduction_notes,
        early_leave_deduction: item.early_leave_deduction!,
        early_leave_deduction_notes: item.early_leave_deduction_notes,
        absent_deduction: item.absent_deduction!,
        absent_deduction_notes: item.absent_deduction_notes,
        other_deductions: item.other_deductions!,
        other_deductions_notes: item.other_deductions_notes,
        bpjs_kesehatan: item.bpjs_kesehatan!,
        bpjs_ketenagakerjaan: item.bpjs_ketenagakerjaan!,
        pph21: item.pph21!,
        total_income: item.total_income!,
        total_deduction: item.total_deduction!,
        take_home_pay: item.take_home_pay!
      }));

      await financeService.upsertPayrollItems(itemsToSave);
      
      Swal.fire('Berhasil!', status === 'Pending' ? 'Payroll berhasil diajukan untuk verifikasi.' : 'Data payroll berhasil disimpan.', 'success');
      onBack();
    } catch (error) {
      console.error('Error saving payroll:', error);
      Swal.fire('Gagal!', 'Terjadi kesalahan saat menyimpan data.', 'error');
    } finally {
      setLoading(false);
    }
  };

  const filteredAccounts = accounts.filter(acc => 
    acc.email !== 'fakhriwildana.work@gmail.com' && (
      acc.full_name.toLowerCase().includes(searchQuery.toLowerCase()) ||
      acc.internal_nik.toLowerCase().includes(searchQuery.toLowerCase())
    )
  );

  const toggleAccount = (id: string) => {
    setSelectedAccountIds(prev => 
      prev.includes(id) ? prev.filter(i => i !== id) : [...prev, id]
    );
  };

  const isLocked = payroll && (payroll.status === 'Approved' || payroll.status === 'Paid');

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-4">
          <button
            onClick={onBack}
            className="p-2 hover:bg-gray-100 rounded-full transition-colors text-gray-500"
          >
            <ArrowLeft size={20} />
          </button>
          <div>
            <h2 className="text-2xl font-bold text-gray-800">{payroll ? 'Proses Payroll' : 'Buat Payroll Baru'}</h2>
            <p className="text-sm text-gray-500">Hitung gaji karyawan berdasarkan presensi dan skema gaji.</p>
          </div>
        </div>
        {payroll && (
          <div className="text-right">
            <div className="text-[10px] text-gray-400 uppercase tracking-widest">Audit Trail</div>
            <div className="text-xs text-gray-500">
              Dibuat oleh: <span className="font-bold text-gray-700">{payroll.creator?.full_name || '-'}</span>
            </div>
            {payroll.updater && (
              <div className="text-xs text-gray-500">
                Diubah oleh: <span className="font-bold text-gray-700">{payroll.updater.full_name}</span>
              </div>
            )}
          </div>
        )}
      </div>

      <div className="space-y-6">
        {/* Horizontal Config Bar */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          {/* Periode & Analisa */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 space-y-4">
            <div className="flex items-center gap-2 text-[#006E62] mb-2">
              <Calendar size={18} />
              <h3 className="font-bold text-sm uppercase tracking-wider">Periode & Analisa</h3>
            </div>

            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Bulan</label>
                <select
                  value={config.month}
                  onChange={(e) => setConfig({ ...config, month: Number(e.target.value) })}
                  disabled={!!payroll}
                  className="w-full px-3 py-2 bg-gray-50 border-none rounded-lg text-sm font-medium focus:ring-2 focus:ring-[#006E62]"
                >
                  {Array.from({ length: 12 }, (_, i) => (
                    <option key={i + 1} value={i + 1}>
                      {new Date(0, i).toLocaleString('id-ID', { month: 'long' })}
                    </option>
                  ))}
                </select>
              </div>

              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Tahun</label>
                <select
                  value={config.year}
                  onChange={(e) => setConfig({ ...config, year: Number(e.target.value) })}
                  disabled={!!payroll}
                  className="w-full px-3 py-2 bg-gray-50 border-none rounded-lg text-sm font-medium focus:ring-2 focus:ring-[#006E62]"
                >
                  {Array.from({ length: 5 }, (_, i) => (
                    <option key={i} value={new Date().getFullYear() - 2 + i}>
                      {new Date().getFullYear() - 2 + i}
                    </option>
                  ))}
                </select>
              </div>

              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Mulai</label>
                <input
                  type="date"
                  value={config.start_date}
                  onChange={(e) => setConfig({ ...config, start_date: e.target.value })}
                  disabled={!!payroll}
                  className="w-full px-3 py-2 bg-gray-50 border-none rounded-lg text-xs font-medium focus:ring-2 focus:ring-[#006E62]"
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Selesai</label>
                <input
                  type="date"
                  value={config.end_date}
                  onChange={(e) => setConfig({ ...config, end_date: e.target.value })}
                  disabled={!!payroll}
                  className="w-full px-3 py-2 bg-gray-50 border-none rounded-lg text-xs font-medium focus:ring-2 focus:ring-[#006E62]"
                />
              </div>
            </div>

            <div className="pt-2">
              <button
                onClick={handleCalculate}
                disabled={calculating || isLocked}
                className="w-full flex items-center justify-center gap-2 py-2.5 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 transition-all text-xs font-bold uppercase tracking-wider shadow-sm disabled:opacity-50"
              >
                {calculating ? <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></div> : <Calculator size={16} />}
                Kalkulasi Otomatis
              </button>
            </div>
          </div>

          {/* Pilih Karyawan */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 flex flex-col h-full">
            <div className="flex items-center gap-2 text-[#006E62] mb-4">
              <Users size={18} />
              <h3 className="font-bold text-sm uppercase tracking-wider">Pilih Karyawan</h3>
            </div>
            {payroll ? (
              <div className="flex-1 flex items-center justify-center text-gray-400 text-xs italic">
                Data karyawan sudah terkunci untuk payroll ini.
              </div>
            ) : (
              <>
                <div className="relative mb-4">
                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
                  <input
                    type="text"
                    placeholder="Cari..."
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="w-full pl-10 pr-4 py-2 bg-gray-50 border-none rounded-lg text-xs focus:ring-2 focus:ring-[#006E62]"
                  />
                </div>
                <div className="flex-1 overflow-y-auto space-y-2 pr-2 custom-scrollbar max-h-[120px]">
                  <button
                    onClick={() => {
                      if (selectedAccountIds.length === filteredAccounts.length) setSelectedAccountIds([]);
                      else setSelectedAccountIds(filteredAccounts.map(a => a.id));
                    }}
                    className="w-full text-left px-3 py-2 text-[10px] font-bold text-[#006E62] uppercase tracking-widest hover:bg-emerald-50 rounded-lg"
                  >
                    {selectedAccountIds.length === filteredAccounts.length ? 'Batal Pilih Semua' : 'Pilih Semua'}
                  </button>
                  {filteredAccounts.map((acc) => (
                    <button
                      key={acc.id}
                      onClick={() => toggleAccount(acc.id)}
                      className={`w-full flex items-center justify-between p-2 rounded-lg border transition-all text-left ${
                        selectedAccountIds.includes(acc.id)
                          ? 'border-[#006E62] bg-emerald-50'
                          : 'border-gray-100 hover:border-gray-200'
                      }`}
                    >
                      <div className="truncate pr-2">
                        <div className="text-[10px] font-bold text-gray-800 truncate">{acc.full_name}</div>
                        <div className="text-[8px] text-gray-500 uppercase tracking-wider">{acc.internal_nik}</div>
                      </div>
                      {selectedAccountIds.includes(acc.id) && (
                        <div className="bg-[#006E62] text-white p-0.5 rounded-full shrink-0">
                          <Check size={8} />
                        </div>
                      )}
                    </button>
                  ))}
                </div>
              </>
            )}
          </div>

          {/* Verifikasi */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 space-y-4">
            <div className="flex items-center gap-2 text-[#006E62] mb-2">
              <UserCheck size={18} />
              <h3 className="font-bold text-sm uppercase tracking-wider">Verifikasi</h3>
            </div>

            <div className="space-y-3">
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Pilih Verifikator</label>
                <select
                  value={config.verifier_id}
                  onChange={(e) => setConfig({ ...config, verifier_id: e.target.value })}
                  disabled={isLocked}
                  className="w-full px-3 py-2 bg-gray-50 border-none rounded-lg text-sm font-medium focus:ring-2 focus:ring-[#006E62]"
                >
                  <option value="">-- Pilih User --</option>
                  {verifiers.map(v => (
                    <option key={v.id} value={v.id}>{v.full_name}</option>
                  ))}
                </select>
              </div>

              <div className="grid grid-cols-2 gap-2">
                <button
                  onClick={() => handleSave('Draft')}
                  disabled={loading || isLocked}
                  className="flex items-center justify-center gap-2 py-2.5 bg-white border border-gray-200 text-gray-600 rounded-lg hover:bg-gray-50 transition-all text-[10px] font-bold uppercase tracking-wider shadow-sm disabled:opacity-50"
                >
                  <Save size={14} />
                  Draft
                </button>
                <button
                  onClick={() => handleSave('Pending')}
                  disabled={loading || isLocked}
                  className="flex items-center justify-center gap-2 py-2.5 bg-[#006E62] text-white rounded-lg hover:bg-[#005a50] transition-all text-[10px] font-bold uppercase tracking-wider shadow-md disabled:opacity-50"
                >
                  <Check size={14} />
                  Ajukan
                </button>
              </div>
            </div>
          </div>
        </div>

        {/* Main Table - Full Width */}
        <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
            <div className="p-4 border-b border-gray-50 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Info size={16} className="text-blue-500" />
                <span className="text-xs font-medium text-gray-500 italic">Klik pada nominal untuk melakukan koreksi manual.</span>
              </div>
              <div className="text-xs font-bold text-gray-800 uppercase tracking-widest">
                Total: {payrollItems.length} Karyawan
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-gray-50 text-gray-500 text-[9px] uppercase tracking-widest font-bold border-b border-gray-100">
                    <th className="px-4 py-3 sticky left-0 bg-gray-50 z-10">Karyawan</th>
                    <th className="px-4 py-3">Gaji Pokok</th>
                    <th className="px-4 py-3">Tunj. Jabatan</th>
                    <th className="px-4 py-3">Tunj. Penempatan</th>
                    <th className="px-4 py-3">Tunj. Lain</th>
                    <th className="px-4 py-3 text-emerald-600">Upah Lembur</th>
                    <th className="px-4 py-3">Tambahan Lain</th>
                    <th className="px-4 py-3 text-rose-600">Pot. Telat</th>
                    <th className="px-4 py-3 text-rose-600">Pot. Pulang Cepat</th>
                    <th className="px-4 py-3 text-rose-600">Pot. Absen</th>
                    <th className="px-4 py-3 text-rose-600">Pot. Lain</th>
                    <th className="px-4 py-3 text-rose-600">BPJS Kes</th>
                    <th className="px-4 py-3 text-rose-600">BPJS TK</th>
                    <th className="px-4 py-3 text-rose-600">PPh21</th>
                    <th className="px-4 py-3 bg-emerald-50 text-[#006E62] sticky right-0 z-10">Total THP</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-50">
                  {payrollItems.length === 0 ? (
                    <tr>
                      <td colSpan={14} className="px-6 py-24 text-center">
                        <div className="flex flex-col items-center gap-2 text-gray-400">
                          <Calculator size={48} strokeWidth={1} />
                          <p className="text-sm font-medium">Belum ada data. Silakan pilih karyawan dan klik Kalkulasi Otomatis.</p>
                        </div>
                      </td>
                    </tr>
                  ) : (
                    payrollItems.map((item, idx) => (
                      <tr key={item.account_id} className="hover:bg-gray-50 transition-colors text-xs group">
                        <td className="px-4 py-3 sticky left-0 bg-white group-hover:bg-gray-50 z-10 border-r border-gray-50">
                          <div className="font-bold text-gray-800 whitespace-nowrap">{item.account?.full_name}</div>
                          <div className="text-[9px] text-gray-400 uppercase tracking-widest">
                            {item.account?.internal_nik} • {item.salary_type}
                            {item.account?.position && ` • ${item.account.position}`}
                          </div>
                        </td>
                        
                        {/* Editable Cells */}
                        {[
                          { field: 'basic_salary', value: item.basic_salary, notes: item.basic_salary_notes },
                          { field: 'position_allowance', value: item.position_allowance, notes: item.position_allowance_notes },
                          { field: 'placement_allowance', value: item.placement_allowance, notes: item.placement_allowance_notes },
                          { field: 'other_allowance', value: item.other_allowance, notes: item.other_allowance_notes },
                          { field: 'overtime_pay', value: item.overtime_pay, notes: item.overtime_pay_notes, color: 'text-emerald-600' },
                          { field: 'other_additions', value: item.other_additions, notes: item.other_additions_notes },
                          { field: 'late_deduction', value: item.late_deduction, notes: item.late_deduction_notes, color: 'text-rose-600' },
                          { field: 'early_leave_deduction', value: item.early_leave_deduction, notes: item.early_leave_deduction_notes, color: 'text-rose-600' },
                          { field: 'absent_deduction', value: item.absent_deduction, notes: item.absent_deduction_notes, color: 'text-rose-600' },
                          { field: 'other_deductions', value: item.other_deductions, notes: item.other_deductions_notes, color: 'text-rose-600' },
                          { field: 'bpjs_kesehatan', value: item.bpjs_kesehatan, color: 'text-rose-600' },
                          { field: 'bpjs_ketenagakerjaan', value: item.bpjs_ketenagakerjaan, color: 'text-rose-600' },
                          { field: 'pph21', value: item.pph21, color: 'text-rose-600' },
                        ].map((cell) => (
                          <td 
                            key={cell.field} 
                            className={`px-4 py-3 cursor-pointer hover:bg-emerald-50 transition-all font-mono ${cell.color || 'text-gray-700'}`}
                            onClick={async () => {
                              if (isLocked) return;
                              const { value: formValues } = await Swal.fire({
                                title: 'Koreksi Komponen',
                                html: `
                                  <div class="space-y-4 text-left">
                                    <div>
                                      <label class="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Nominal Baru</label>
                                      <input id="swal-input1" class="swal2-input w-full" type="number" value="${cell.value}">
                                    </div>
                                    <div>
                                      <label class="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Keterangan Koreksi</label>
                                      <textarea id="swal-input2" class="swal2-textarea w-full" placeholder="Alasan koreksi...">${cell.notes || ''}</textarea>
                                    </div>
                                  </div>
                                `,
                                focusConfirm: false,
                                showCancelButton: true,
                                preConfirm: () => {
                                  return [
                                    (document.getElementById('swal-input1') as HTMLInputElement).value,
                                    (document.getElementById('swal-input2') as HTMLTextAreaElement).value
                                  ];
                                }
                              });

                              if (formValues) {
                                handleManualOverride(idx, {
                                  [cell.field]: Number(formValues[0]),
                                  [`${cell.field}_notes`]: formValues[1]
                                });
                              }
                            }}
                          >
                            <div className="flex flex-col">
                              <span>{cell.value?.toLocaleString('id-ID')}</span>
                              {cell.notes && <span className="text-[8px] text-gray-400 truncate max-w-[80px]">{cell.notes}</span>}
                            </div>
                          </td>
                        ))}

                        <td className="px-4 py-3 bg-emerald-50 text-[#006E62] font-bold font-mono sticky right-0 group-hover:bg-emerald-100 z-10 border-l border-emerald-100">
                          Rp {item.take_home_pay?.toLocaleString('id-ID')}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    );
  };

export default PayrollProcess;
