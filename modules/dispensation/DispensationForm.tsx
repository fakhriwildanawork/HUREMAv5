import React, { useState, useEffect } from 'react';
import { X, Save, AlertCircle, Calendar, Clock, FileText, Upload, CheckCircle2, Loader2, Info, ClipboardList } from 'lucide-react';
import { dispensationService } from '../../services/dispensationService';
import { presenceService } from '../../services/presenceService';
import { settingsService } from '../../services/settingsService';
import { scheduleService } from '../../services/scheduleService';
import { googleDriveService } from '../../services/googleDriveService';
import { authService } from '../../services/authService';
import { leaveService } from '../../services/leaveService';
import { permissionService } from '../../services/permissionService';
import { maternityLeaveService } from '../../services/maternityLeaveService';
import { accountService } from '../../services/accountService';
import { DispensationRequest, DispensationIssueType, DispensationIssue, Attendance, ScheduleRule, Account, Schedule } from '../../types';
import Swal from 'sweetalert2';

interface DispensationFormProps {
  onClose: () => void;
  onSuccess: () => void;
  editData: DispensationRequest | null;
}

interface AttendanceProblem {
  date: string;
  presence_id: string | null;
  issues: DispensationIssueType[];
  details: string;
}

const DispensationForm: React.FC<DispensationFormProps> = ({ onClose, onSuccess, editData }) => {
  const user = authService.getCurrentUser();
  const [isLoading, setIsLoading] = useState(false);
  const [isDetecting, setIsDetecting] = useState(false);
  const [problems, setProblems] = useState<AttendanceProblem[]>([]);
  const [selectedProblem, setSelectedProblem] = useState<AttendanceProblem | null>(null);
  const [selectedIssues, setSelectedIssues] = useState<DispensationIssueType[]>([]);
  const [reason, setReason] = useState(editData?.reason || '');
  const [file, setFile] = useState<File | null>(null);
  const [isUploading, setIsUploading] = useState(false);

  useEffect(() => {
    if (!editData) {
      detectProblems();
    } else {
      // If editing, we just show the current data
      setReason(editData.reason);
      setSelectedIssues(editData.issues.map(i => i.type));
      // We still need to show the date info
      setSelectedProblem({
        date: editData.date,
        presence_id: editData.presence_id,
        issues: editData.issues.map(i => i.type),
        details: 'Mode Edit Pengajuan'
      });
    }
  }, [editData]);

  const detectProblems = async () => {
    try {
      setIsDetecting(true);
      
      // Fetch full account details to get location_id and schedule_id
      const fullAccount = await accountService.getById(user!.id);
      if (!fullAccount) throw new Error("Data akun tidak ditemukan");

      const settings = await settingsService.getAll();
      const windowDays = settings.find(s => s.key === 'dispensation_window_days')?.value || 7;
      
      // Gunakan waktu lokal Jakarta (WIB)
      const now = new Date();
      const startDate = new Date(now);
      startDate.setDate(now.getDate() - (windowDays - 1));
      
      const startDateStr = startDate.toISOString().split('T')[0];
      const endDateStr = now.toISOString().split('T')[0];

      // Fetch all relevant data for the window in parallel
      const [
        attendance,
        annualLeaves,
        mandatoryLeaves,
        permissions,
        maternityLeaves,
        existingRequests,
        specialHolidays
      ] = await Promise.all([
        presenceService.getAttendanceByRange(startDateStr, endDateStr, user!.id),
        leaveService.getAnnualByRange(user!.id, startDateStr, endDateStr),
        leaveService.getMandatoryByRange(user!.id, startDateStr, endDateStr),
        permissionService.getByRange(user!.id, startDateStr, endDateStr),
        maternityLeaveService.getByRange(user!.id, startDateStr, endDateStr),
        dispensationService.getByRange(user!.id, startDateStr, endDateStr),
        scheduleService.getSpecialHolidaysByRange(fullAccount.location_id, startDateStr, endDateStr)
      ]);

      const filteredAttendance = attendance; // Already filtered by range

      // Fetch schedule rules
      const scheduleId = fullAccount.schedule_id;
      let rules: ScheduleRule[] = [];
      if (scheduleId) {
        const schedule = await scheduleService.getById(scheduleId);
        rules = schedule.rules || [];
      }

      // Detect issues for each day in window
      const detected: AttendanceProblem[] = [];
      for (let i = 0; i < windowDays; i++) {
        const d = new Date(now);
        d.setDate(now.getDate() - i);
        const dStr = d.toISOString().split('T')[0];
        
        // Skip if it's today (cannot request dispensation for today yet)
        if (dStr === now.toISOString().split('T')[0]) continue;

        // Check if this date is a Special Holiday (Type 3)
        const isSpecialHoliday = specialHolidays.some(h => {
          if (!h.start_date || !h.end_date) return false;
          const isDateInRange = dStr >= h.start_date && dStr <= h.end_date;
          const isUserExcluded = h.excluded_account_ids?.includes(user!.id);
          return isDateInRange && !isUserExcluded;
        });

        if (isSpecialHoliday) continue;

        const dayOfWeek = d.getDay(); // 0 (Sun) to 6 (Sat)
        
        const rule = rules.find(r => r.day_of_week === dayOfWeek);
        if (!rule || rule.is_holiday) continue; // Skip holidays or no rule

        // Check if there is already an approved or pending leave/permission/maternity leave for this date
        const hasAnnual = annualLeaves.some(l => dStr >= l.start_date && dStr <= l.end_date);
        const hasMandatory = mandatoryLeaves.some(l => dStr >= l.start_date && dStr <= l.end_date);
        const hasPermission = permissions.some(p => dStr >= p.start_date && dStr <= p.end_date);
        const hasMaternity = maternityLeaves.some(m => dStr >= m.start_date && dStr <= m.end_date);
        
        if (hasAnnual || hasMandatory || hasPermission || hasMaternity) continue;

        // Check if there is already a pending or approved dispensation request for this date
        const hasExistingRequest = existingRequests.some(r => r.date === dStr);
        if (hasExistingRequest) continue;

        const dailyPresence = filteredAttendance.find(a => a.created_at?.split('T')[0] === dStr);
        
        const issues: DispensationIssueType[] = [];
        let details = "";

        if (!dailyPresence) {
          issues.push('ABSENT');
          details = "Tidak ada data presensi (Mangkir)";
        } else {
          if (dailyPresence.late_minutes > 0) {
            issues.push('LATE');
            details += `Terlambat ${dailyPresence.late_minutes}m. `;
          }
          if (dailyPresence.early_departure_minutes > 0) {
            issues.push('EARLY_LEAVE');
            details += `Pulang Awal ${dailyPresence.early_departure_minutes}m. `;
          }
          if (dailyPresence.check_in && !dailyPresence.check_out) {
            issues.push('NO_CLOCK_OUT');
            details += "Tanpa Presensi Pulang. ";
          }
        }

        if (issues.length > 0) {
          detected.push({
            date: dStr,
            presence_id: dailyPresence?.id || null,
            issues,
            details: details.trim()
          });
        }
      }

      setProblems(detected);
    } catch (error) {
      console.error(error);
    } finally {
      setIsDetecting(false);
    }
  };

  const handleIssueToggle = (issue: DispensationIssueType) => {
    setSelectedIssues(prev => 
      prev.includes(issue) ? prev.filter(i => i !== issue) : [...prev, issue]
    );
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedProblem || selectedIssues.length === 0 || !reason) {
      Swal.fire('Peringatan', 'Mohon lengkapi data pengajuan.', 'warning');
      return;
    }

    try {
      setIsLoading(true);
      let fileId = editData?.file_id || null;

      if (file) {
        setIsUploading(true);
        // Upload to Google Drive (folder Dispensasi)
        const folderId = (import.meta as any).env.VITE_DRIVE_FOLDER_DISPENSATION || '';
        fileId = await googleDriveService.uploadFile(file, folderId);
        setIsUploading(false);
      }

      const issues: DispensationIssue[] = selectedIssues.map(type => ({
        type,
        status: 'PENDING'
      }));

      if (editData) {
        await dispensationService.update(editData.id, {
          issues,
          reason,
          file_id: fileId
        });
      } else {
        await dispensationService.create({
          account_id: user!.id,
          presence_id: selectedProblem.presence_id,
          date: selectedProblem.date,
          issues,
          reason,
          file_id: fileId
        });
      }

      Swal.fire({
        title: 'Berhasil!',
        text: 'Pengajuan dispensasi telah dikirim.',
        icon: 'success',
        timer: 2000,
        showConfirmButton: false
      });
      onSuccess();
    } catch (error) {
      console.error(error);
      Swal.fire('Gagal', 'Terjadi kesalahan saat menyimpan data.', 'error');
    } finally {
      setIsLoading(false);
      setIsUploading(false);
    }
  };

  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[60] flex items-center justify-center p-4 animate-in fade-in duration-300">
      <div className="bg-white w-full max-w-2xl rounded-3xl shadow-2xl overflow-hidden flex flex-col max-h-[90vh]">
        {/* Header */}
        <div className="px-8 py-6 bg-gradient-to-r from-[#006E62] to-[#008a7b] text-white flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-white/20 rounded-xl flex items-center justify-center">
              <ClipboardList size={24} />
            </div>
            <div>
              <h3 className="text-lg font-bold tracking-tight">{editData ? 'Edit Pengajuan' : 'Buat Pengajuan Dispensasi'}</h3>
              <p className="text-[10px] text-white/70 font-bold uppercase tracking-widest">Lengkapi Detail Koreksi Presensi</p>
            </div>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-white/10 rounded-full transition-colors">
            <X size={24} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="flex-1 overflow-y-auto p-8 space-y-8">
          {/* Step 1: Pilih Tanggal & Masalah */}
          {!editData && (
            <div className="space-y-4">
              <div className="flex items-center gap-2 mb-2">
                <Calendar size={18} className="text-[#006E62]" />
                <h4 className="text-sm font-bold text-gray-800 uppercase tracking-wider">1. Pilih Tanggal Bermasalah</h4>
              </div>
              
              {isDetecting ? (
                <div className="py-8 flex flex-col items-center justify-center text-gray-400 bg-gray-50 rounded-2xl border border-dashed border-gray-200">
                  <Loader2 className="animate-spin mb-2" size={24} />
                  <p className="text-[10px] font-bold uppercase tracking-widest">Mendeteksi masalah presensi...</p>
                </div>
              ) : problems.length === 0 ? (
                <div className="p-6 bg-emerald-50 rounded-2xl border border-emerald-100 flex items-center gap-4">
                  <CheckCircle2 className="text-emerald-500 shrink-0" size={24} />
                  <p className="text-xs text-emerald-700 font-medium">Luar biasa! Tidak ditemukan masalah presensi dalam batas hari yang ditentukan.</p>
                </div>
              ) : (
                <div className="grid grid-cols-1 gap-3">
                  {problems.map((p, idx) => (
                    <button
                      key={idx}
                      type="button"
                      onClick={() => {
                        setSelectedProblem(p);
                        setSelectedIssues([]); // Reset issues when date changes
                      }}
                      className={`p-4 rounded-2xl border text-left transition-all flex items-center justify-between ${
                        selectedProblem?.date === p.date 
                        ? 'bg-[#006E62]/5 border-[#006E62] shadow-md shadow-[#006E62]/5' 
                        : 'bg-white border-gray-100 hover:border-gray-300'
                      }`}
                    >
                      <div className="flex items-center gap-4">
                        <div className={`w-10 h-10 rounded-xl flex items-center justify-center ${selectedProblem?.date === p.date ? 'bg-[#006E62] text-white' : 'bg-gray-100 text-gray-400'}`}>
                          <Calendar size={20} />
                        </div>
                        <div>
                          <p className="text-xs font-bold text-gray-800">{new Date(p.date).toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long' })}</p>
                          <p className="text-[10px] text-gray-500 italic">{p.details}</p>
                        </div>
                      </div>
                      {selectedProblem?.date === p.date && <CheckCircle2 className="text-[#006E62]" size={20} />}
                    </button>
                  ))}
                </div>
              )}
            </div>
          )}

          {/* Step 2: Pilih Jenis Dispensasi */}
          {selectedProblem && (
            <div className="space-y-4 animate-in slide-in-from-top-4 duration-500">
              <div className="flex items-center gap-2 mb-2">
                <Clock size={18} className="text-[#006E62]" />
                <h4 className="text-sm font-bold text-gray-800 uppercase tracking-wider">2. Pilih Masalah yang Diajukan</h4>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                {selectedProblem.issues.map((issue, idx) => (
                  <button
                    key={idx}
                    type="button"
                    onClick={() => handleIssueToggle(issue)}
                    className={`p-4 rounded-2xl border text-left transition-all flex items-center gap-3 ${
                      selectedIssues.includes(issue)
                      ? 'bg-emerald-50 border-emerald-500 text-emerald-700'
                      : 'bg-white border-gray-100 text-gray-500 hover:border-gray-300'
                    }`}
                  >
                    <div className={`w-6 h-6 rounded-full flex items-center justify-center border-2 ${selectedIssues.includes(issue) ? 'bg-emerald-500 border-emerald-500 text-white' : 'border-gray-200'}`}>
                      {selectedIssues.includes(issue) && <CheckCircle2 size={14} />}
                    </div>
                    <span className="text-[11px] font-bold uppercase tracking-wider">{issue.replace('_', ' ')}</span>
                  </button>
                ))}
              </div>
            </div>
          )}

          {/* Step 3: Alasan & Bukti */}
          {selectedIssues.length > 0 && (
            <div className="space-y-6 animate-in slide-in-from-top-4 duration-500">
              <div className="flex items-center gap-2 mb-2">
                <FileText size={18} className="text-[#006E62]" />
                <h4 className="text-sm font-bold text-gray-800 uppercase tracking-wider">3. Alasan & Bukti Pendukung</h4>
              </div>

              <div className="space-y-2">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest ml-1">Alasan Pengajuan</label>
                <textarea
                  required
                  value={reason}
                  onChange={(e) => setReason(e.target.value)}
                  placeholder="Jelaskan alasan Anda mengajukan dispensasi ini..."
                  className="w-full px-5 py-4 rounded-2xl border border-gray-200 text-sm focus:ring-4 focus:ring-[#006E62]/10 focus:border-[#006E62] outline-none transition-all min-h-[120px] resize-none"
                />
              </div>

              <div className="space-y-2">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest ml-1">Upload Bukti (Foto/Dokumen)</label>
                <div className="relative group">
                  <input
                    type="file"
                    onChange={(e) => setFile(e.target.files?.[0] || null)}
                    className="hidden"
                    id="evidence-upload"
                    accept="image/*,application/pdf"
                  />
                  <label
                    htmlFor="evidence-upload"
                    className={`flex flex-col items-center justify-center w-full p-8 rounded-2xl border-2 border-dashed transition-all cursor-pointer ${
                      file ? 'bg-emerald-50 border-emerald-300' : 'bg-gray-50 border-gray-200 hover:bg-gray-100 hover:border-gray-300'
                    }`}
                  >
                    {file ? (
                      <>
                        <CheckCircle2 className="text-emerald-500 mb-2" size={32} />
                        <p className="text-xs font-bold text-emerald-700">{file.name}</p>
                        <p className="text-[10px] text-emerald-500 mt-1 uppercase tracking-widest">Klik untuk mengganti file</p>
                      </>
                    ) : (
                      <>
                        <Upload className="text-gray-400 mb-2 group-hover:scale-110 transition-transform" size={32} />
                        <p className="text-xs font-bold text-gray-600">Pilih File Bukti</p>
                        <p className="text-[10px] text-gray-400 mt-1 uppercase tracking-widest text-center">Format: JPG, PNG, PDF (Maks. 5MB)</p>
                      </>
                    )}
                  </label>
                </div>
              </div>
            </div>
          )}
        </form>

        {/* Footer */}
        <div className="px-8 py-6 bg-gray-50 border-t border-gray-100 flex items-center justify-between">
          <div className="flex items-center gap-2 text-blue-600">
            <Info size={14} />
            <p className="text-[10px] font-medium italic">Pengajuan akan diverifikasi oleh Admin.</p>
          </div>
          <div className="flex gap-3">
            <button
              type="button"
              onClick={onClose}
              className="px-6 py-3 rounded-xl text-xs font-bold text-gray-500 uppercase tracking-wider hover:bg-gray-200 transition-all"
            >
              Batal
            </button>
            <button
              onClick={handleSubmit}
              disabled={isLoading || isUploading || !selectedProblem || selectedIssues.length === 0 || !reason}
              className="flex items-center gap-2 bg-[#006E62] text-white px-8 py-3 rounded-xl font-bold text-xs uppercase tracking-wider hover:bg-[#005c52] transition-all shadow-lg shadow-[#006E62]/20 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {isLoading || isUploading ? (
                <>
                  <Loader2 className="animate-spin" size={18} />
                  {isUploading ? 'Mengupload...' : 'Menyimpan...'}
                </>
              ) : (
                <>
                  <Save size={18} />
                  Kirim Pengajuan
                </>
              )}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

export default DispensationForm;
