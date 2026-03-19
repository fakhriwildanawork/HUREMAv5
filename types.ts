export interface EmployeeReportData {
  totalEmployees: number;
  newEmployees: number;
  exitEmployees: number;
  genderRatio: { name: string; value: number }[];
  ageDistribution: { name: string; value: number }[];
  educationDistribution: { name: string; value: number }[];
  locationDistribution: { name: string; value: number }[];
  positionDistribution: { name: string; value: number }[];
  contractTypeDistribution: { name: string; value: number }[];
  tenureDistribution: { name: string; value: number }[];
  healthRiskProfile: { name: string; value: number }[];
  disciplineSummary: { name: string; value: number }[];
}

export interface Location {
  id: string;
  name: string;
  location_type: string;
  address: string;
  city: string;
  province: string;
  zip_code: string;
  phone: string;
  latitude: number;
  longitude: number;
  radius: number;
  description: string;
  search_all: string;
  image_google_id?: string | null;
  created_at?: string;
  updated_at?: string;
}

export interface LocationAdministration {
  id: string;
  location_id: string;
  admin_date: string;
  status: 'Milik Sendiri' | 'Sewa/Kontrak' | 'Kerjasama';
  due_date?: string | null;
  description?: string;
  file_ids: string[]; // Array ID Google Drive
  created_at?: string;
}

export interface Schedule {
  id: string;
  name: string;
  type: 1 | 2 | 3 | 4;
  tolerance_minutes: number;
  tolerance_checkin_minutes: number;
  start_date?: string | null;
  end_date?: string | null;
  excluded_account_ids: string[];
  created_at?: string;
  updated_at?: string;
  rules?: ScheduleRule[];
  location_ids?: string[];
}

export interface ScheduleRule {
  id: string;
  schedule_id: string;
  day_of_week?: number;
  check_in_time?: string | null;
  check_out_time?: string | null;
  is_holiday: boolean;
}

export interface AuthUser {
  id: string;
  full_name: string;
  internal_nik: string;
  access_code: string;
  role: 'admin' | 'user';
  gender: 'Laki-laki' | 'Perempuan';
  schedule_type: string;
  photo_google_id?: string | null;
  // Admin Permissions
  is_hr_admin?: boolean;
  is_performance_admin?: boolean;
  is_finance_admin?: boolean;
}

export interface Account {
  id: string;
  // Identitas
  full_name: string;
  nik_ktp: string;
  photo_google_id?: string | null;
  ktp_google_id?: string | null;
  gender: 'Laki-laki' | 'Perempuan';
  religion: string;
  dob: string | null;
  // Kontak & Sosial
  address: string;
  phone: string;
  email: string;
  marital_status: string;
  dependents_count: number;
  emergency_contact_name: string;
  emergency_contact_rel: string;
  emergency_contact_phone: string;
  // Pendidikan
  last_education: string;
  major: string; // Jurusan
  diploma_google_id?: string | null;
  // Karier & Penempatan
  internal_nik: string;
  position: string;
  grade: string;
  location_id: string | null; // Relasi ke Location (UUID)
  schedule_id: string | null; // Relasi ke Schedule (UUID)
  // FIX: Added location property to support joined data from Supabase
  location?: any;
  schedule?: Schedule;
  employee_type: 'Tetap' | 'Kontrak' | 'Harian' | 'Magang';
  start_date: string | null;
  end_date?: string | null;
  // Pengaturan Kerja & Presensi
  schedule_type: string;
  leave_quota: number;
  is_leave_accumulated: boolean;
  max_carry_over_days: number;
  carry_over_quota: number;
  maternity_leave_quota: number;
  is_presence_limited_checkin: boolean;
  is_presence_limited_checkout: boolean;
  is_presence_limited_ot_in: boolean;
  is_presence_limited_ot_out: boolean;
  // Keamanan & Medis
  access_code: string;
  password?: string;
  role: 'admin' | 'user';
  mcu_status?: string | null;
  health_risk?: string | null;
  
  created_at?: string;
  updated_at?: string;
  search_all?: string;
}

export interface Attendance {
  id: string;
  account_id: string;
  check_in: string | null;
  check_out: string | null;
  in_latitude: number | null;
  in_longitude: number | null;
  out_latitude: number | null;
  out_longitude: number | null;
  in_photo_id: string | null;
  out_photo_id: string | null;
  in_address: string | null;
  out_address: string | null;
  late_minutes: number;
  early_departure_minutes: number;
  late_reason: string | null;
  early_departure_reason: string | null;
  status_in: string;
  status_out: string;
  work_duration: string | null;
  created_at?: string;
}

export interface Overtime {
  id: string;
  account_id: string;
  check_in: string | null;
  check_out: string | null;
  in_latitude: number | null;
  in_longitude: number | null;
  out_latitude: number | null;
  out_longitude: number | null;
  in_photo_id: string | null;
  out_photo_id: string | null;
  in_address: string | null;
  out_address: string | null;
  duration_minutes: number;
  work_duration: string | null;
  reason: string | null;
  created_at?: string;
}

export interface CareerLog {
  id: string;
  account_id: string;
  position: string;
  grade: string;
  location_id?: string | null;
  location_name: string;
  schedule_id?: string | null;
  schedule_type?: string | null;
  file_sk_id?: string | null;
  notes?: string | null;
  change_date: string;
}

export interface CareerLogExtended extends CareerLog {
  account?: {
    full_name: string;
    internal_nik: string;
  };
}

export interface HealthLog {
  id: string;
  account_id: string;
  mcu_status: string;
  health_risk: string;
  file_mcu_id?: string | null;
  notes?: string | null;
  change_date: string;
}

export interface HealthLogExtended extends HealthLog {
  account?: {
    full_name: string;
    internal_nik: string;
  };
}

export interface AccountContract {
  id: string;
  account_id: string;
  contract_number: string;
  contract_type: string;
  start_date: string;
  end_date?: string | null;
  file_id?: string | null;
  notes?: string | null;
  created_at?: string;
  updated_at?: string;
}

export interface AccountContractExtended extends AccountContract {
  account?: {
    full_name: string;
    internal_nik: string;
  };
}

export interface AccountCertification {
  id: string;
  account_id: string;
  entry_date: string;
  cert_type: string;
  cert_name: string;
  cert_date: string;
  file_id?: string | null;
  notes?: string | null;
  created_at?: string;
  updated_at?: string;
}

export interface AccountCertificationExtended extends AccountCertification {
  account?: {
    full_name: string;
    internal_nik: string;
  };
}

export interface WarningLog {
  id: string;
  account_id: string;
  warning_type: 'Teguran' | 'SP1' | 'SP2' | 'SP3';
  reason: string;
  issue_date: string;
  file_id?: string | null;
  created_at?: string;
}

export interface WarningLogExtended extends WarningLog {
  account?: {
    full_name: string;
    internal_nik: string;
  };
}

export interface TerminationLog {
  id: string;
  account_id: string;
  termination_type: 'Pemecatan' | 'Resign';
  termination_date: string;
  reason: string;
  severance_amount: number;
  penalty_amount: number;
  file_id?: string | null;
  created_at?: string;
}

export interface TerminationLogExtended extends TerminationLog {
  account?: {
    full_name: string;
    internal_nik: string;
  };
}

export interface LeaveRequest {
  id: string;
  account_id: string;
  start_date: string;
  end_date: string;
  description: string;
  status: 'pending' | 'approved' | 'rejected';
  created_at?: string;
  updated_at?: string;
}

export interface LeaveRequestExtended extends LeaveRequest {
  account?: {
    full_name: string;
    internal_nik: string;
  };
}

export interface AnnualLeaveRequest {
  id: string;
  account_id: string;
  start_date: string;
  end_date: string;
  description: string;
  status: 'pending' | 'approved' | 'rejected' | 'negotiating' | 'cancelled';
  file_id?: string | null;
  negotiation_data: {
    role: 'admin' | 'user';
    start_date: string;
    end_date: string;
    reason: string;
    timestamp: string;
  }[];
  current_negotiator_role: 'admin' | 'user';
  created_at?: string;
  updated_at?: string;
  account?: {
    full_name: string;
    internal_nik: string;
  };
}

export interface MaternityLeaveRequest {
  id: string;
  account_id: string;
  start_date: string;
  end_date: string;
  description: string;
  status: 'pending' | 'approved' | 'rejected' | 'negotiating' | 'cancelled';
  file_id?: string | null;
  negotiation_data: {
    role: 'admin' | 'user';
    start_date: string;
    end_date: string;
    reason: string;
    timestamp: string;
  }[];
  current_negotiator_role: 'admin' | 'user';
  created_at?: string;
  updated_at?: string;
  account?: {
    full_name: string;
    internal_nik: string;
  };
}

export interface PermissionRequest {
  id: string;
  account_id: string;
  permission_type: string;
  start_date: string;
  end_date: string;
  description: string;
  status: 'pending' | 'approved' | 'rejected' | 'negotiating' | 'cancelled';
  file_id?: string | null;
  negotiation_data: {
    role: 'admin' | 'user';
    start_date: string;
    end_date: string;
    reason: string;
    timestamp: string;
  }[];
  current_negotiator_role: 'admin' | 'user';
  created_at?: string;
  updated_at?: string;
  account?: {
    full_name: string;
    internal_nik: string;
  };
}

export interface DigitalDocument {
  id: string;
  name: string;
  doc_type: string;
  file_id: string;
  description?: string;
  created_at?: string;
  updated_at?: string;
  allowed_account_ids?: string[];
}

export type SubmissionStatus = 'Pending' | 'Disetujui' | 'Ditolak' | 'Dibatalkan';

export interface Submission {
  id: string;
  account_id: string;
  type: 'Lembur' | 'Cuti' | 'Izin' | 'Koreksi Absen' | string;
  status: SubmissionStatus;
  submission_data: any; // Flexible JSON Data
  description: string;
  verifier_id?: string | null;
  verified_at?: string | null;
  verification_notes?: string | null;
  file_id?: string | null;
  created_at?: string;
  updated_at?: string;
  // Join properties
  account?: {
    full_name: string;
    internal_nik: string;
  };
}

export interface AppSetting {
  key: string;
  value: any;
  description?: string;
  updated_at?: string;
}

export type LocationInput = Omit<Location, 'id' | 'created_at' | 'updated_at' | 'search_all'>;
export type LocationAdminInput = Omit<LocationAdministration, 'id' | 'created_at'>;
export type AccountInput = Omit<Account, 'id' | 'created_at' | 'updated_at' | 'search_all' | 'location'>;
export type CareerLogInput = Omit<CareerLog, 'id'>;
export type HealthLogInput = Omit<HealthLog, 'id'>;
export type AccountContractInput = Omit<AccountContract, 'id' | 'created_at' | 'updated_at'>;
export type AccountContractInputExtended = AccountContractInput;
export type AccountCertificationInput = Omit<AccountCertification, 'id' | 'created_at' | 'updated_at'>;
export type WarningLogInput = Omit<WarningLog, 'id' | 'created_at'>;
export type TerminationLogInput = Omit<TerminationLog, 'id' | 'created_at'>;
export type LeaveRequestInput = Omit<LeaveRequest, 'id' | 'created_at' | 'updated_at' | 'status'>;
export type AnnualLeaveRequestInput = {
  account_id: string;
  start_date: string;
  end_date: string;
  description: string;
  file_id?: string | null;
};

export type PermissionRequestInput = {
  account_id: string;
  permission_type: string;
  start_date: string;
  end_date: string;
  description: string;
  file_id?: string | null;
};

export type MaternityLeaveRequestInput = {
  account_id: string;
  start_date: string;
  end_date: string;
  description: string;
  file_id?: string | null;
};
export type DocumentInput = Omit<DigitalDocument, 'id' | 'created_at' | 'updated_at' | 'allowed_account_ids'> & {
  allowed_account_ids: string[];
};
export type SubmissionInput = Omit<Submission, 'id' | 'created_at' | 'updated_at' | 'status' | 'verifier_id' | 'verified_at' | 'verification_notes' | 'account'>;

export type ScheduleInput = Omit<Schedule, 'id' | 'created_at' | 'updated_at' | 'rules' | 'location_ids'> & {
  rules: Omit<ScheduleRule, 'id' | 'schedule_id'>[];
  location_ids: string[];
};

export interface GoogleDriveFile {
  id: string;
  name: string;
  mimeType: string;
}

export type AttendanceInput = Omit<Attendance, 'id' | 'created_at'>;
export type OvertimeInput = Omit<Overtime, 'id' | 'created_at'>;

export interface KPI {
  id: string;
  account_id: string;
  title: string;
  description: string;
  weight: number;
  start_date: string;
  deadline: string;
  status: 'Active' | 'Pause' | 'Unverified' | 'Verified' | 'Unreported';
  supporting_links: string[];
  report_data?: {
    description: string;
    file_ids: string[];
    links: string[];
    self_assessment: number;
    reported_at: string;
  } | null;
  verification_data?: {
    score: number;
    notes: string;
    verified_at: string;
    verifier_id: string;
  } | null;
  created_at?: string;
  updated_at?: string;
  account?: {
    full_name: string;
    internal_nik: string;
  };
}

export type KPIInput = Omit<KPI, 'id' | 'created_at' | 'updated_at' | 'status' | 'report_data' | 'verification_data' | 'account'>;

export interface KeyActivity {
  id: string;
  title: string;
  description: string;
  weight: number;
  start_date: string;
  end_date: string;
  status: 'Active' | 'Pause' | 'Completed';
  recurrence_type: 'Once' | 'Daily' | 'Weekly' | 'Monthly' | 'EndOfMonth';
  recurrence_rule: {
    days_of_week?: number[];
    dates_of_month?: number[];
  } | null;
  supporting_links: string[];
  created_at?: string;
  updated_at?: string;
  assignments?: KeyActivityAssignment[];
}

export interface KeyActivityAssignment {
  id: string;
  activity_id: string;
  account_id: string;
  account?: {
    full_name: string;
    internal_nik: string;
  };
}

export interface KeyActivityReport {
  id: string;
  activity_id: string;
  account_id: string;
  due_date: string;
  reported_at: string;
  description: string;
  file_ids: string[];
  links: string[];
  status: 'Unverified' | 'Verified';
  verification_data?: {
    score: number;
    notes: string;
    verified_at: string;
    verifier_id: string;
  } | null;
  created_at?: string;
  updated_at?: string;
  account?: {
    full_name: string;
    internal_nik: string;
  };
  activity?: KeyActivity;
}

export interface SalesReport {
  id: string;
  account_id: string;
  customer_name: string;
  activity_type: 'Cold Call' | 'Site Survey' | 'Product Demo' | 'Offering' | 'Negotiation' | 'Closing' | 'Maintenance';
  description: string;
  latitude: number;
  longitude: number;
  address?: string;
  photo_urls: string[];
  file_ids: string[];
  reported_at: string;
  created_at?: string;
  updated_at?: string;
  account?: {
    full_name: string;
    internal_nik: string;
  };
}

export type SalesReportInput = Omit<SalesReport, 'id' | 'created_at' | 'updated_at' | 'account' | 'reported_at'>;

export type KeyActivityInput = Omit<KeyActivity, 'id' | 'created_at' | 'updated_at' | 'status' | 'assignments'> & {
  assigned_account_ids: string[];
};

export type KeyActivityReportInput = Omit<KeyActivityReport, 'id' | 'created_at' | 'updated_at' | 'status' | 'verification_data' | 'account' | 'activity' | 'reported_at'>;

export interface Feedback {
  id: string;
  account_id: string;
  category: 'Gaji' | 'Fasilitas' | 'Hubungan Kerja' | 'Lainnya' | string;
  priority: 'Low' | 'Medium' | 'High';
  is_anonymous: boolean;
  description: string;
  attachments: string[]; // Array of file IDs
  links: string[]; // Array of URLs
  status: 'Unread' | 'Read';
  created_at: string;
  updated_at?: string;
  account?: {
    full_name: string;
    internal_nik: string;
  };
}

export type FeedbackInput = Omit<Feedback, 'id' | 'created_at' | 'updated_at' | 'status' | 'account'>;

export interface Whistleblowing {
  id: string;
  account_id: string;
  category: 'Pencurian' | 'Perusakan' | 'Bullying' | 'Fraud' | 'Pelanggaran SOP' | 'Lainnya' | string;
  description: string;
  reported_account_ids: string[]; // Array of account IDs being reported
  attachments: string[]; // Array of file IDs
  links: string[]; // Array of URLs
  status: 'Unread' | 'Read';
  created_at: string;
  updated_at?: string;
  account?: {
    full_name: string;
    internal_nik: string;
  };
  reported_accounts?: {
    id: string;
    full_name: string;
    internal_nik: string;
  }[];
}

export type WhistleblowingInput = Omit<Whistleblowing, 'id' | 'created_at' | 'updated_at' | 'status' | 'account' | 'reported_accounts'>;

export interface MeetingNote {
  id: string;
  meeting_id: string;
  content: string;
  attachments: string[];
  links: string[];
  created_at: string;
  updated_at?: string;
}

export interface Announcement {
  id: string;
  title: string;
  content: string;
  category: 'Urgent' | 'Info' | 'Event' | 'Policy';
  target_type: 'All' | 'Department' | 'Individual';
  target_ids: string[]; // Department names or User IDs
  publish_start: string;
  publish_end: string;
  attachments: string[];
  created_by: string;
  created_at: string;
  updated_at: string;
  creator?: { full_name: string };
  is_read?: boolean;
  read_count?: number;
}

export interface AnnouncementRead {
  id: string;
  announcement_id: string;
  user_id: string;
  read_at: string;
}

export interface Meeting {
  id: string;
  title: string;
  description: string;
  scheduled_at: string;
  started_at?: string;
  ended_at?: string;
  location_type: 'Online' | 'Offline';
  location_detail: string; // URL for online, address for offline
  latitude?: number;
  longitude?: number;
  participant_ids: string[];
  notulen_ids: string[];
  minutes_content?: string;
  attachments: string[]; // Array of file IDs
  links: string[]; // Array of URLs
  status: 'Scheduled' | 'In Progress' | 'Completed' | 'Cancelled';
  created_by: string;
  created_at: string;
  updated_at?: string;
  creator?: {
    full_name: string;
  };
  participants?: {
    id: string;
    full_name: string;
    internal_nik: string;
  }[];
  notulens?: {
    id: string;
    full_name: string;
    internal_nik: string;
  }[];
  notes?: MeetingNote[];
}

export type MeetingInput = Omit<Meeting, 'id' | 'created_at' | 'updated_at' | 'creator' | 'participants' | 'notulens' | 'notes'>;

export interface SalaryScheme {
  id: string;
  name: string;
  description: string | null;
  type: 'Harian' | 'Bulanan';
  basic_salary: number;
  position_allowance: number;
  placement_allowance: number;
  other_allowance: number;
  overtime_rate_per_hour: number;
  late_deduction_per_minute: number;
  early_leave_deduction_per_minute: number;
  no_clock_out_deduction_per_day: number;
  absent_deduction_per_day: number;
  created_at?: string;
  updated_at?: string;
}

export type SalarySchemeInput = Omit<SalaryScheme, 'id' | 'created_at' | 'updated_at'>;

export interface SalaryAssignment {
  id: string;
  scheme_id: string;
  account_id: string;
  created_at?: string;
  updated_at?: string;
}

export interface SalaryAssignmentExtended extends SalaryAssignment {
  account?: {
    full_name: string;
    internal_nik: string;
    position: string;
    grade: string;
    location_id: string;
  };
  scheme?: SalaryScheme;
}

export interface SalaryAdjustment {
  id: string;
  account_id: string;
  type: 'Addition' | 'Deduction';
  amount: number;
  month: number;
  year: number;
  description: string;
  created_at: string;
  updated_at: string;
  account?: {
    full_name: string;
    internal_nik: string;
  };
}

export type SalaryAdjustmentInput = Omit<SalaryAdjustment, 'id' | 'created_at' | 'updated_at' | 'account'>;

export interface PayrollStatus {
  id: string;
  month: number;
  year: number;
  status: 'Draft' | 'Approved' | 'Paid';
  updated_at: string;
}

export interface Payroll {
  id: string;
  month: number;
  year: number;
  start_date: string;
  end_date: string;
  status: 'Draft' | 'Pending' | 'Approved' | 'Paid';
  verifier_id?: string;
  verified_at?: string;
  verification_notes?: string;
  created_by?: string;
  updated_by?: string;
  created_at: string;
  updated_at: string;
  verifier?: {
    full_name: string;
  };
  creator?: {
    full_name: string;
  };
  updater?: {
    full_name: string;
  };
}

export interface PayrollItem {
  id: string;
  payroll_id: string;
  account_id: string;
  
  salary_type: 'Harian' | 'Bulanan';
  
  basic_salary: number;
  basic_salary_notes?: string;
  
  position_allowance: number;
  position_allowance_notes?: string;
  
  placement_allowance: number;
  placement_allowance_notes?: string;
  
  other_allowance: number;
  other_allowance_notes?: string;
  
  overtime_pay: number;
  overtime_pay_notes?: string;
  
  other_additions: number;
  other_additions_notes?: string;
  
  late_deduction: number;
  late_deduction_notes?: string;
  
  early_leave_deduction: number;
  early_leave_deduction_notes?: string;
  
  absent_deduction: number;
  absent_deduction_notes?: string;
  
  other_deductions: number;
  other_deductions_notes?: string;
  
  bpjs_kesehatan: number;
  bpjs_ketenagakerjaan: number;
  pph21: number;
  
  total_income: number;
  total_deduction: number;
  take_home_pay: number;
  
  created_at: string;
  updated_at: string;
  payroll?: Payroll;
  account?: {
    full_name: string;
    internal_nik: string;
    position: string;
    department: string;
    location: string;
    grade?: string;
  };
}

export type DispensationIssueType = 'LATE' | 'EARLY_LEAVE' | 'NO_CLOCK_OUT' | 'ABSENT';
export type DispensationIssueStatus = 'PENDING' | 'APPROVED' | 'REJECTED';

export interface DispensationIssue {
  type: DispensationIssueType;
  status: DispensationIssueStatus;
  admin_notes?: string;
}

export interface DispensationRequest {
  id: string;
  account_id: string;
  presence_id: string | null;
  date: string;
  issues: DispensationIssue[];
  reason: string;
  file_id: string | null;
  is_read: boolean;
  status: 'PENDING' | 'APPROVED' | 'REJECTED' | 'PARTIAL';
  created_at: string;
  updated_at?: string;
  account?: {
    full_name: string;
    internal_nik: string;
  };
}

export type DispensationRequestInput = Omit<DispensationRequest, 'id' | 'created_at' | 'updated_at' | 'status' | 'is_read' | 'account'>;

export interface PayrollSettings {
  id: string;
  company_name: string;
  company_address?: string;
  company_phone?: string;
  company_email?: string;
  company_website?: string;
  company_logo_url?: string;
  updated_at: string;
}

export type ReimbursementStatus = 'Pending' | 'Approved' | 'Partially Approved' | 'Rejected';
export type ReimbursementCategory = 'Operasional' | 'Akomodasi' | 'Inventaris' | 'Lainnya';

export interface Reimbursement {
  id: string;
  account_id: string;
  transaction_date: string;
  category: string;
  description: string;
  amount_requested: number;
  amount_approved: number | null;
  proof_file_id: string | null;
  payment_method: 'Cash' | 'Transfer';
  target_type: 'Bank' | 'E-Wallet' | null;
  target_name: string | null;
  account_number: string | null;
  account_holder: string | null;
  status: ReimbursementStatus;
  is_read: boolean;
  admin_notes: string | null;
  payment_proof_id: string | null;
  verifier_id: string | null;
  verified_at: string | null;
  created_at: string;
  updated_at: string;
  // Join properties
  account?: {
    full_name: string;
    internal_nik: string;
  };
}

export type ReimbursementInput = Omit<Reimbursement, 'id' | 'account_id' | 'status' | 'is_read' | 'amount_approved' | 'admin_notes' | 'payment_proof_id' | 'verifier_id' | 'verified_at' | 'created_at' | 'updated_at' | 'account'>;

export type EarlySalaryStatus = 'Pending' | 'Approved' | 'Paid' | 'Rejected';

export interface EarlySalaryRequest {
  id: string;
  account_id: string;
  month: number;
  year: number;
  amount: number;
  reason: string;
  status: EarlySalaryStatus;
  payment_proof_id: string | null;
  rejection_reason: string | null;
  verifier_id: string | null;
  verified_at: string | null;
  created_at: string;
  updated_at: string;
  // Join properties
  account?: {
    full_name: string;
    internal_nik: string;
  };
}

export type EarlySalaryRequestInput = Omit<EarlySalaryRequest, 'id' | 'account_id' | 'status' | 'payment_proof_id' | 'rejection_reason' | 'verifier_id' | 'verified_at' | 'created_at' | 'updated_at' | 'account'>;

export type CompensationStatus = 'Pending' | 'Completed';

export interface Compensation {
  id: string;
  account_id: string;
  termination_type: 'Resign' | 'Pemecatan';
  termination_date: string;
  amount: number;
  type: 'Severance' | 'Penalty';
  reason: string;
  status: CompensationStatus;
  is_read: boolean;
  transaction_date?: string | null;
  processed_amount?: number | null;
  notes?: string | null;
  proof_file_id?: string | null;
  created_at: string;
  updated_at: string;
  // Join properties
  account?: {
    full_name: string;
    internal_nik: string;
  };
}

export type CompensationInput = Omit<Compensation, 'id' | 'status' | 'is_read' | 'transaction_date' | 'processed_amount' | 'notes' | 'proof_file_id' | 'created_at' | 'updated_at' | 'account'>;

export interface EmployeeOfThePeriod {
  id: string;
  account_ids: string[];
  month: number;
  year: number;
  reason: string | null;
  created_at: string;
  updated_at: string;
  created_by?: string;
  // Join properties
  accounts?: {
    id: string;
    full_name: string;
    internal_nik: string;
    photo_google_id?: string | null;
    position: string;
  }[];
}

export type EmployeeOfThePeriodInput = Omit<EmployeeOfThePeriod, 'id' | 'created_at' | 'updated_at' | 'accounts'>;

// --- Report Types ---
export interface AttendanceSummary {
  accountId: string;
  fullName: string;
  nik: string;
  totalDays: number;
  present: number;
  late: number;
  lateMinutes: number;
  earlyDeparture: number;
  earlyDepartureMinutes: number;
  absent: number;
  leave: number;
  maternityLeave: number;
  permission: number;
  holiday: number;
  specialHoliday: number;
  noClockOut: number;
  dispensationCount: number;
  attendanceRate: number;
  dailyDetails: {
    date: string;
    status: 'PRESENT' | 'ABSENT' | 'LEAVE' | 'MATERNITY' | 'PERMISSION' | 'HOLIDAY' | 'SPECIAL_HOLIDAY' | 'WEEKEND';
    isLate: boolean;
    isEarlyDeparture: boolean;
    isNoClockOut: boolean;
  }[];
}

export interface OvertimeSummary {
  accountId: string;
  fullName: string;
  nik: string;
  totalOvertimeMinutes: number;
  totalOvertimeHours: number;
  overtimeCount: number;
  estimatedCost: number;
}

export interface LeaveSummary {
  accountId: string;
  fullName: string;
  nik: string;
  totalQuota: number;
  usedQuota: number;
  remainingQuota: number;
  carryOverQuota: number;
  maternityQuota: number;
  maternityUsed: number;
  permissionCount: number;
}

export interface PayrollSummary {
  month: number;
  year: number;
  totalBasicSalary: number;
  totalAllowances: number;
  totalOvertime: number;
  totalDeductions: number;
  totalPph21: number;
  totalBpjsKesehatan: number;
  totalBpjsKetenagakerjaan: number;
  totalTakeHomePay: number;
  itemCount: number;
  items: PayrollItem[];
}

export interface ExpenseSummary {
  totalRequested: number;
  totalApproved: number;
  count: number;
  categories: { [key: string]: number };
  items: Reimbursement[];
}

export interface CompensationSummary {
  totalSeverance: number;
  totalPenalty: number;
  count: number;
  items: Compensation[];
}
