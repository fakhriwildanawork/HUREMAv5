
import React, { useState, useEffect } from 'react';
import { QRCodeSVG } from 'qrcode.react';
import { X, Edit2, Trash2, User, UserCircle, Phone, Mail, Calendar, MapPin, Briefcase, Shield, Heart, GraduationCap, Download, ExternalLink, Clock, Activity, Plus, Paperclip, FileBadge, Award, ShieldAlert, LogOut } from 'lucide-react';
import Swal from 'sweetalert2';
import { Account, CareerLog, HealthLog, AccountContract, AccountCertification, WarningLog, TerminationLog, SalaryScheme } from '../../types';
import { accountService } from '../../services/accountService';
import { contractService } from '../../services/contractService';
import { certificationService } from '../../services/certificationService';
import { disciplineService } from '../../services/disciplineService';
import { financeService } from '../../services/financeService';
import { googleDriveService } from '../../services/googleDriveService';
import LoadingSpinner from '../../components/Common/LoadingSpinner';
import LogForm from './LogForm';
import CertificationFormModal from '../certification/CertificationFormModal';
import ContractFormModal from '../contract/ContractFormModal';
import WarningForm from '../discipline/WarningForm';
import TerminationForm from '../discipline/TerminationForm';
import ContractDetailModal from '../contract/ContractDetailModal';
import CareerDetailModal from './CareerDetailModal';
import HealthDetailModal from './HealthDetailModal';
import CertificationDetailModal from './CertificationDetailModal';
import WarningDetailModal from './WarningDetailModal';
import TerminationDetailModal from './TerminationDetailModal';

interface AccountDetailProps {
  id: string;
  onClose: () => void;
  onEdit: (account: Account) => void;
  onDelete: (id: string) => void;
  isReadOnly?: boolean;
}

const AccountDetail: React.FC<AccountDetailProps> = ({ id, onClose, onEdit, onDelete, isReadOnly = false }) => {
  const [account, setAccount] = useState<Account | null>(null);
  const [careerLogs, setCareerLogs] = useState<CareerLog[]>([]);
  const [healthLogs, setHealthLogs] = useState<HealthLog[]>([]);
  const [contracts, setContracts] = useState<AccountContract[]>([]);
  const [certs, setCerts] = useState<AccountCertification[]>([]);
  const [warnings, setWarnings] = useState<WarningLog[]>([]);
  const [termination, setTermination] = useState<TerminationLog | null>(null);
  const [salaryScheme, setSalaryScheme] = useState<SalaryScheme | null>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [showLogForm, setShowLogForm] = useState<{ type: 'career' | 'health', data?: any, isEdit?: boolean } | null>(null);
  const [showCertForm, setShowCertForm] = useState<{ show: boolean, data?: any }>({ show: false });
  const [showContractForm, setShowContractForm] = useState<{ show: boolean, data?: any }>({ show: false });
  const [showWarningForm, setShowWarningForm] = useState(false);
  const [showTerminationForm, setShowTerminationForm] = useState(false);
  const [selectedContractDetail, setSelectedContractDetail] = useState<AccountContract | null>(null);
  const [selectedCareerDetail, setSelectedCareerDetail] = useState<CareerLog | null>(null);
  const [selectedHealthDetail, setSelectedHealthDetail] = useState<HealthLog | null>(null);
  const [selectedCertDetail, setSelectedCertDetail] = useState<AccountCertification | null>(null);
  const [selectedWarningDetail, setSelectedWarningDetail] = useState<WarningLog | null>(null);
  const [selectedTerminationDetail, setSelectedTerminationDetail] = useState<TerminationLog | null>(null);
  
  // Media Preview States
  const [previewMedia, setPreviewMedia] = useState<{ url: string, title: string, type: 'image' | 'qr' } | null>(null);

  useEffect(() => {
    fetchData();
  }, [id]);

  const fetchData = async () => {
    try {
      setIsLoading(true);
      const [acc, careers, healths, contractList, certList, warningList, term, assignment] = await Promise.all([
        accountService.getById(id),
        accountService.getCareerLogs(id),
        accountService.getHealthLogs(id),
        contractService.getByAccountId(id),
        certificationService.getByAccountId(id),
        disciplineService.getWarningsByAccountId(id),
        disciplineService.getTerminationByAccountId(id),
        financeService.getAssignmentByAccountId(id)
      ]);
      setAccount(acc as any);
      setCareerLogs(careers);
      setHealthLogs(healths);
      setContracts(contractList);
      setCerts(certList);
      setWarnings(warningList);
      setTermination(term || null);
      if (assignment && assignment.scheme) {
        setSalaryScheme(assignment.scheme);
      } else {
        setSalaryScheme(null);
      }
    } catch (error) {
      console.error(error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleLogSubmit = async (data: any) => {
    setIsSaving(true);
    const type = showLogForm?.type;
    const isEdit = showLogForm?.isEdit;
    const tempId = data.id || `temp-${Date.now()}`;
    const optimisticEntry = { ...data, id: tempId, change_date: data.change_date || new Date().toISOString() };

    try {
      if (type === 'career') {
        if (isEdit) {
          const updated = await accountService.updateCareerLog(data.id, data);
          setCareerLogs(prev => prev.map(l => l.id === data.id ? updated : l));
        } else {
          setCareerLogs(prev => [optimisticEntry, ...prev]);
          const created = await accountService.createCareerLog(data);
          setCareerLogs(prev => prev.map(l => l.id === tempId ? created : l));
        }
        setAccount(prev => prev ? { ...prev, position: data.position, grade: data.grade, location_id: data.location_id, location: { ...prev.location, name: data.location_name } } : null);
      } else if (type === 'health') {
        if (isEdit) {
          const updated = await accountService.updateHealthLog(data.id, data);
          setHealthLogs(prev => prev.map(l => l.id === data.id ? updated : l));
        } else {
          setHealthLogs(prev => [optimisticEntry, ...prev]);
          const created = await accountService.createHealthLog(data);
          setHealthLogs(prev => prev.map(l => l.id === tempId ? created : l));
        }
        setAccount(prev => prev ? { ...prev, mcu_status: data.mcu_status, health_risk: data.health_risk } : null);
      }
      setShowLogForm(null);
      Swal.fire({ title: 'Berhasil!', text: `Riwayat telah ${isEdit ? 'diperbarui' : 'ditambahkan'}.`, icon: 'success', timer: 1000, showConfirmButton: false });
    } catch (error) {
      if (type === 'career' && !isEdit) setCareerLogs(prev => prev.filter(l => l.id !== tempId));
      else if (type === 'health' && !isEdit) setHealthLogs(prev => prev.filter(l => l.id !== tempId));
      Swal.fire('Gagal', 'Gagal menyimpan riwayat', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleDeleteLog = async (logId: string, type: 'career' | 'health') => {
    const result = await Swal.fire({
      title: 'Hapus riwayat?',
      text: "Data ini tidak dapat dikembalikan.",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#006E62',
      cancelButtonColor: '#ef4444',
      confirmButtonText: 'Ya, hapus!',
      cancelButtonText: 'Batal'
    });

    if (result.isConfirmed) {
      setIsSaving(true);
      try {
        if (type === 'career') {
          await accountService.deleteCareerLog(logId);
          setCareerLogs(prev => prev.filter(l => l.id !== logId));
        } else if (type === 'health') {
          await accountService.deleteHealthLog(logId);
          setHealthLogs(prev => prev.filter(l => l.id !== logId));
        }
        Swal.fire('Terhapus!', 'Riwayat telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Gagal menghapus data', 'error');
      } finally {
        setIsSaving(false);
      }
    }
  };

  const handleDeleteContract = async (contractId: string) => {
    const result = await Swal.fire({
      title: 'Hapus Kontrak?',
      text: "Data kontrak akan dihapus permanen.",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#006E62',
      cancelButtonColor: '#ef4444',
      confirmButtonText: 'Ya, hapus!'
    });

    if (result.isConfirmed) {
      setIsSaving(true);
      try {
        await contractService.delete(contractId);
        setContracts(prev => prev.filter(c => c.id !== contractId));
        Swal.fire('Terhapus!', 'Kontrak telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Gagal menghapus.', 'error');
      } finally {
        setIsSaving(false);
      }
    }
  };

  const handleDeleteCert = async (certId: string) => {
    const result = await Swal.fire({
      title: 'Hapus Sertifikasi?',
      text: "Data sertifikasi ini akan dihapus permanen.",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#006E62',
      cancelButtonColor: '#ef4444',
      confirmButtonText: 'Ya, hapus!',
      cancelButtonText: 'Batal'
    });

    if (result.isConfirmed) {
      setIsSaving(true);
      try {
        await certificationService.delete(certId);
        setCerts(prev => prev.filter(c => c.id !== certId));
        Swal.fire('Terhapus!', 'Data sertifikasi telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Gagal menghapus data', 'error');
      } finally {
        setIsSaving(false);
      }
    }
  };

  const handleDeleteWarning = async (logId: string) => {
    const res = await Swal.fire({ 
      title: 'Hapus riwayat peringatan?', 
      icon: 'warning', 
      showCancelButton: true, 
      confirmButtonColor: '#006E62',
      confirmButtonText: 'Ya, Hapus'
    });
    if (res.isConfirmed) {
      try {
        setIsSaving(true);
        await disciplineService.deleteWarning(logId);
        setWarnings(prev => prev.filter(w => w.id !== logId));
        Swal.fire('Terhapus', '', 'success');
      } catch (e) { Swal.fire('Gagal', 'Gagal menghapus data', 'error'); }
      finally { setIsSaving(false); }
    }
  };

  const downloadQR = () => {
    const svg = document.getElementById('qr-member-code');
    if (!svg) return;
    const svgData = new XMLSerializer().serializeToString(svg);
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    const img = new Image();
    img.onload = () => {
      canvas.width = img.width;
      canvas.height = img.height;
      ctx?.drawImage(img, 0, 0);
      const pngFile = canvas.toDataURL('image/png');
      const downloadLink = document.createElement('a');
      downloadLink.download = `QR_${account?.full_name}_${account?.internal_nik}.png`;
      downloadLink.href = `${pngFile}`;
      downloadLink.click();
    };
    img.src = 'data:image/svg+xml;base64,' + btoa(svgData);
  };

  if (isLoading) return <div className="flex justify-center py-20"><div className="w-8 h-8 border-4 border-[#006E62] border-t-transparent rounded-full animate-spin"></div></div>;
  if (!account) return null;

  const today = new Date().toISOString().split('T')[0];
  const isInactive = account.end_date && account.end_date <= today;

  // Sync data with latest career log
  const latestCareer = careerLogs[0];
  const currentPosition = latestCareer?.position || account.position;
  const currentGrade = latestCareer?.grade || account.grade;
  const currentLocation = latestCareer?.location_name || account.location?.name || '-';

  const DetailSection = ({ icon: Icon, title, onAdd, children, isScrollable = false }: { icon: any, title: string, onAdd?: () => void, children: React.ReactNode, isScrollable?: boolean }) => (
    <div className={`bg-white border border-gray-100 p-5 rounded-md shadow-sm flex flex-col ${isScrollable ? 'h-[320px]' : ''}`}>
      <div className="flex items-center justify-between border-b border-gray-50 pb-3 mb-4 shrink-0">
        <div className="flex items-center gap-2">
          <Icon size={16} className="text-[#006E62]" />
          <h4 className="text-[11px] font-bold uppercase tracking-widest text-gray-400">{title}</h4>
        </div>
        {onAdd && !isReadOnly && (
          <button onClick={onAdd} className="p-1 hover:bg-gray-50 text-[#006E62] rounded transition-colors">
            <Plus size={16} />
          </button>
        )}
      </div>
      <div className={`space-y-4 ${isScrollable ? 'flex-1 overflow-y-auto pr-2 scrollbar-thin' : ''}`}>{children}</div>
    </div>
  );

  const DataRow = ({ label, value, isFile = false }: { label: string, value: any, isFile?: boolean }) => (
    <div>
      <p className="text-[9px] font-bold text-gray-400 uppercase tracking-tighter mb-0.5">{label}</p>
      {isFile && value ? (
        <button 
          onClick={() => window.open(googleDriveService.getViewerUrl(value), '_blank')}
          className="flex items-center gap-1.5 text-[11px] text-[#006E62] font-bold hover:underline"
        >
          <Paperclip size={10} /> LIHAT DOKUMEN
        </button>
      ) : (
        <p className="text-xs text-gray-700 font-medium leading-tight">{value || '-'}</p>
      )}
    </div>
  );

  const formatDate = (dateStr: string) => {
    if (!dateStr) return '-';
    return new Date(dateStr).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' });
  };

  const formatCurrency = (amount?: number | null) => {
    if (!amount) return 'Rp 0';
    return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(amount);
  };

  return (
    <div className="space-y-6 pb-20">
      {isSaving && <LoadingSpinner />}
      
      {/* Header Profile */}
      <div className="bg-white rounded-md border border-gray-100 p-6 flex flex-col md:flex-row gap-6 items-start shadow-sm">
        <div 
          className="w-32 h-32 rounded-md border-4 border-gray-50 overflow-hidden shrink-0 shadow-inner cursor-pointer hover:opacity-90 transition-opacity"
          onClick={() => account.photo_google_id && window.open(googleDriveService.getViewerUrl(account.photo_google_id), '_blank')}
        >
          {account.photo_google_id ? (
            <img src={googleDriveService.getFileUrl(account.photo_google_id)} className="w-full h-full object-cover" />
          ) : (
            <div className="w-full h-full flex items-center justify-center bg-gray-100 text-gray-400"><UserCircle size={48} /></div>
          )}
        </div>
        
        <div className="flex-1 space-y-2">
          <div className="flex items-center gap-3">
             <h2 className="text-2xl font-bold text-gray-800 tracking-tight">{account.full_name}</h2>
             <span className={`px-2 py-0.5 text-[10px] font-bold uppercase rounded ${isInactive ? 'bg-gray-100 text-gray-400' : 'bg-[#006E62]/10 text-[#006E62]'}`}>
               {account.employee_type}
             </span>
             {(termination || isInactive) && <span className="px-2 py-0.5 bg-red-50 text-red-600 text-[10px] font-bold uppercase rounded">NON-AKTIF</span>}
          </div>
          <p className="text-sm font-bold text-gray-500 uppercase tracking-widest">{currentPosition} • {currentGrade} • {account.internal_nik}</p>
          <div className="flex flex-wrap gap-4 pt-2">
             <div className="flex items-center gap-1.5 text-xs text-gray-600"><MapPin size={14} className="text-gray-400" /> {currentLocation}</div>
             <div className="flex items-center gap-1.5 text-xs text-gray-600"><Mail size={14} className="text-gray-400" /> {account.email || '-'}</div>
             <div className="flex items-center gap-1.5 text-xs text-gray-600"><Phone size={14} className="text-gray-400" /> {account.phone || '-'}</div>
          </div>
        </div>

        <div 
          className="p-3 bg-gray-50 rounded border border-gray-100 flex flex-col items-center gap-2 cursor-pointer hover:bg-gray-100 transition-colors"
          onClick={() => setPreviewMedia({ url: account.id, title: 'Member QR Code', type: 'qr' })}
        >
          <QRCodeSVG id="qr-member-code" value={account.id} size={80} bgColor="#F9FAFB" />
          <p className="text-[8px] font-bold text-gray-400 tracking-widest uppercase">Member ID</p>
        </div>

        <div className="flex gap-2">
           {!isReadOnly && (
             <>
               <button onClick={() => onEdit(account)} className="p-2 border border-gray-100 rounded text-[#006E62] bg-emerald-50/50 hover:bg-emerald-50 transition-colors"><Edit2 size={16} /></button>
               <button onClick={() => onDelete(account.id)} className="p-2 border border-gray-100 rounded text-red-500 bg-red-50/50 hover:bg-red-50 transition-colors"><Trash2 size={16} /></button>
             </>
           )}
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        {/* a. Informasi Personal */}
        <DetailSection icon={User} title="Informasi Personal">
          <div className="grid grid-cols-2 gap-4">
             <DataRow label="NIK KTP" value={account.nik_ktp} />
             <DataRow label="Tanggal Lahir" value={formatDate(account.dob || '')} />
             <DataRow label="Gender" value={account.gender} />
             <DataRow label="Agama" value={account.religion} />
             <DataRow label="Status Nikah" value={account.marital_status} />
             <DataRow label="Tanggungan" value={account.dependents_count} />
          </div>
          <DataRow label="Scan KTP" value={account.ktp_google_id} isFile />
          <DataRow label="Alamat Domisili" value={account.address} />
        </DetailSection>

        {/* b. Karier & Penempatan */}
        <DetailSection icon={Briefcase} title="Karier & Penempatan">
          <div className="grid grid-cols-2 gap-4">
             <DataRow label="Jabatan" value={currentPosition} />
             <DataRow label="Golongan" value={currentGrade} />
             <DataRow label="NIK Internal" value={account.internal_nik} />
             <DataRow label="Jadwal" value={account.schedule_type} />
             <DataRow label="Mulai Kerja" value={formatDate(account.start_date || '')} />
             <DataRow label="Akhir Kerja" value={account.end_date ? formatDate(account.end_date) : 'Aktif'} />
          </div>
        </DetailSection>

        {/* c. Presensi & Akses */}
        <DetailSection icon={Shield} title="Presensi & Akses">
           <div className="space-y-3">
              {!isReadOnly && (
                <div className="flex items-center justify-between p-2 bg-gray-50 rounded text-[11px] font-bold">
                  <span className="text-gray-500">KODE AKSES</span>
                  <span className="text-[#006E62] tracking-widest">{account.access_code}</span>
                </div>
              )}
              <p className="text-[9px] font-bold text-gray-400 uppercase mt-2">Kebijakan Radius Presensi</p>
              <div className="grid grid-cols-2 gap-2">
                 {[
                   { id: 'is_presence_limited_checkin', label: 'Check-in Datang' },
                   { id: 'is_presence_limited_checkout', label: 'Check-out Pulang' },
                   { id: 'is_presence_limited_ot_in', label: 'Check-in Lembur' },
                   { id: 'is_presence_limited_ot_out', label: 'Check-out Lembur' }
                 ].map(item => (
                   <div key={item.id} className="flex items-center justify-between px-2 py-1.5 border border-gray-100 rounded bg-gray-50/50">
                      <span className="text-[9px] font-medium text-gray-600">{item.label}</span>
                      <span className={`text-[8px] font-bold uppercase ${account[item.id as keyof Account] ? 'text-[#006E62]' : 'text-orange-500'}`}>{account[item.id as keyof Account] ? 'Terbatas' : 'Bebas'}</span>
                   </div>
                 ))}
              </div>
           </div>
        </DetailSection>

        {/* d. Pendidikan & Dokumen */}
        <DetailSection icon={GraduationCap} title="Pendidikan & Dokumen">
           <div>
             <p className="text-[9px] font-bold text-gray-400 uppercase tracking-tighter mb-0.5">Pendidikan Terakhir</p>
             <p className="text-xs text-gray-700 font-medium leading-tight">{account.last_education} {account.major ? `- ${account.major}` : ''}</p>
           </div>
           <div className="grid grid-cols-1 gap-4 pt-2">
              <DataRow label="Scan Ijazah" value={account.diploma_google_id} isFile />
           </div>
        </DetailSection>

        {/* e. Kontak Darurat */}
        <DetailSection icon={Heart} title="Kontak Darurat">
           <div className="mt-2">
              <div className="space-y-3">
                <DataRow label="Nama Kontak" value={account.emergency_contact_name} />
                <div className="grid grid-cols-2 gap-4">
                  <DataRow label="Hubungan" value={account.emergency_contact_rel} />
                  <DataRow label="No HP" value={account.emergency_contact_phone} />
                </div>
              </div>
           </div>
        </DetailSection>

        {/* f. Riwayat Kontrak Kerja */}
        <DetailSection 
          icon={FileBadge} 
          title="Riwayat Kontrak Kerja"
          onAdd={() => setShowContractForm({ show: true, data: { account_id: id } })}
          isScrollable
        >
           <div className="space-y-3">
            {contracts.length === 0 ? (
              <p className="text-[10px] text-gray-400 italic">Belum ada riwayat kontrak.</p>
            ) : (
              contracts.map(c => (
                <div 
                  key={c.id} 
                  className="flex group justify-between items-start border-l-2 border-emerald-100 pl-3 py-1 relative cursor-pointer hover:bg-gray-50/50 transition-colors"
                  onClick={() => setSelectedContractDetail(c)}
                >
                  <div className="flex-1 min-w-0 pr-4">
                    <p className="text-[10px] font-bold text-[#006E62] leading-tight">{c.contract_number}</p>
                    <p className="text-[9px] text-gray-500 font-medium uppercase tracking-tighter">{c.contract_type}</p>
                    <div className="flex justify-between items-center mt-1">
                      <p className="text-[8px] text-gray-400 uppercase font-bold">{formatDate(c.start_date)} - {c.end_date ? formatDate(c.end_date) : 'TETAP'}</p>
                    </div>
                  </div>
                  <div className="flex gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                    {!isReadOnly && (
                      <>
                        <button 
                          onClick={(e) => { e.stopPropagation(); setShowContractForm({ show: true, data: c }); }} 
                          className="text-[#006E62] hover:opacity-80"
                        >
                          <Edit2 size={12} />
                        </button>
                        <button 
                          onClick={(e) => { e.stopPropagation(); handleDeleteContract(c.id); }} 
                          className="text-red-500 hover:opacity-80"
                        >
                          <Trash2 size={12} />
                        </button>
                      </>
                    )}
                  </div>
                </div>
              ))
            )}
           </div>
        </DetailSection>

        {/* g. Riwayat Karir */}
        <DetailSection 
          icon={Clock} 
          title="Riwayat Karir" 
          onAdd={() => setShowLogForm({ type: 'career', data: account })}
          isScrollable
        >
          <div className="space-y-3">
            {careerLogs.length === 0 ? (
              <p className="text-[10px] text-gray-400 italic">Belum ada riwayat perubahan karir.</p>
            ) : (
              careerLogs.map((log) => {
                if (!log) return null; // Safety guard
                return (
                  <div 
                    key={log.id} 
                    className="flex group justify-between items-start border-l-2 border-emerald-100 pl-3 py-1 relative cursor-pointer hover:bg-gray-50/50 transition-colors"
                    onClick={() => setSelectedCareerDetail(log)}
                  >
                    <div className="flex-1 min-w-0 pr-4">
                      <p className="text-[10px] font-bold text-[#006E62] leading-tight">{log.position} • {log.grade}</p>
                      <p className="text-[9px] text-gray-400 font-medium uppercase tracking-tighter">{log.location_name}</p>
                      <div className="flex items-center gap-2 mt-0.5">
                        <p className="text-[8px] text-gray-400 font-bold uppercase">{formatDate(log.change_date)}</p>
                      </div>
                    </div>
                    <div className="flex gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                      {!isReadOnly && (
                        <>
                          <button 
                            onClick={(e) => { e.stopPropagation(); setShowLogForm({ type: 'career', data: log, isEdit: true }); }} 
                            className="text-[#006E62] hover:opacity-80"
                          >
                            <Edit2 size={12} />
                          </button>
                          <button 
                            onClick={(e) => { e.stopPropagation(); handleDeleteLog(log.id, 'career'); }} 
                            className="text-red-500 hover:opacity-80"
                          >
                            <Trash2 size={12} />
                          </button>
                        </>
                      )}
                    </div>
                  </div>
                );
              })
            )}
          </div>
        </DetailSection>

        {/* h. Daftar Sertifikasi */}
        <DetailSection 
          icon={Award} 
          title="Daftar Sertifikasi" 
          onAdd={() => setShowCertForm({ show: true, data: { account_id: id } })}
          isScrollable
        >
          <div className="space-y-3">
            {certs.length === 0 ? (
              <p className="text-[10px] text-gray-400 italic">Belum ada data sertifikasi.</p>
            ) : (
              certs.map((cert) => (
                <div 
                  key={cert.id} 
                  className="flex group justify-between items-start border-l-2 border-emerald-100 pl-3 py-1 relative cursor-pointer hover:bg-gray-50/50 transition-colors"
                  onClick={() => setSelectedCertDetail(cert)}
                >
                  <div className="flex-1 min-w-0 pr-4">
                    <p className="text-[10px] font-bold text-[#006E62] leading-tight">{cert.cert_name}</p>
                    <p className="text-[9px] text-gray-500 font-medium uppercase tracking-tighter">{cert.cert_type}</p>
                    <div className="flex items-center gap-2 mt-0.5">
                      <p className="text-[8px] text-gray-400 font-bold uppercase">{formatDate(cert.cert_date)}</p>
                    </div>
                  </div>
                  <div className="flex gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                    {!isReadOnly && (
                      <>
                        <button 
                          onClick={(e) => { e.stopPropagation(); setShowCertForm({ show: true, data: cert }); }} 
                          className="text-[#006E62] hover:opacity-80"
                        >
                          <Edit2 size={12} />
                        </button>
                        <button 
                          onClick={(e) => { e.stopPropagation(); handleDeleteCert(cert.id); }} 
                          className="text-red-500 hover:opacity-80"
                        >
                          <Trash2 size={12} />
                        </button>
                      </>
                    )}
                  </div>
                </div>
              ))
            )}
          </div>
        </DetailSection>

        {/* i. Riwayat Kesehatan */}
        <DetailSection 
          icon={Heart} 
          title="Riwayat Kesehatan" 
          onAdd={() => setShowLogForm({ type: 'health', data: { account_id: id } })}
          isScrollable
        >
          <div className="space-y-3">
            {healthLogs.length === 0 ? (
              <p className="text-[10px] text-gray-400 italic">Belum ada riwayat kesehatan.</p>
            ) : (
              healthLogs.map((log) => (
                <div 
                  key={log.id} 
                  className="flex group justify-between items-start border-l-2 border-emerald-100 pl-3 py-1 relative cursor-pointer hover:bg-gray-50/50 transition-colors"
                  onClick={() => setSelectedHealthDetail(log)}
                >
                  <div className="flex-1 min-w-0 pr-4">
                    <p className="text-[10px] font-bold text-[#006E62] leading-tight">{log.mcu_status}</p>
                    <p className="text-[9px] text-gray-500 font-medium uppercase tracking-tighter">{log.health_risk}</p>
                    <div className="flex items-center gap-2 mt-0.5">
                      <p className="text-[8px] text-gray-400 font-bold uppercase">{formatDate(log.change_date)}</p>
                    </div>
                  </div>
                  <div className="flex gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                    {!isReadOnly && (
                      <>
                        <button 
                          onClick={(e) => { e.stopPropagation(); setShowLogForm({ type: 'health', data: log, isEdit: true }); }} 
                          className="text-[#006E62] hover:opacity-80"
                        >
                          <Edit2 size={12} />
                        </button>
                        <button 
                          onClick={(e) => { e.stopPropagation(); handleDeleteLog(log.id, 'health'); }} 
                          className="text-red-500 hover:opacity-80"
                        >
                          <Trash2 size={12} />
                        </button>
                      </>
                    )}
                  </div>
                </div>
              ))
            )}
          </div>
        </DetailSection>

        {/* j. Status Kedisiplinan */}
        <DetailSection 
          icon={ShieldAlert} 
          title="Status Kedisiplinan" 
          onAdd={() => setShowWarningForm(true)} 
          isScrollable
        >
          <div className="space-y-3">
            {warnings.length === 0 ? (
              <p className="text-[10px] text-gray-400 italic">Belum ada riwayat peringatan.</p>
            ) : (
              warnings.map(w => (
                <div 
                  key={w.id} 
                  className="flex group justify-between items-start border-l-2 border-emerald-100 pl-3 py-1 relative cursor-pointer hover:bg-gray-50/50 transition-colors"
                  onClick={() => setSelectedWarningDetail(w)}
                >
                  <div className="flex-1 min-w-0 pr-4">
                    <p className="text-[10px] font-bold text-[#006E62] leading-tight">{w.warning_type}</p>
                    <p className="text-[8px] text-gray-400 uppercase font-bold">{formatDate(w.issue_date)}</p>
                    <p className="text-[10px] text-gray-600 mt-1 line-clamp-1 italic">"{w.reason}"</p>
                  </div>
                  <div className="flex gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                    {!isReadOnly && (
                      <button 
                        onClick={(e) => { e.stopPropagation(); handleDeleteWarning(w.id); }} 
                        className="text-red-500 hover:opacity-80"
                      >
                        <Trash2 size={12} />
                      </button>
                    )}
                  </div>
                </div>
              ))
            )}
          </div>
        </DetailSection>

        {/* k. Status Exit / Pemberhentian */}
        <DetailSection 
          icon={LogOut} 
          title="Status Exit / Pemberhentian" 
          onAdd={!termination && !isInactive ? () => setShowTerminationForm(true) : undefined}
          isScrollable
        >
          {termination || isInactive ? (
            <div 
              className="space-y-3 p-3 bg-red-50/50 border border-red-100 rounded cursor-pointer hover:bg-red-50 transition-colors"
              onClick={() => termination && setSelectedTerminationDetail(termination)}
            >
              <div className="flex justify-between items-center">
                <span className="text-[10px] font-bold text-red-600 uppercase tracking-widest">{termination?.termination_type || 'KONTRAK BERAKHIR'}</span>
                <span className="text-[10px] font-bold text-gray-500">{formatDate(termination?.termination_date || account.end_date || '')}</span>
              </div>
              <DataRow label="Alasan Keluar" value={termination?.reason || 'Masa kontrak telah habis atau akun dinonaktifkan secara otomatis.'} />
              {termination?.termination_type === 'Pemecatan' && (
                <DataRow label="Uang Pesangon" value={formatCurrency(termination.severance_amount)} />
              )}
              {termination?.termination_type === 'Resign' && (
                <DataRow label="Biaya Penalti" value={formatCurrency(termination.penalty_amount)} />
              )}
              {!isReadOnly && (
                <button 
                  onClick={async (e) => {
                    e.stopPropagation();
                    const res = await Swal.fire({ title: 'Batalkan Pemberhentian?', text: 'Akun akan diaktifkan kembali.', icon: 'question', showCancelButton: true, confirmButtonColor: '#006E62' });
                    if (res.isConfirmed) {
                      setIsSaving(true);
                      if (termination) {
                        await disciplineService.deleteTermination(termination.id, id);
                      } else {
                        await accountService.update(id, { end_date: null });
                      }
                      setTermination(null);
                      setAccount(prev => prev ? { ...prev, end_date: null } : null);
                      setIsSaving(false);
                    }
                  }}
                  className="w-full mt-2 py-1.5 text-[10px] font-bold uppercase text-red-600 border border-red-200 rounded hover:bg-white transition-colors"
                >
                  Batalkan Exit / Aktifkan Kembali
                </button>
              )}
            </div>
          ) : (
            <div className="flex flex-col items-center justify-center py-6 text-gray-300">
               <LogOut size={32} strokeWidth={1} />
               <p className="text-[10px] font-bold uppercase tracking-widest mt-2 text-gray-400">Status: Aktif Bekerja</p>
            </div>
          )}
        </DetailSection>
      </div>

      {/* Modal Preview Media */}
      {previewMedia && (
        <div className="fixed inset-0 z-[200] bg-black/90 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white rounded-md max-w-4xl w-full overflow-hidden flex flex-col max-h-[90vh]">
             <div className="px-6 py-3 border-b flex justify-between items-center">
                <h4 className="text-sm font-bold text-[#006E62] uppercase tracking-widest">{previewMedia.title}</h4>
                <div className="flex gap-4 items-center">
                   {previewMedia.type === 'qr' && (
                     <button onClick={downloadQR} className="text-[#006E62] hover:text-[#005a50] flex items-center gap-1 text-xs font-bold"><Download size={14} /> DOWNLOAD</button>
                   )}
                   <button onClick={() => setPreviewMedia(null)} className="text-gray-400 hover:text-gray-600"><X size={20} /></button>
                </div>
             </div>
             <div className="flex-1 overflow-auto bg-gray-50 flex items-center justify-center p-8">
                {previewMedia.type === 'image' ? (
                  <img src={previewMedia.url} className="max-w-full max-h-full object-contain shadow-xl rounded" />
                ) : (
                  <div className="bg-white p-8 rounded-xl shadow-2xl">
                    <QRCodeSVG value={previewMedia.url} size={250} />
                    <p className="center mt-4 font-mono text-gray-400 text-xs">ID: {previewMedia.url}</p>
                  </div>
                )}
             </div>
          </div>
        </div>
      )}

      {showLogForm && (
        <LogForm type={showLogForm.type} accountId={id} initialData={showLogForm.data} isEdit={showLogForm.isEdit} onClose={() => setShowLogForm(null)} onSubmit={handleLogSubmit} />
      )}

      {showCertForm.show && (
        <CertificationFormModal onClose={() => setShowCertForm({ show: false })} onSuccess={() => { setShowCertForm({ show: false }); fetchData(); }} initialData={showCertForm.data} />
      )}

      {showContractForm.show && (
        <ContractFormModal onClose={() => setShowContractForm({ show: false })} onSuccess={() => { setShowContractForm({ show: false }); fetchData(); }} initialData={showContractForm.data} />
      )}

      {showWarningForm && (
        <WarningForm accountId={id} onClose={() => setShowWarningForm(false)} onSuccess={() => { setShowWarningForm(false); fetchData(); }} />
      )}
      
      {showTerminationForm && (
        <TerminationForm accountId={id} onClose={() => setShowTerminationForm(false)} onSuccess={() => { setShowTerminationForm(false); fetchData(); }} />
      )}

      {selectedContractDetail && (
        <ContractDetailModal contract={selectedContractDetail} onClose={() => setSelectedContractDetail(null)} />
      )}

      {selectedCareerDetail && (
        <CareerDetailModal log={selectedCareerDetail} onClose={() => setSelectedCareerDetail(null)} />
      )}

      {selectedHealthDetail && (
        <HealthDetailModal log={selectedHealthDetail} onClose={() => setSelectedHealthDetail(null)} />
      )}

      {selectedCertDetail && (
        <CertificationDetailModal cert={selectedCertDetail} onClose={() => setSelectedCertDetail(null)} />
      )}

      {selectedWarningDetail && (
        <WarningDetailModal log={selectedWarningDetail} onClose={() => setSelectedWarningDetail(null)} />
      )}

      {selectedTerminationDetail && (
        <TerminationDetailModal log={selectedTerminationDetail} onClose={() => setSelectedTerminationDetail(null)} />
      )}
    </div>
  );
};

export default AccountDetail;
