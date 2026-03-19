
import React, { useState, useEffect } from 'react';
import { X, Edit2, Trash2, MapPin, Calendar, Clock, Image as ImageIcon, Phone, Target, Layers, Plus, FileText, ExternalLink, Paperclip } from 'lucide-react';
import Swal from 'sweetalert2';
import { Location, LocationAdministration } from '../../types';
import { locationService } from '../../services/locationService';
import { googleDriveService } from '../../services/googleDriveService';
import LocationAdminForm from './LocationAdminForm';
import LoadingSpinner from '../../components/Common/LoadingSpinner';

interface LocationDetailProps {
  id: string;
  onClose: () => void;
  onEdit: (location: Location) => void;
  onDelete: (id: string) => void;
}

const LocationDetail: React.FC<LocationDetailProps> = ({ id, onClose, onEdit, onDelete }) => {
  const [location, setLocation] = useState<Location | null>(null);
  const [administrations, setAdministrations] = useState<LocationAdministration[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [showAdminForm, setShowAdminForm] = useState(false);
  const [editingAdmin, setEditingAdmin] = useState<LocationAdministration | null>(null);

  useEffect(() => {
    fetchData();
  }, [id]);

  const fetchData = async () => {
    try {
      setIsLoading(true);
      const [locData, adminData] = await Promise.all([
        locationService.getById(id),
        locationService.getAdministrations(id)
      ]);
      setLocation(locData);
      setAdministrations(adminData);
    } catch (error) {
      console.error(error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleSaveAdmin = async (adminInput: any) => {
    setIsSaving(true);
    try {
      // Pastikan data bersih: ubah string kosong ke null untuk konsistensi database
      const cleanInput = {
        ...adminInput,
        due_date: adminInput.due_date === '' ? null : adminInput.due_date
      };

      const newAdmin = await locationService.createAdministration({
        ...cleanInput,
        location_id: id
      });
      setAdministrations(prev => [newAdmin, ...prev]);
      
      setShowAdminForm(false);
      setEditingAdmin(null);
      Swal.fire({
        title: 'Berhasil!',
        text: 'Data administrasi telah disimpan.',
        icon: 'success',
        timer: 1500,
        showConfirmButton: false
      });
    } catch (error) {
      Swal.fire('Gagal', 'Terjadi kesalahan saat menyimpan data administrasi', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleDeleteAdmin = async (adminId: string) => {
    const result = await Swal.fire({
      title: 'Hapus data administrasi?',
      text: "Data ini akan dihapus permanen.",
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
        await locationService.deleteAdministration(adminId);
        setAdministrations(prev => prev.filter(a => a.id !== adminId));
        Swal.fire('Terhapus!', 'Data administrasi telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Gagal menghapus data', 'error');
      } finally {
        setIsSaving(false);
      }
    }
  };

  if (isLoading) return (
    <div className="flex items-center justify-center h-64">
      <div className="w-8 h-8 border-4 border-[#006E62] border-t-transparent rounded-full animate-spin"></div>
    </div>
  );
  if (!location) return null;

  const formatDate = (dateStr?: string) => {
    if (!dateStr) return '-';
    return new Date(dateStr).toLocaleDateString('id-ID', {
      day: 'numeric',
      month: 'long',
      year: 'numeric'
    });
  };

  return (
    <div className="bg-white rounded-md shadow-sm border border-gray-100 overflow-hidden flex flex-col md:flex-row gap-8 p-6 md:p-8">
      {isSaving && <LoadingSpinner />}
      
      {/* Info Utama */}
      <div className="w-full md:w-1/3 space-y-6">
        <div className="space-y-4">
          <div className="flex justify-between items-start">
            <div>
              <div className="flex items-center gap-2 mb-1">
                <Layers size={14} className="text-[#006E62]" />
                <span className="text-[10px] font-bold uppercase tracking-widest text-[#006E62]">{location.location_type || 'Lokasi'}</span>
              </div>
              <h2 className="text-2xl font-bold text-gray-800 tracking-tight">{location.name}</h2>
            </div>
            <div className="flex gap-2">
              <button onClick={() => onEdit(location)} className="p-2 text-gray-400 hover:text-[#006E62] border border-gray-100 rounded hover:bg-gray-50 transition-all shadow-sm">
                <Edit2 size={16} />
              </button>
              <button onClick={() => onDelete(location.id)} className="p-2 text-gray-400 hover:text-red-500 border border-gray-100 rounded hover:bg-gray-50 transition-all shadow-sm">
                <Trash2 size={16} />
              </button>
            </div>
          </div>

          <div className="grid grid-cols-1 gap-4">
            <div className="space-y-1">
              <p className="text-[9px] font-bold uppercase tracking-widest text-gray-400">Alamat Lengkap</p>
              <div className="flex items-start gap-2">
                <MapPin size={12} className="text-gray-400 mt-0.5 shrink-0" />
                <p className="text-xs text-gray-600 leading-relaxed font-medium">
                  {location.address}, {location.city}, {location.province} {location.zip_code}
                </p>
              </div>
            </div>
            <div className="flex justify-between items-center bg-gray-50 p-3 rounded-md border border-gray-100">
              <div className="space-y-1">
                <p className="text-[9px] font-bold uppercase tracking-widest text-gray-400">Telepon</p>
                <div className="flex items-center gap-2 text-xs font-semibold text-gray-700">
                  <Phone size={12} className="text-[#006E62]" /> {location.phone || '-'}
                </div>
              </div>
              <div className="space-y-1 text-right">
                <p className="text-[9px] font-bold uppercase tracking-widest text-gray-400">Radius</p>
                <div className="flex items-center gap-2 text-xs font-semibold text-gray-700 justify-end">
                  <Target size={12} className="text-[#006E62]" /> {location.radius}m
                </div>
              </div>
            </div>
            {location.description && (
              <div className="space-y-1 pt-2">
                <p className="text-[9px] font-bold uppercase tracking-widest text-gray-400">Keterangan</p>
                <p className="text-xs text-gray-600 italic leading-relaxed bg-gray-50 p-3 rounded border border-gray-100">
                  "{location.description}"
                </p>
              </div>
            )}
          </div>
          
          <div className="pt-4 border-t border-gray-100 flex justify-between text-[9px] text-gray-400 uppercase font-bold">
             <span>Dibuat: {formatDate(location.created_at)}</span>
             <span>ID: {location.id.slice(0, 8)}</span>
          </div>
        </div>
      </div>

      {/* Kolom Kanan: Data Administrasi */}
      <div className="flex-1 space-y-6">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="p-2 bg-[#006E62]/10 rounded-md">
              <FileText size={20} className="text-[#006E62]" />
            </div>
            <div>
              <h3 className="text-lg font-bold text-gray-800">Data Administrasi</h3>
              <p className="text-[10px] text-gray-400 uppercase font-bold tracking-widest">Legalitas & Dokumen</p>
            </div>
          </div>
          <button 
            onClick={() => setShowAdminForm(true)}
            className="flex items-center gap-2 bg-[#006E62] text-white px-4 py-2 rounded-md hover:bg-[#005a50] transition-colors shadow-sm text-xs font-bold uppercase"
          >
            <Plus size={14} /> Tambah Data
          </button>
        </div>

        {administrations.length === 0 ? (
          <div className="flex flex-col items-center justify-center py-12 bg-gray-50 rounded-md border border-dashed border-gray-200 text-gray-400">
            <FileText size={40} strokeWidth={1} className="mb-2" />
            <p className="text-xs font-medium uppercase tracking-widest">Belum ada data administrasi</p>
          </div>
        ) : (
          <div className="space-y-4">
            {administrations.map((admin) => (
              <div key={admin.id} className="bg-white border border-gray-100 rounded-md p-4 shadow-sm hover:shadow-md transition-shadow relative group">
                <div className="absolute top-4 right-4 flex gap-2">
                  <button 
                    onClick={() => handleDeleteAdmin(admin.id)}
                    className="text-gray-300 hover:text-red-500 transition-colors"
                  >
                    <Trash2 size={14} />
                  </button>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                  <div className="space-y-1">
                    <p className="text-[9px] font-bold uppercase tracking-widest text-gray-400">Tanggal</p>
                    <p className="text-xs font-bold text-gray-700">{formatDate(admin.admin_date)}</p>
                  </div>
                  <div className="space-y-1">
                    <p className="text-[9px] font-bold uppercase tracking-widest text-gray-400">Status</p>
                    <span className={`inline-block px-2 py-0.5 rounded text-[9px] font-bold uppercase ${
                      admin.status === 'Milik Sendiri' ? 'bg-green-50 text-green-600' :
                      admin.status === 'Sewa/Kontrak' ? 'bg-blue-50 text-blue-600' : 'bg-orange-50 text-orange-600'
                    }`}>
                      {admin.status}
                    </span>
                  </div>
                  {admin.due_date && (
                    <div className="space-y-1">
                      <p className="text-[9px] font-bold uppercase tracking-widest text-gray-400">Jatuh Tempo</p>
                      <p className="text-xs font-bold text-red-500">{formatDate(admin.due_date)}</p>
                    </div>
                  )}
                  <div className="space-y-1 lg:col-span-1">
                    <p className="text-[9px] font-bold uppercase tracking-widest text-gray-400">Lampiran</p>
                    <div className="flex flex-wrap gap-1">
                      {admin.file_ids?.map((fid, idx) => (
                        <a 
                          key={idx} 
                          href={googleDriveService.getFileUrl(fid)} 
                          target="_blank" 
                          rel="noreferrer"
                          className="flex items-center gap-1 bg-gray-100 text-[9px] font-bold p-1 rounded hover:bg-gray-200 transition-colors"
                        >
                          <Paperclip size={10} /> DOK {idx + 1}
                        </a>
                      ))}
                      {(!admin.file_ids || admin.file_ids.length === 0) && <span className="text-xs text-gray-300">-</span>}
                    </div>
                  </div>
                </div>
                {admin.description && (
                  <div className="mt-3 pt-3 border-t border-gray-50">
                    <p className="text-[9px] font-bold uppercase tracking-widest text-gray-400 mb-1">Keterangan</p>
                    <p className="text-xs text-gray-600 leading-relaxed">{admin.description}</p>
                  </div>
                )}
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Modal Form Administrasi */}
      {showAdminForm && (
        <LocationAdminForm 
          initialData={editingAdmin}
          onClose={() => {
            setShowAdminForm(false);
            setEditingAdmin(null);
          }}
          onSubmit={handleSaveAdmin}
        />
      )}
    </div>
  );
};

export default LocationDetail;
