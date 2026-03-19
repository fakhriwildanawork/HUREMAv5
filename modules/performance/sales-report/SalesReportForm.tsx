import React, { useState, useEffect, useRef } from 'react';
import { X, Save, MapPin, Camera, Info, Loader2, Link as LinkIcon, Trash2, Plus, FileText, Upload } from 'lucide-react';
import { MapContainer, TileLayer, Marker, useMap } from 'react-leaflet';
import L from 'leaflet';
import 'leaflet/dist/leaflet.css';
import { SalesReportInput } from '../../../types';
import { googleDriveService } from '../../../services/googleDriveService';

// Fix Leaflet marker icon issue
const DefaultIcon = L.icon({
  iconUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon.png',
  shadowUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-shadow.png',
  iconSize: [25, 41],
  iconAnchor: [12, 41]
});

L.Marker.prototype.options.icon = DefaultIcon;

// Helper component to update map view
const ChangeView = ({ center }: { center: [number, number] }) => {
  const map = useMap();
  map.setView(center, 15);
  return null;
};

interface SalesReportFormProps {
  onClose: () => void;
  onSubmit: (data: SalesReportInput) => void;
}

const SalesReportForm: React.FC<SalesReportFormProps> = ({ onClose, onSubmit }) => {
  const [formData, setFormData] = useState<SalesReportInput>({
    account_id: '',
    customer_name: '',
    activity_type: 'Cold Call',
    description: '',
    latitude: 0,
    longitude: 0,
    address: '',
    photo_urls: [],
    file_ids: []
  });

  const [isLocating, setIsLocating] = useState(false);
  const [isGeocoding, setIsGeocoding] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [locationError, setLocationError] = useState<string | null>(null);
  const [newLink, setNewLink] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);
  const docInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    handleGetLocation();
  }, []);

  const reverseGeocode = async (lat: number, lon: number) => {
    try {
      setIsGeocoding(true);
      const response = await fetch(`https://nominatim.openstreetmap.org/reverse?format=json&lat=${lat}&lon=${lon}&zoom=18&addressdetails=1`);
      const data = await response.json();
      if (data && data.display_name) {
        setFormData(prev => ({ ...prev, address: data.display_name }));
      }
    } catch (error) {
      console.error('Geocoding error:', error);
    } finally {
      setIsGeocoding(false);
    }
  };

  const handleGetLocation = () => {
    if (!navigator.geolocation) {
      setLocationError('Geolocation tidak didukung oleh browser Anda.');
      return;
    }

    setIsLocating(true);
    setLocationError(null);

    navigator.geolocation.getCurrentPosition(
      (position) => {
        const { latitude, longitude } = position.coords;
        setFormData(prev => ({
          ...prev,
          latitude,
          longitude
        }));
        setIsLocating(false);
        reverseGeocode(latitude, longitude);
      },
      (error) => {
        console.error('Geolocation error:', error);
        let msg = 'Gagal mendapatkan lokasi.';
        if (error.code === error.PERMISSION_DENIED) msg = 'Izin lokasi ditolak. Mohon aktifkan GPS.';
        setLocationError(msg);
        setIsLocating(false);
      },
      { enableHighAccuracy: true, timeout: 10000, maximumAge: 0 }
    );
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => {
    const { name, value } = e.target;
    setFormData(prev => ({ ...prev, [name]: value }));
  };

  const handleTakePhoto = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      setIsUploading(true);
      const fileId = await googleDriveService.uploadFile(file);
      const url = googleDriveService.getFileUrl(fileId);
      setFormData(prev => ({ ...prev, photo_urls: [...prev.photo_urls, url] }));
    } catch (error) {
      alert('Gagal mengunggah foto.');
    } finally {
      setIsUploading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const handleAttachDoc = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      setIsUploading(true);
      const fileId = await googleDriveService.uploadFile(file);
      setFormData(prev => ({ ...prev, file_ids: [...prev.file_ids, fileId] }));
    } catch (error) {
      alert('Gagal mengunggah dokumen.');
    } finally {
      setIsUploading(false);
      if (docInputRef.current) docInputRef.current.value = '';
    }
  };

  const addLink = () => {
    if (newLink && !formData.photo_urls.includes(newLink)) {
      setFormData(prev => ({ ...prev, photo_urls: [...prev.photo_urls, newLink] }));
      setNewLink('');
    }
  };

  const removePhoto = (url: string) => {
    setFormData(prev => ({ ...prev, photo_urls: prev.photo_urls.filter(u => u !== url) }));
  };

  const removeFile = (id: string) => {
    setFormData(prev => ({ ...prev, file_ids: prev.file_ids.filter(f => f !== id) }));
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (formData.latitude === 0 || formData.longitude === 0) {
      alert('Lokasi GPS wajib didapatkan sebelum mengirim laporan.');
      return;
    }
    onSubmit(formData);
  };

  const mapPosition: [number, number] = [formData.latitude || -6.200000, formData.longitude || 106.816666];

  return (
    <div className="fixed inset-0 z-[70] flex items-center justify-center bg-black/40 p-4 backdrop-blur-sm">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in duration-200 flex flex-col max-h-[90vh]">
        <div className="px-6 py-4 border-b border-gray-50 flex items-center justify-between bg-gray-50/50 shrink-0">
          <div className="flex items-center gap-2">
            <MapPin size={18} className="text-[#006E62]" />
            <h3 className="text-sm font-bold text-gray-800 uppercase tracking-tight">Check-in Kunjungan Sales</h3>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={20} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-6 space-y-5 overflow-y-auto scrollbar-thin">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="space-y-5">
              {/* Geolocation Status */}
              <div className={`p-4 rounded-2xl border flex flex-col gap-3 ${locationError ? 'bg-rose-50 border-rose-100 text-rose-600' : 'bg-emerald-50 border-emerald-100 text-[#006E62]'}`}>
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-3">
                    {isLocating ? (
                      <Loader2 size={20} className="animate-spin" />
                    ) : locationError ? (
                      <Info size={20} />
                    ) : (
                      <MapPin size={20} />
                    )}
                    <div>
                      <p className="text-[10px] font-bold uppercase tracking-widest mb-0.5">Status Lokasi GPS</p>
                      <p className="text-xs font-bold">
                        {isLocating ? 'Mencari Koordinat...' : locationError ? locationError : `${formData.latitude.toFixed(6)}, ${formData.longitude.toFixed(6)}`}
                      </p>
                    </div>
                  </div>
                  {!isLocating && (
                    <button 
                      type="button" 
                      onClick={handleGetLocation}
                      className="text-[10px] font-bold uppercase underline tracking-widest"
                    >
                      Refresh
                    </button>
                  )}
                </div>
                
                {formData.address && (
                  <div className="pt-2 border-t border-emerald-100/50">
                    <p className="text-[9px] font-bold uppercase tracking-widest mb-1 opacity-60">Alamat Terdeteksi</p>
                    <p className="text-[10px] font-medium leading-relaxed">{formData.address}</p>
                  </div>
                )}
              </div>

              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Nama Klien / Toko</label>
                <input 
                  required
                  type="text" 
                  name="customer_name" 
                  value={formData.customer_name} 
                  onChange={handleChange}
                  placeholder="Contoh: Toko Berkah Jaya"
                  className="w-full px-3 py-2.5 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] text-xs font-bold"
                />
              </div>

              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Tahap Pipeline</label>
                <select
                  required
                  name="activity_type"
                  value={formData.activity_type}
                  onChange={handleChange}
                  className="w-full px-3 py-2.5 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] text-xs font-bold appearance-none"
                >
                  <option value="Cold Call">Cold Call (Perkenalan)</option>
                  <option value="Site Survey">Site Survey (Kunjungan Lokasi)</option>
                  <option value="Product Demo">Product Demo (Presentasi)</option>
                  <option value="Offering">Offering (Penawaran Harga)</option>
                  <option value="Negotiation">Negotiation (Negosiasi)</option>
                  <option value="Closing">Closing (Deal/Kontrak)</option>
                  <option value="Maintenance">Maintenance (Kunjungan Rutin)</option>
                </select>
              </div>

              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Catatan Kunjungan</label>
                <textarea 
                  required
                  name="description" 
                  value={formData.description} 
                  onChange={handleChange}
                  rows={3}
                  placeholder="Tuliskan hasil pertemuan atau kendala di lapangan..."
                  className="w-full px-3 py-2.5 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] text-xs font-medium resize-none"
                />
              </div>
            </div>

            <div className="space-y-5">
              {/* Map View */}
              <div className="h-[200px] rounded-2xl overflow-hidden border border-gray-100 shadow-inner relative z-0">
                <MapContainer center={mapPosition} zoom={15} scrollWheelZoom={false} style={{ height: '100%', width: '100%' }}>
                  <TileLayer
                    attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
                    url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
                  />
                  <Marker position={mapPosition} />
                  <ChangeView center={mapPosition} />
                </MapContainer>
              </div>

              {/* Photo Evidence */}
              <div className="space-y-3">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Bukti Foto & Dokumen</label>
                <div className="grid grid-cols-2 gap-3">
                  <button 
                    type="button" 
                    onClick={() => fileInputRef.current?.click()}
                    disabled={isUploading}
                    className="flex flex-col items-center justify-center gap-2 p-4 border-2 border-dashed border-gray-100 rounded-2xl hover:border-[#006E62] hover:bg-emerald-50 transition-all group"
                  >
                    <Camera size={24} className="text-gray-300 group-hover:text-[#006E62]" />
                    <span className="text-[10px] font-bold text-gray-400 group-hover:text-[#006E62] uppercase">Ambil Foto</span>
                    <input 
                      ref={fileInputRef}
                      type="file" 
                      accept="image/*" 
                      capture="environment" 
                      onChange={handleTakePhoto}
                      className="hidden"
                    />
                  </button>
                  <button 
                    type="button" 
                    onClick={() => docInputRef.current?.click()}
                    disabled={isUploading}
                    className="flex flex-col items-center justify-center gap-2 p-4 border-2 border-dashed border-gray-100 rounded-2xl hover:border-[#006E62] hover:bg-emerald-50 transition-all group"
                  >
                    <Upload size={24} className="text-gray-300 group-hover:text-[#006E62]" />
                    <span className="text-[10px] font-bold text-gray-400 group-hover:text-[#006E62] uppercase">Attach File</span>
                    <input 
                      ref={docInputRef}
                      type="file" 
                      onChange={handleAttachDoc}
                      className="hidden"
                    />
                  </button>
                </div>

                {isUploading && (
                  <div className="flex items-center gap-2 px-3 py-2 bg-blue-50 text-blue-600 rounded-xl animate-pulse">
                    <Loader2 size={14} className="animate-spin" />
                    <span className="text-[10px] font-bold uppercase tracking-widest">Mengunggah...</span>
                  </div>
                )}

                {/* Photo List */}
                {formData.photo_urls.length > 0 && (
                  <div className="space-y-2">
                    {formData.photo_urls.map(url => (
                      <div key={url} className="flex items-center justify-between px-3 py-2 bg-gray-50 border border-gray-100 rounded-xl">
                        <div className="flex items-center gap-2 overflow-hidden">
                          <Camera size={12} className="text-gray-400 shrink-0" />
                          <span className="text-[10px] text-gray-600 truncate">{url}</span>
                        </div>
                        <button type="button" onClick={() => removePhoto(url)} className="text-rose-500 hover:text-rose-700">
                          <Trash2 size={14} />
                        </button>
                      </div>
                    ))}
                  </div>
                )}

                {/* File List */}
                {formData.file_ids.length > 0 && (
                  <div className="space-y-2">
                    {formData.file_ids.map(id => (
                      <div key={id} className="flex items-center justify-between px-3 py-2 bg-blue-50/50 border border-blue-100 rounded-xl">
                        <div className="flex items-center gap-2 overflow-hidden">
                          <FileText size={12} className="text-blue-400 shrink-0" />
                          <span className="text-[10px] text-blue-600 truncate">Doc ID: {id}</span>
                        </div>
                        <button type="button" onClick={() => removeFile(id)} className="text-rose-500 hover:text-rose-700">
                          <Trash2 size={14} />
                        </button>
                      </div>
                    ))}
                  </div>
                )}

                <div className="space-y-1.5">
                  <label className="text-[9px] font-bold text-gray-400 uppercase tracking-widest px-1">Link Eksternal (Opsional)</label>
                  <div className="flex gap-2">
                    <input 
                      type="url" 
                      value={newLink}
                      onChange={(e) => setNewLink(e.target.value)}
                      placeholder="https://..."
                      className="flex-1 px-3 py-2 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] text-[10px]"
                    />
                    <button 
                      type="button" 
                      onClick={addLink}
                      className="px-3 py-2 bg-[#006E62] text-white rounded-xl hover:bg-[#005a50] transition-all"
                    >
                      <Plus size={16} />
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </form>

        <div className="px-6 py-4 border-t border-gray-50 bg-gray-50/30 flex gap-3 shrink-0">
          <button 
            type="button" 
            onClick={onClose}
            className="flex-1 py-3 border border-gray-100 text-gray-400 bg-white rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-gray-50 transition-all"
          >
            Batal
          </button>
          <button 
            onClick={handleSubmit}
            disabled={isLocating || isUploading}
            className="flex-[2] py-3 bg-[#006E62] text-white rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-[#005a50] shadow-lg shadow-[#006E62]/20 transition-all flex items-center justify-center gap-2 disabled:opacity-50"
          >
            <Save size={16} />
            Kirim Laporan Kunjungan
          </button>
        </div>
      </div>
    </div>
  );
};

export default SalesReportForm;
