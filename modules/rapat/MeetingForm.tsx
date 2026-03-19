import React, { useState, useRef, useEffect } from 'react';
import { X, Save, Calendar, Clock, MapPin, Users, Plus, Trash2, Search, Video, Globe, Map as MapIcon, Loader2, Info } from 'lucide-react';
import { MeetingInput, Account } from '../../types';
import { MapContainer, TileLayer, Marker, useMap } from 'react-leaflet';
import L from 'leaflet';
import 'leaflet/dist/leaflet.css';

// Fix Leaflet marker icon issue
const DefaultIcon = L.icon({
  iconUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon.png',
  shadowUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-shadow.png',
  iconSize: [25, 41],
  iconAnchor: [12, 41]
});
L.Marker.prototype.options.icon = DefaultIcon;

const ChangeView = ({ center }: { center: [number, number] }) => {
  const map = useMap();
  map.setView(center, 15);
  return null;
};

interface MeetingFormProps {
  onClose: () => void;
  onSubmit: (data: any) => void;
  accounts: Account[];
}

const MeetingForm: React.FC<MeetingFormProps> = ({ onClose, onSubmit, accounts }) => {
  const [formData, setFormData] = useState<any>({
    title: '',
    description: '',
    scheduled_at: '',
    location_type: 'Online',
    location_detail: '',
    latitude: -6.200000,
    longitude: 106.816666,
    participant_ids: [],
    notulen_ids: [],
    attachments: [],
    links: []
  });

  const [accountSearch, setAccountSearch] = useState('');
  const [isLocating, setIsLocating] = useState(false);
  const [isGeocoding, setIsGeocoding] = useState(false);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => {
    const { name, value } = e.target;
    setFormData(prev => ({ ...prev, [name]: value }));
  };

  const toggleSelection = (id: string, field: 'participant_ids' | 'notulen_ids') => {
    setFormData(prev => ({
      ...prev,
      [field]: prev[field].includes(id) 
        ? prev[field].filter(i => i !== id) 
        : [...prev[field], id]
    }));
  };

  const handleGetLocation = () => {
    if (!navigator.geolocation) return;
    setIsLocating(true);
    navigator.geolocation.getCurrentPosition(
      (position) => {
        const { latitude, longitude } = position.coords;
        setFormData(prev => ({ ...prev, latitude, longitude }));
        setIsLocating(false);
        reverseGeocode(latitude, longitude);
      },
      () => setIsLocating(false),
      { enableHighAccuracy: true }
    );
  };

  const reverseGeocode = async (lat: number, lon: number) => {
    try {
      setIsGeocoding(true);
      const response = await fetch(`https://nominatim.openstreetmap.org/reverse?format=json&lat=${lat}&lon=${lon}&zoom=18`);
      const data = await response.json();
      if (data && data.display_name) {
        setFormData(prev => ({ ...prev, location_detail: data.display_name }));
      }
    } catch (error) {
      console.error(error);
    } finally {
      setIsGeocoding(false);
    }
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (formData.participant_ids.length === 0) {
      alert('Pilih setidaknya satu peserta rapat.');
      return;
    }
    if (formData.notulen_ids.length === 0) {
      alert('Pilih setidaknya satu notulen rapat.');
      return;
    }
    onSubmit(formData);
  };

  const filteredAccounts = accounts.filter(a => 
    a.full_name.toLowerCase().includes(accountSearch.toLowerCase()) ||
    a.internal_nik.toLowerCase().includes(accountSearch.toLowerCase())
  );

  const mapPosition: [number, number] = [formData.latitude, formData.longitude];

  const handleMarkerDragEnd = (e: any) => {
    const marker = e.target;
    const position = marker.getLatLng();
    setFormData(prev => ({ ...prev, latitude: position.lat, longitude: position.lng }));
    reverseGeocode(position.lat, position.lng);
  };

  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
      <div className="bg-white w-full max-w-4xl rounded-3xl overflow-hidden shadow-2xl flex flex-col animate-in fade-in zoom-in duration-300 max-h-[95vh]">
        <div className="p-6 border-b border-gray-100 flex items-center justify-between shrink-0 bg-gray-50/50">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-[#006E62] text-white flex items-center justify-center">
              <Video size={20} />
            </div>
            <div>
              <h3 className="text-lg font-bold text-gray-800 tracking-tight">Jadwalkan Rapat Baru</h3>
              <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Atur waktu, lokasi, dan peserta rapat</p>
            </div>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-gray-100 rounded-xl transition-colors text-gray-400">
            <X size={24} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-6 overflow-y-auto space-y-6 scrollbar-none">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
            {/* Left Column: Basic Info & Location */}
            <div className="space-y-6">
              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Judul Rapat</label>
                <input 
                  required
                  type="text" 
                  name="title" 
                  value={formData.title} 
                  onChange={handleChange}
                  placeholder="Contoh: Rapat Koordinasi Mingguan"
                  className="w-full px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] transition-all text-sm font-bold text-gray-700"
                />
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1.5">
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Waktu Rapat</label>
                  <input 
                    required
                    type="datetime-local" 
                    name="scheduled_at" 
                    value={formData.scheduled_at} 
                    onChange={handleChange}
                    className="w-full px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] transition-all text-sm font-bold text-gray-700"
                  />
                </div>
                <div className="space-y-1.5">
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Tipe Lokasi</label>
                  <select 
                    name="location_type" 
                    value={formData.location_type} 
                    onChange={handleChange}
                    className="w-full px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] transition-all text-sm font-bold text-gray-700"
                  >
                    <option value="Online">Online Meeting</option>
                    <option value="Offline">Offline (Tatap Muka)</option>
                  </select>
                </div>
              </div>

              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">
                  {formData.location_type === 'Online' ? 'Link Meeting' : 'Lokasi / Alamat'}
                </label>
                <div className="relative">
                  <input 
                    required
                    type="text" 
                    name="location_detail" 
                    value={formData.location_detail} 
                    onChange={handleChange}
                    placeholder={formData.location_type === 'Online' ? 'https://zoom.us/j/...' : 'Gedung A, Lantai 3'}
                    className="w-full px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] transition-all text-sm font-bold text-gray-700"
                  />
                  {formData.location_type === 'Offline' && (
                    <button 
                      type="button"
                      onClick={handleGetLocation}
                      className="absolute right-3 top-1/2 -translate-y-1/2 text-[#006E62] hover:bg-emerald-50 p-1.5 rounded-lg transition-all"
                      title="Gunakan Lokasi Saat Ini"
                    >
                      {isLocating ? <Loader2 size={16} className="animate-spin" /> : <MapIcon size={16} />}
                    </button>
                  )}
                </div>
              </div>

              {formData.location_type === 'Offline' && (
                <div className="h-[200px] rounded-2xl overflow-hidden border border-gray-100 shadow-inner relative z-0">
                  <MapContainer center={mapPosition} zoom={15} scrollWheelZoom={false} style={{ height: '100%', width: '100%' }}>
                    <TileLayer
                      attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
                      url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
                    />
                    <Marker 
                      position={mapPosition} 
                      draggable={true}
                      eventHandlers={{
                        dragend: handleMarkerDragEnd,
                      }}
                    />
                    <ChangeView center={mapPosition} />
                  </MapContainer>
                </div>
              )}

              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Agenda / Deskripsi</label>
                <textarea 
                  required
                  name="description" 
                  value={formData.description} 
                  onChange={handleChange}
                  rows={3}
                  placeholder="Tuliskan poin-poin pembahasan rapat..."
                  className="w-full px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] transition-all text-sm font-medium text-gray-700 resize-none"
                />
              </div>
            </div>

            {/* Right Column: Participants & Notulens */}
            <div className="space-y-6">
              <div className="space-y-1.5">
                <div className="flex items-center justify-between px-1">
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Pilih Peserta & Notulen</label>
                  <div className="relative w-40">
                    <Search className="absolute left-2 top-1/2 -translate-y-1/2 text-gray-400" size={12} />
                    <input 
                      type="text"
                      value={accountSearch}
                      onChange={(e) => setAccountSearch(e.target.value)}
                      placeholder="Cari..."
                      className="w-full pl-7 pr-3 py-1 bg-gray-50 border border-gray-100 rounded-lg focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 text-[10px] font-bold"
                    />
                  </div>
                </div>
                
                <div className="bg-gray-50 border border-gray-100 rounded-2xl overflow-hidden h-[500px] overflow-y-auto scrollbar-thin">
                  <table className="w-full text-left border-collapse">
                    <thead className="sticky top-0 bg-gray-100 z-10">
                      <tr>
                        <th className="px-4 py-2 text-[9px] font-bold text-gray-400 uppercase">Pegawai</th>
                        <th className="px-4 py-2 text-[9px] font-bold text-gray-400 uppercase text-center">Peserta</th>
                        <th className="px-4 py-2 text-[9px] font-bold text-gray-400 uppercase text-center">Notulen</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100">
                      {filteredAccounts.map(account => (
                        <tr key={account.id} className="hover:bg-white transition-colors">
                          <td className="px-4 py-2">
                            <div className="flex items-center gap-2">
                              <div className="w-6 h-6 rounded-full bg-gray-200 flex items-center justify-center text-[10px] font-bold text-gray-500">
                                {account.full_name.charAt(0)}
                              </div>
                              <div className="overflow-hidden">
                                <p className="text-[10px] font-bold text-gray-700 truncate">{account.full_name}</p>
                                <p className="text-[8px] text-gray-400 font-medium">{account.internal_nik}</p>
                              </div>
                            </div>
                          </td>
                          <td className="px-4 py-2 text-center">
                            <input 
                              type="checkbox" 
                              checked={formData.participant_ids.includes(account.id)}
                              onChange={() => toggleSelection(account.id, 'participant_ids')}
                              className="w-4 h-4 rounded border-gray-300 text-[#006E62] focus:ring-[#006E62]"
                            />
                          </td>
                          <td className="px-4 py-2 text-center">
                            <input 
                              type="checkbox" 
                              checked={formData.notulen_ids.includes(account.id)}
                              onChange={() => toggleSelection(account.id, 'notulen_ids')}
                              className="w-4 h-4 rounded border-gray-300 text-[#006E62] focus:ring-[#006E62]"
                            />
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <div className="flex items-center gap-4 px-2 pt-2">
                  <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest">
                    Terpilih: <span className="text-[#006E62]">{formData.participant_ids.length} Peserta</span>
                  </p>
                  <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest">
                    Notulen: <span className="text-amber-600">{formData.notulen_ids.length} Orang</span>
                  </p>
                </div>
              </div>
            </div>
          </div>
        </form>

        <div className="p-6 border-t border-gray-100 bg-gray-50/50 flex gap-3 shrink-0">
          <button 
            type="button"
            onClick={onClose}
            className="flex-1 py-3 border border-gray-200 text-gray-600 rounded-xl font-bold text-xs uppercase tracking-widest hover:bg-white transition-all"
          >
            Batal
          </button>
          <button 
            onClick={handleSubmit}
            className="flex-[2] py-3 bg-[#006E62] text-white rounded-xl font-bold text-xs uppercase tracking-widest hover:bg-[#005a50] transition-all shadow-lg shadow-[#006E62]/20 flex items-center justify-center gap-2"
          >
            <Save size={16} />
            Jadwalkan Rapat
          </button>
        </div>
      </div>
    </div>
  );
};

export default MeetingForm;
