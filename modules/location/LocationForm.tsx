
import React, { useState, useEffect, useRef } from 'react';
import { X, Save, MapPin, Navigation, AlertCircle } from 'lucide-react';
import Swal from 'sweetalert2';
import { LocationInput } from '../../types';
import { googleDriveService } from '../../services/googleDriveService';

declare var L: any;

interface LocationFormProps {
  onClose: () => void;
  onSubmit: (data: LocationInput) => void;
  initialData?: Partial<LocationInput>;
}

const LocationForm: React.FC<LocationFormProps> = ({ onClose, onSubmit, initialData }) => {
  const [formData, setFormData] = useState<LocationInput>({
    name: initialData?.name || '',
    location_type: initialData?.location_type || 'Kantor',
    address: initialData?.address || '',
    city: initialData?.city || '',
    province: initialData?.province || '',
    zip_code: initialData?.zip_code || '',
    phone: initialData?.phone || '',
    latitude: initialData?.latitude || 0,
    longitude: initialData?.longitude || 0,
    radius: initialData?.radius || 100,
    description: initialData?.description || '',
    image_google_id: initialData?.image_google_id || '',
  });

  const mapRef = useRef<any>(null);
  const markerRef = useRef<any>(null);

  useEffect(() => {
    const defaultLat = formData.latitude || -6.200000;
    const defaultLng = formData.longitude || 106.816666;

    if (!mapRef.current) {
      mapRef.current = L.map('map-container').setView([defaultLat, defaultLng], 15);
      L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        attribution: '© OpenStreetMap contributors'
      }).addTo(mapRef.current);

      markerRef.current = L.marker([defaultLat, defaultLng], { draggable: true }).addTo(mapRef.current);

      markerRef.current.on('dragend', function (event: any) {
        const position = event.target.getLatLng();
        updateCoords(position.lat, position.lng);
      });

      mapRef.current.on('click', function (e: any) {
        const { lat, lng } = e.latlng;
        markerRef.current.setLatLng([lat, lng]);
        updateCoords(lat, lng);
      });

      if (!initialData) {
        detectGPS();
      }
    }

    return () => {
      if (mapRef.current) {
        mapRef.current.remove();
        mapRef.current = null;
      }
    };
  }, []);

  const detectGPS = () => {
    if (navigator.geolocation) {
      navigator.geolocation.getCurrentPosition(
        (position) => {
          const { latitude, longitude } = position.coords;
          updateCoords(latitude, longitude);
          if (mapRef.current) {
            mapRef.current.setView([latitude, longitude], 17);
            markerRef.current.setLatLng([latitude, longitude]);
          }
        },
        (error) => { console.error("GPS Error: ", error); },
        { enableHighAccuracy: true, timeout: 5000, maximumAge: 0 }
      );
    }
  };

  const updateCoords = async (lat: number, lng: number) => {
    setFormData(prev => ({ ...prev, latitude: lat, longitude: lng }));
    try {
      const response = await fetch(`https://nominatim.openstreetmap.org/reverse?format=json&lat=${lat}&lon=${lng}&zoom=18&addressdetails=1`);
      const data = await response.json();
      if (data && data.address) {
        const addr = data.display_name || '';
        const city = data.address.city || data.address.town || data.address.city_district || '';
        const province = data.address.state || '';
        const zip = data.address.postcode || '';
        
        setFormData(prev => ({
          ...prev,
          address: addr,
          city: city,
          province: province,
          zip_code: zip
        }));
      }
    } catch (err) { console.error("Reverse Geotag Error:", err); }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => {
    const { name, value } = e.target;
    setFormData(prev => ({ ...prev, [name]: value }));
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    
    // Validasi Kolom Wajib
    if (
      !formData.name.trim() || 
      !formData.location_type || 
      !formData.address.trim() || 
      !formData.city.trim() || 
      !formData.province.trim() || 
      !formData.radius ||
      !formData.latitude ||
      !formData.longitude
    ) {
      Swal.fire({
        title: 'Data Belum Lengkap',
        text: 'Mohon lengkapi Jenis, Nama, Map, Alamat, Kota, Provinsi, dan Radius sebelum menyimpan.',
        icon: 'warning',
        confirmButtonColor: '#006E62'
      });
      return;
    }

    onSubmit(formData);
  };

  const Label = ({ children, required = false }: { children: React.ReactNode, required?: boolean }) => (
    <label className="text-[10px] font-bold text-gray-500 uppercase tracking-wider flex items-center gap-1">
      {children}
      {required && <span className="text-red-500">*</span>}
    </label>
  );

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-40 p-4">
      <div className="bg-white rounded-md shadow-2xl w-full max-w-2xl overflow-hidden flex flex-col max-h-[90vh]">
        <div className="px-6 py-3 border-b border-gray-100 flex items-center justify-between bg-white sticky top-0">
          <h3 className="text-base font-bold text-[#006E62]">
            {initialData ? 'Ubah Lokasi' : 'Tambah Lokasi Baru'}
          </h3>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={18} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="flex-1 overflow-y-auto p-5 space-y-4">
          <div className="bg-orange-50/50 border border-orange-100 p-2 rounded mb-4 flex items-center gap-2">
            <AlertCircle size={14} className="text-orange-400 shrink-0" />
            <p className="text-[10px] text-orange-600 font-medium">Kolom bertanda <span className="text-red-500 font-bold">*</span> wajib diisi agar data lokasi dapat disimpan.</p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
            <div className="space-y-1">
              <Label required>Jenis Lokasi</Label>
              <select
                name="location_type"
                value={formData.location_type}
                onChange={handleChange}
                className="w-full px-2 py-1.5 text-[13px] border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] focus:border-[#006E62] outline-none bg-gray-50"
              >
                <option value="Kantor">Kantor</option>
                <option value="Gudang">Gudang</option>
                <option value="Proyek">Proyek</option>
                <option value="Toko">Toko</option>
                <option value="Lainnya">Lainnya</option>
              </select>
            </div>
            <div className="md:col-span-2 space-y-1">
              <Label required>Nama Lokasi</Label>
              <input
                required
                name="name"
                value={formData.name}
                onChange={handleChange}
                placeholder="cth: Kantor Pusat Jakarta"
                className="w-full px-2 py-1.5 text-[13px] border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] focus:border-[#006E62] outline-none bg-white"
              />
            </div>
          </div>

          <div className="space-y-1">
            <div className="flex items-center justify-between mb-1">
              <Label required>Peta Geotag (Map)</Label>
              <button 
                type="button"
                onClick={detectGPS}
                className="flex items-center gap-1 text-[9px] font-bold text-[#006E62] hover:underline"
              >
                <Navigation size={10} /> DETEKSI POSISI
              </button>
            </div>
            <div id="map-container" className="mb-1 border border-gray-200 rounded shadow-sm"></div>
            <div className="flex gap-4">
              <div className="text-[9px] text-gray-400 font-mono">LAT: {formData.latitude.toFixed(6)}</div>
              <div className="text-[9px] text-gray-400 font-mono">LNG: {formData.longitude.toFixed(6)}</div>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
            <div className="md:col-span-3 space-y-1">
              <Label required>Alamat Lengkap</Label>
              <input
                required
                name="address"
                value={formData.address}
                onChange={handleChange}
                className="w-full px-2 py-1.5 text-[13px] border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] focus:border-[#006E62] outline-none"
              />
            </div>
            <div className="space-y-1">
              <Label required>Kota</Label>
              <input
                required
                name="city"
                value={formData.city}
                onChange={handleChange}
                className="w-full px-2 py-1.5 text-[13px] border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] focus:border-[#006E62] outline-none"
              />
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
            <div className="space-y-1">
              <Label required>Provinsi</Label>
              <input
                name="province"
                value={formData.province}
                onChange={handleChange}
                className="w-full px-2 py-1.5 text-[13px] border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] focus:border-[#006E62] outline-none"
              />
            </div>
            <div className="space-y-1">
              <label className="text-[10px] font-bold text-gray-500 uppercase tracking-wider">Kode Pos</label>
              <input
                name="zip_code"
                value={formData.zip_code}
                onChange={handleChange}
                className="w-full px-2 py-1.5 text-[13px] border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] focus:border-[#006E62] outline-none"
              />
            </div>
            <div className="space-y-1">
              <label className="text-[10px] font-bold text-gray-500 uppercase tracking-wider">Telepon</label>
              <input
                name="phone"
                value={formData.phone}
                onChange={handleChange}
                className="w-full px-2 py-1.5 text-[13px] border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] focus:border-[#006E62] outline-none"
              />
            </div>
            <div className="space-y-1">
              <Label required>Radius (m)</Label>
              <input
                type="number"
                name="radius"
                value={formData.radius}
                onChange={handleChange}
                className="w-full px-2 py-1.5 text-[13px] border border-gray-200 rounded focus:ring-1 focus:ring-[#006E62] focus:border-[#006E62] outline-none"
              />
            </div>
          </div>
        </form>

        <div className="p-4 border-t border-gray-100 flex justify-end gap-3 bg-gray-50">
          <button type="button" onClick={onClose} className="px-4 py-2 text-xs font-bold text-gray-500 hover:text-gray-700 transition-colors uppercase">
            Batal
          </button>
          <button
            onClick={handleSubmit}
            className="flex items-center gap-2 bg-[#006E62] text-white px-6 py-2 rounded-md hover:bg-[#005a50] transition-colors shadow-md text-xs font-bold uppercase"
          >
            <Save size={14} /> Simpan Data
          </button>
        </div>
      </div>
    </div>
  );
};

export default LocationForm;
