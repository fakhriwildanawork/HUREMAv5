import React from 'react';
import { X, MapPin, Calendar, User, Info, Camera, Link as LinkIcon, ExternalLink, Navigation, FileText } from 'lucide-react';
import { MapContainer, TileLayer, Marker, Popup } from 'react-leaflet';
import L from 'leaflet';
import 'leaflet/dist/leaflet.css';
import { SalesReport } from '../../../types';

// Fix Leaflet marker icon issue
const DefaultIcon = L.icon({
  iconUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon.png',
  shadowUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-shadow.png',
  iconSize: [25, 41],
  iconAnchor: [12, 41]
});

L.Marker.prototype.options.icon = DefaultIcon;

interface SalesReportDetailProps {
  report: SalesReport;
  onClose: () => void;
}

const SalesReportDetail: React.FC<SalesReportDetailProps> = ({ report, onClose }) => {
  const position: [number, number] = [report.latitude, report.longitude];

  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 p-4 backdrop-blur-sm">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-3xl overflow-hidden animate-in zoom-in duration-200 flex flex-col max-h-[90vh]">
        {/* Header */}
        <div className="px-6 py-4 border-b border-gray-50 flex items-center justify-between bg-gray-50/50 shrink-0">
          <div className="flex items-center gap-2">
            <MapPin size={18} className="text-[#006E62]" />
            <h3 className="text-sm font-bold text-gray-800 uppercase tracking-tight">Detail Kunjungan Sales</h3>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={20} />
          </button>
        </div>

        {/* Content */}
        <div className="p-6 overflow-y-auto scrollbar-thin space-y-6">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            {/* Left Column: Info */}
            <div className="space-y-5">
              <div className="p-4 bg-emerald-50 rounded-2xl border border-emerald-100">
                <p className="text-[10px] font-bold text-emerald-600 uppercase tracking-widest mb-1">Nama Klien / Toko</p>
                <h4 className="text-base font-bold text-gray-800">{report.customer_name}</h4>
                <div className="flex items-center gap-2 mt-2">
                  <span className="px-2 py-0.5 bg-white border border-emerald-100 rounded text-[10px] font-bold text-[#006E62] uppercase">
                    {report.activity_type}
                  </span>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Sales Person</p>
                  <div className="flex items-center gap-2">
                    <div className="w-8 h-8 rounded-full bg-gray-100 flex items-center justify-center text-gray-500">
                      <User size={14} />
                    </div>
                    <div>
                      <p className="text-xs font-bold text-gray-800">{report.account?.full_name || 'N/A'}</p>
                      <p className="text-[9px] text-gray-400 font-medium">{report.account?.internal_nik || 'N/A'}</p>
                    </div>
                  </div>
                </div>
                <div>
                  <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Waktu Laporan</p>
                  <div className="flex items-center gap-2">
                    <div className="w-8 h-8 rounded-full bg-gray-100 flex items-center justify-center text-gray-500">
                      <Calendar size={14} />
                    </div>
                    <div>
                      <p className="text-xs font-bold text-gray-800">
                        {new Date(report.reported_at).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' })}
                      </p>
                      <p className="text-[9px] text-gray-400 font-medium">
                        {new Date(report.reported_at).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })} WIB
                      </p>
                    </div>
                  </div>
                </div>
              </div>

              <div>
                <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-2">Catatan Kunjungan</p>
                <div className="p-4 bg-gray-50 rounded-2xl border border-gray-100 min-h-[100px]">
                  <p className="text-xs text-gray-600 leading-relaxed whitespace-pre-wrap italic">"{report.description}"</p>
                </div>
              </div>

              {report.photo_urls && report.photo_urls.length > 0 && (
                <div>
                  <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-2">Bukti Foto ({report.photo_urls.length})</p>
                  <div className="flex flex-wrap gap-2">
                    {report.photo_urls.map((url, idx) => (
                      <a 
                        key={idx} 
                        href={url} 
                        target="_blank" 
                        rel="noopener noreferrer"
                        className="flex items-center gap-1.5 px-3 py-2 bg-white border border-gray-100 rounded-xl hover:bg-emerald-50 hover:border-emerald-200 transition-all group"
                      >
                        <Camera size={14} className="text-gray-400 group-hover:text-[#006E62]" />
                        <span className="text-[10px] font-bold text-gray-500 group-hover:text-[#006E62]">Foto {idx + 1}</span>
                        <ExternalLink size={10} className="text-gray-300 group-hover:text-[#006E62]" />
                      </a>
                    ))}
                  </div>
                </div>
              )}
            </div>

            {/* Right Column: Map */}
            <div className="space-y-4">
              <div className="flex items-center justify-between">
                <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest">Lokasi Geotagging</p>
                <a 
                  href={`https://www.google.com/maps?q=${report.latitude},${report.longitude}`}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="flex items-center gap-1.5 text-[10px] font-bold text-[#006E62] hover:underline"
                >
                  <Navigation size={12} />
                  Buka di Google Maps
                </a>
              </div>
              
              <div className="h-[300px] rounded-2xl overflow-hidden border border-gray-100 shadow-inner z-0">
                <MapContainer center={position} zoom={15} scrollWheelZoom={false} style={{ height: '100%', width: '100%' }}>
                  <TileLayer
                    attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
                    url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
                  />
                  <Marker position={position}>
                    <Popup>
                      <div className="text-xs font-bold">
                        {report.customer_name}<br/>
                        <span className="text-[10px] font-normal text-gray-500">{report.latitude}, {report.longitude}</span>
                      </div>
                    </Popup>
                  </Marker>
                </MapContainer>
              </div>

              <div className="p-3 bg-blue-50 rounded-xl space-y-3">
                <div className="flex items-start gap-3">
                  <Info size={16} className="text-blue-500 shrink-0 mt-0.5" />
                  <p className="text-[9px] text-blue-700 leading-relaxed font-medium">
                    Titik koordinat diambil secara otomatis menggunakan GPS perangkat sales saat laporan dikirimkan.
                  </p>
                </div>
                {report.address && (
                  <div className="pt-2 border-t border-blue-100 flex items-start gap-2">
                    <MapPin size={14} className="text-blue-500 shrink-0 mt-0.5" />
                    <p className="text-[10px] text-blue-800 font-bold leading-tight">{report.address}</p>
                  </div>
                )}
              </div>

              {/* Attached Files */}
              {report.file_ids && report.file_ids.length > 0 && (
                <div className="space-y-2">
                  <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest px-1">Dokumen Terlampir ({report.file_ids.length})</p>
                  <div className="space-y-2">
                    {report.file_ids.map((id, idx) => (
                      <a 
                        key={idx}
                        href={`https://drive.google.com/file/d/${id}/view`}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="flex items-center justify-between p-3 bg-gray-50 border border-gray-100 rounded-xl hover:bg-emerald-50 hover:border-emerald-200 transition-all group"
                      >
                        <div className="flex items-center gap-3">
                          <div className="w-8 h-8 bg-white rounded-lg flex items-center justify-center text-[#006E62] shadow-sm">
                            <FileText size={16} />
                          </div>
                          <div>
                            <p className="text-[10px] font-bold text-gray-700">Lampiran Dokumen {idx + 1}</p>
                            <p className="text-[8px] text-gray-400 font-medium uppercase tracking-tighter">Google Drive File</p>
                          </div>
                        </div>
                        <ExternalLink size={14} className="text-gray-300 group-hover:text-[#006E62]" />
                      </a>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className="px-6 py-4 border-t border-gray-50 bg-gray-50/30 flex justify-end shrink-0">
          <button 
            onClick={onClose}
            className="px-6 py-2.5 bg-white border border-gray-100 text-gray-400 rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-gray-50 transition-all shadow-sm"
          >
            Tutup Detail
          </button>
        </div>
      </div>
    </div>
  );
};

export default SalesReportDetail;
