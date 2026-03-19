import React from 'react';
import { X, MapPin, Navigation, Clock, User, Calendar } from 'lucide-react';
import { MapContainer, TileLayer, Marker, Popup, Polyline } from 'react-leaflet';
import L from 'leaflet';
import 'leaflet/dist/leaflet.css';
import { SalesReport } from '../../../types';

// Fix Leaflet icon issue
// @ts-ignore
import icon from 'leaflet/dist/images/marker-icon.png';
// @ts-ignore
import iconShadow from 'leaflet/dist/images/marker-shadow.png';

let DefaultIcon = L.icon({
  iconUrl: icon,
  shadowUrl: iconShadow,
  iconSize: [25, 41],
  iconAnchor: [12, 41]
});

L.Marker.prototype.options.icon = DefaultIcon;

interface SalesRouteMapProps {
  reports: SalesReport[];
  onClose: () => void;
  onViewDetail: (report: SalesReport) => void;
}

const SalesRouteMap: React.FC<SalesRouteMapProps> = ({ reports, onClose, onViewDetail }) => {
  // Sort reports by time to draw the polyline correctly
  const sortedReports = [...reports].sort((a, b) => 
    new Date(a.reported_at).getTime() - new Date(b.reported_at).getTime()
  );

  const positions = sortedReports.map(r => [r.latitude, r.longitude] as [number, number]);
  
  // Calculate center of the map
  const center: [number, number] = positions.length > 0 
    ? [
        positions.reduce((sum, p) => sum + p[0], 0) / positions.length,
        positions.reduce((sum, p) => sum + p[1], 0) / positions.length
      ]
    : [-6.200000, 106.816666]; // Default to Jakarta

  const employeeName = reports[0]?.account?.full_name || 'Sales';
  const reportDate = new Date(reports[0]?.reported_at).toLocaleDateString('id-ID', { 
    day: 'numeric', 
    month: 'long', 
    year: 'numeric' 
  });

  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
      <div className="bg-white w-full max-w-5xl h-[90vh] rounded-3xl overflow-hidden shadow-2xl flex flex-col animate-in fade-in zoom-in duration-300">
        {/* Header */}
        <div className="p-6 border-b border-gray-100 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-4">
            <div className="w-12 h-12 rounded-2xl bg-emerald-50 flex items-center justify-center text-[#006E62]">
              <Navigation size={24} />
            </div>
            <div>
              <h3 className="text-lg font-bold text-gray-800 tracking-tight">Rute Kunjungan Sales</h3>
              <div className="flex items-center gap-3 mt-0.5">
                <div className="flex items-center gap-1 text-[10px] font-bold text-[#006E62] uppercase tracking-wider">
                  <User size={10} />
                  <span>{employeeName}</span>
                </div>
                <div className="flex items-center gap-1 text-[10px] font-bold text-gray-400 uppercase tracking-wider">
                  <Calendar size={10} />
                  <span>{reportDate}</span>
                </div>
              </div>
            </div>
          </div>
          <button 
            onClick={onClose}
            className="p-2 hover:bg-gray-100 rounded-xl transition-colors text-gray-400 hover:text-gray-600"
          >
            <X size={24} />
          </button>
        </div>

        {/* Map Area */}
        <div className="flex-1 relative">
          <MapContainer 
            center={center} 
            zoom={13} 
            scrollWheelZoom={true} 
            style={{ height: '100%', width: '100%' }}
          >
            <TileLayer
              attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
              url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
            />
            
            {/* Draw the route line */}
            <Polyline 
              positions={positions} 
              color="#006E62" 
              weight={3} 
              opacity={0.6} 
              dashArray="10, 10"
            />

            {/* Markers for each visit */}
            {sortedReports.map((report, index) => (
              <Marker 
                key={report.id} 
                position={[report.latitude, report.longitude]}
              >
                <Popup>
                  <div className="p-1 min-w-[150px]">
                    <div className="flex items-center gap-2 mb-2">
                      <div className="w-6 h-6 rounded-full bg-[#006E62] text-white flex items-center justify-center text-[10px] font-bold">
                        {index + 1}
                      </div>
                      <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">
                        {new Date(report.reported_at).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })} WIB
                      </span>
                    </div>
                    <h4 className="text-xs font-bold text-gray-800 mb-1">{report.customer_name}</h4>
                    <p className="text-[10px] text-gray-500 line-clamp-2 mb-2 italic">"{report.description}"</p>
                    <button 
                      onClick={() => onViewDetail(report)}
                      className="w-full py-1.5 bg-gray-50 hover:bg-emerald-50 text-[#006E62] rounded text-[9px] font-bold uppercase tracking-wider transition-colors border border-gray-100"
                    >
                      Lihat Detail
                    </button>
                  </div>
                </Popup>
              </Marker>
            ))}
          </MapContainer>

          {/* Legend / Stats overlay */}
          <div className="absolute bottom-6 left-6 z-[1000] bg-white/90 backdrop-blur-md p-4 rounded-2xl border border-gray-100 shadow-xl max-w-xs">
            <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-3">Ringkasan Rute</h4>
            <div className="space-y-3">
              <div className="flex items-center gap-3">
                <div className="w-8 h-8 rounded-lg bg-emerald-50 flex items-center justify-center text-[#006E62]">
                  <MapPin size={16} />
                </div>
                <div>
                  <p className="text-[11px] font-bold text-gray-800">{reports.length} Lokasi</p>
                  <p className="text-[9px] text-gray-400 font-medium">Total kunjungan hari ini</p>
                </div>
              </div>
              <div className="flex items-center gap-3">
                <div className="w-8 h-8 rounded-lg bg-blue-50 flex items-center justify-center text-blue-600">
                  <Clock size={16} />
                </div>
                <div>
                  <p className="text-[11px] font-bold text-gray-800">
                    {new Date(sortedReports[0].reported_at).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })} - 
                    {new Date(sortedReports[sortedReports.length - 1].reported_at).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })}
                  </p>
                  <p className="text-[9px] text-gray-400 font-medium">Rentang waktu kunjungan</p>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default SalesRouteMap;
