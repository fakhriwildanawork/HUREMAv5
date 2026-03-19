
import React, { useEffect, useRef } from 'react';

declare var L: any;

interface PresenceMapProps {
  userLat: number;
  userLng: number;
  officeLat: number;
  officeLng: number;
  radius: number;
}

const PresenceMap: React.FC<PresenceMapProps> = ({ userLat, userLng, officeLat, officeLng, radius }) => {
  const mapRef = useRef<any>(null);
  const containerId = "map-presence-container";

  useEffect(() => {
    if (!mapRef.current) {
      mapRef.current = L.map(containerId, { zoomControl: false }).setView([officeLat, officeLng], 16);
      L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png').addTo(mapRef.current);
      
      // Office Circle
      L.circle([officeLat, officeLng], {
        color: '#006E62',
        fillColor: '#006E62',
        fillOpacity: 0.1,
        radius: radius
      }).addTo(mapRef.current);

      // Office Marker (Pusat Lokasi)
      L.marker([officeLat, officeLng], {
        icon: L.divIcon({ 
          className: 'bg-[#006E62] w-3 h-3 rounded-full border-2 border-white' 
        })
      }).addTo(mapRef.current);

      // User ACTUAL Location Marker (HERE()) - Red Static Pin
      L.marker([userLat, userLng], {
        icon: L.divIcon({ 
          className: 'bg-red-600 w-3 h-3 rounded-full border-2 border-white shadow-lg' 
        })
      }).addTo(mapRef.current);

      const bounds = L.latLngBounds([[userLat, userLng], [officeLat, officeLng]]);
      mapRef.current.fitBounds(bounds, { padding: [20, 20] });
    }

    return () => {
      if (mapRef.current) {
        mapRef.current.remove();
        mapRef.current = null;
      }
    };
  }, [userLat, userLng]);

  return (
    <div id={containerId} className="w-full h-32 rounded-lg border border-gray-100 shadow-inner overflow-hidden" />
  );
};

export default PresenceMap;
