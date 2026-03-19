
import React, { useState, useEffect } from 'react';
import { Plus, Search, MapPin, Grid, List as ListIcon, Filter, ArrowLeft } from 'lucide-react';
import Swal from 'sweetalert2';
import { locationService } from '../../services/locationService';
import { Location, LocationInput } from '../../types';
import LocationForm from './LocationForm';
import LocationDetail from './LocationDetail';
import { CardSkeleton } from '../../components/Common/Skeleton';
import LoadingSpinner from '../../components/Common/LoadingSpinner';

const LocationMain: React.FC = () => {
  const [locations, setLocations] = useState<Location[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [viewMode, setViewMode] = useState<'grid' | 'list'>('grid');
  const [searchTerm, setSearchTerm] = useState('');
  const [showForm, setShowForm] = useState(false);
  const [selectedLocationId, setSelectedLocationId] = useState<string | null>(null);
  const [editingLocation, setEditingLocation] = useState<Location | null>(null);

  useEffect(() => {
    fetchLocations();
  }, []);

  const fetchLocations = async () => {
    try {
      setIsLoading(true);
      const data = await locationService.getAll();
      setLocations(data);
    } catch (error) {
      Swal.fire('Gagal', 'Gagal memuat data lokasi', 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleCreate = async (input: LocationInput) => {
    setIsSaving(true);
    // Optimistic UI update
    const tempId = `temp-${Math.random().toString(36).substring(7)}`;
    const optimisticLocation: Location = { 
      ...input, 
      id: tempId, 
      created_at: new Date().toISOString(),
      search_all: `${input.name} ${input.location_type} ${input.address} ${input.city}`.toLowerCase()
    };
    
    setLocations(prev => [optimisticLocation, ...prev]);
    setShowForm(false);

    try {
      const created = await locationService.create(input);
      setLocations(prev => prev.map(loc => loc.id === tempId ? created : loc));
      Swal.fire({
        title: 'Berhasil!',
        text: 'Lokasi baru telah ditambahkan.',
        icon: 'success',
        timer: 1500,
        showConfirmButton: false
      });
    } catch (error) {
      setLocations(prev => prev.filter(loc => loc.id !== tempId));
      Swal.fire('Gagal', 'Terjadi kesalahan saat menyimpan data', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleUpdate = async (id: string, input: Partial<LocationInput>) => {
    setIsSaving(true);
    try {
      const updated = await locationService.update(id, input);
      setLocations(prev => prev.map(loc => loc.id === id ? updated : loc));
      setEditingLocation(null);
      setShowForm(false);
      Swal.fire({
        title: 'Terupdate!',
        text: 'Data lokasi berhasil diperbarui.',
        icon: 'success',
        timer: 1500,
        showConfirmButton: false
      });
    } catch (error) {
      Swal.fire('Gagal', 'Gagal memperbarui data', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleDelete = async (id: string) => {
    const result = await Swal.fire({
      title: 'Apakah Anda yakin?',
      text: "Data ini akan dihapus permanen dari sistem.",
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
        await locationService.delete(id);
        setLocations(prev => prev.filter(loc => loc.id !== id));
        setSelectedLocationId(null);
        Swal.fire('Terhapus!', 'Data lokasi telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Gagal menghapus data', 'error');
      } finally {
        setIsSaving(false);
      }
    }
  };

  const filteredLocations = locations.filter(loc => {
    const searchStr = (loc.search_all || `${loc.name} ${loc.address} ${loc.city} ${loc.location_type}`).toLowerCase();
    return searchStr.includes(searchTerm.toLowerCase());
  });

  // Render Halaman Detail jika ada lokasi terpilih
  if (selectedLocationId) {
    return (
      <div className="animate-in fade-in slide-in-from-right duration-300">
        <button 
          onClick={() => setSelectedLocationId(null)}
          className="mb-4 flex items-center gap-2 text-gray-500 hover:text-[#006E62] transition-colors font-bold text-xs uppercase"
        >
          <ArrowLeft size={16} /> Kembali ke Daftar
        </button>
        <LocationDetail
          id={selectedLocationId}
          onClose={() => setSelectedLocationId(null)}
          onEdit={(loc) => { setSelectedLocationId(null); setEditingLocation(loc); setShowForm(true); }}
          onDelete={(id) => handleDelete(id)}
        />
      </div>
    );
  }

  return (
    <div className="space-y-6">
      {isSaving && <LoadingSpinner />}

      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="relative flex-1 max-w-md">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
          <input
            type="text"
            placeholder="Cari (Nama, Alamat, Jenis, Kota)..."
            className="w-full pl-10 pr-4 py-2 bg-gray-50 border border-gray-200 rounded-md focus:outline-none focus:ring-2 focus:ring-[#006E62] focus:border-transparent transition-all text-sm"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </div>
        
        <div className="flex items-center gap-2">
          <div className="flex border border-gray-200 rounded-md overflow-hidden bg-white">
            <button 
              onClick={() => setViewMode('grid')}
              className={`p-2 transition-colors ${viewMode === 'grid' ? 'bg-gray-100 text-[#006E62]' : 'text-gray-400 hover:text-gray-600'}`}
            >
              <Grid size={18} />
            </button>
            <button 
              onClick={() => setViewMode('list')}
              className={`p-2 transition-colors ${viewMode === 'list' ? 'bg-gray-100 text-[#006E62]' : 'text-gray-400 hover:text-gray-600'}`}
            >
              <ListIcon size={18} />
            </button>
          </div>
          <button 
            onClick={() => { setEditingLocation(null); setShowForm(true); }}
            className="flex items-center gap-2 bg-[#006E62] text-white px-4 py-2 rounded-md hover:bg-[#005a50] transition-colors shadow-sm"
          >
            <Plus size={18} />
            <span className="font-medium text-sm">Tambah</span>
          </button>
        </div>
      </div>

      {isLoading ? (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
          {[1, 2, 3, 4, 5, 6].map(i => <CardSkeleton key={i} />)}
        </div>
      ) : filteredLocations.length === 0 ? (
        <div className="flex flex-col items-center justify-center py-20 text-gray-400">
          <MapPin size={48} strokeWidth={1} className="mb-4" />
          <p className="text-lg">Data tidak ditemukan.</p>
        </div>
      ) : viewMode === 'grid' ? (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
          {filteredLocations.map(location => (
            <div 
              key={location.id} 
              onClick={() => setSelectedLocationId(location.id)}
              className="group bg-white border border-gray-100 p-4 rounded-md shadow-sm hover:shadow-md transition-all cursor-pointer border-l-4 border-l-transparent hover:border-l-[#006E62]"
            >
              <div className="flex justify-between items-start mb-2">
                <h3 className="font-bold text-[#006E62] group-hover:text-[#005a50] line-clamp-1 text-sm">{location.name}</h3>
                <span className="text-[9px] font-bold text-gray-400 border border-gray-100 px-1.5 py-0.5 rounded uppercase">{location.location_type || 'Lokasi'}</span>
              </div>
              <p className="text-[11px] text-gray-500 mb-3 line-clamp-1">{location.address}</p>
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2 text-[10px] text-gray-400 font-bold uppercase tracking-wider">
                  <MapPin size={10} /> {location.city}
                </div>
                <div className="text-[10px] text-gray-300">R: {location.radius}m</div>
              </div>
            </div>
          ))}
        </div>
      ) : (
        <div className="bg-white border border-gray-100 rounded-md overflow-hidden shadow-sm overflow-x-auto">
          <table className="w-full text-left min-w-[600px]">
            <thead className="bg-gray-50 text-[10px] font-bold text-gray-500 uppercase">
              <tr>
                <th className="px-6 py-3">Nama & Jenis</th>
                <th className="px-6 py-3">Alamat & Kota</th>
                <th className="px-6 py-3">Telepon</th>
                <th className="px-6 py-3 text-right">Radius</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filteredLocations.map(location => (
                <tr 
                  key={location.id} 
                  onClick={() => setSelectedLocationId(location.id)}
                  className="hover:bg-gray-50 cursor-pointer transition-colors"
                >
                  <td className="px-6 py-4">
                    <div className="font-bold text-[#006E62] text-xs">{location.name}</div>
                    <div className="text-[9px] text-gray-400 uppercase font-bold">{location.location_type}</div>
                  </td>
                  <td className="px-6 py-4">
                    <div className="text-[11px] text-gray-600 truncate max-w-xs">{location.address}</div>
                    <div className="text-[10px] text-gray-400 font-bold uppercase">{location.city}</div>
                  </td>
                  <td className="px-6 py-4 text-xs text-gray-500">{location.phone || '-'}</td>
                  <td className="px-6 py-4 text-xs text-right font-bold text-gray-400">{location.radius}m</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* Modals */}
      {showForm && (
        <LocationForm 
          onClose={() => { setShowForm(false); setEditingLocation(null); }}
          onSubmit={editingLocation ? (data) => handleUpdate(editingLocation.id, data) : handleCreate}
          initialData={editingLocation || undefined}
        />
      )}
    </div>
  );
};

export default LocationMain;
