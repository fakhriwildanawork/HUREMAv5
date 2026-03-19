import React, { useState, useEffect } from 'react';
import { ArrowLeft, Save, Filter, Search, UserCheck, MapPin, Briefcase, GraduationCap, CheckCircle2 } from 'lucide-react';
import { financeService } from '../../services/financeService';
import { accountService } from '../../services/accountService';
import { locationService } from '../../services/locationService';
import { SalaryScheme, Account, Location } from '../../types';
import Swal from 'sweetalert2';

interface SalarySchemeAssignmentProps {
  onBack: () => void;
}

const SalarySchemeAssignment: React.FC<SalarySchemeAssignmentProps> = ({ onBack }) => {
  const [loading, setLoading] = useState(false);
  const [schemes, setSchemes] = useState<SalaryScheme[]>([]);
  const [accounts, setAccounts] = useState<Account[]>([]);
  const [assignments, setAssignments] = useState<any[]>([]);
  const [locations, setLocations] = useState<Location[]>([]);
  const [distinctAttrs, setDistinctAttrs] = useState<{ positions: string[], grades: string[] }>({ positions: [], grades: [] });
  
  const [selectedSchemeId, setSelectedSchemeId] = useState('');
  const [selectedAccountIds, setSelectedAccountIds] = useState<string[]>([]);
  
  // Filters
  const [filterLocation, setFilterLocation] = useState('');
  const [filterGrade, setFilterGrade] = useState('');
  const [filterPosition, setFilterPosition] = useState('');
  const [searchQuery, setSearchQuery] = useState('');

  useEffect(() => {
    const fetchData = async () => {
      try {
        const [sData, aData, lData, dData, asData] = await Promise.all([
          financeService.getSchemes(),
          accountService.getAll(),
          locationService.getAll(),
          accountService.getDistinctAttributes(),
          financeService.getAssignments()
        ]);
        setSchemes(sData);
        setAccounts(aData);
        setLocations(lData);
        setDistinctAttrs(dData);
        setAssignments(asData);
      } catch (error) {
        console.error('Error fetching data:', error);
      }
    };
    fetchData();
  }, []);

  const filteredAccounts = accounts.filter(acc => {
    // Filter out accounts that already have an assignment
    const hasAssignment = assignments.some(as => as.account_id === acc.id);
    if (hasAssignment) return false;

    const matchesLocation = !filterLocation || acc.location_id === filterLocation;
    const matchesGrade = !filterGrade || acc.grade === filterGrade;
    const matchesPosition = !filterPosition || acc.position === filterPosition;
    const matchesSearch = !searchQuery || 
      acc.full_name.toLowerCase().includes(searchQuery.toLowerCase()) ||
      acc.internal_nik.toLowerCase().includes(searchQuery.toLowerCase());
    
    return matchesLocation && matchesGrade && matchesPosition && matchesSearch;
  });

  const handleSelectAll = () => {
    if (selectedAccountIds.length === filteredAccounts.length) {
      setSelectedAccountIds([]);
    } else {
      setSelectedAccountIds(filteredAccounts.map(a => a.id));
    }
  };

  const handleToggleAccount = (id: string) => {
    setSelectedAccountIds(prev => 
      prev.includes(id) ? prev.filter(a => a !== id) : [...prev, id]
    );
  };

  const handleAssign = async () => {
    if (!selectedSchemeId) {
      Swal.fire('Peringatan', 'Silakan pilih skema gaji terlebih dahulu.', 'warning');
      return;
    }
    if (selectedAccountIds.length === 0) {
      Swal.fire('Peringatan', 'Silakan pilih setidaknya satu karyawan.', 'warning');
      return;
    }

    const result = await Swal.fire({
      title: 'Terapkan Skema?',
      text: `Anda akan menerapkan skema ini ke ${selectedAccountIds.length} karyawan. Skema lama (jika ada) akan digantikan.`,
      icon: 'question',
      showCancelButton: true,
      confirmButtonColor: '#006E62',
      confirmButtonText: 'Ya, Terapkan'
    });

    if (result.isConfirmed) {
      setLoading(true);
      try {
        await financeService.assignScheme(selectedSchemeId, selectedAccountIds);
        Swal.fire('Berhasil!', 'Skema gaji berhasil diterapkan ke karyawan terpilih.', 'success');
        onBack();
      } catch (error) {
        console.error('Error assigning scheme:', error);
        Swal.fire('Gagal!', 'Terjadi kesalahan saat menerapkan skema.', 'error');
      } finally {
        setLoading(false);
      }
    }
  };

  return (
    <div className="max-w-6xl mx-auto space-y-6">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-4">
          <button
            onClick={onBack}
            className="p-2 hover:bg-gray-100 rounded-full transition-colors text-gray-500"
          >
            <ArrowLeft size={20} />
          </button>
          <div>
            <h2 className="text-2xl font-bold text-gray-800">Penugasan Skema Gaji</h2>
            <p className="text-sm text-gray-500">Hubungkan karyawan dengan skema gaji yang sesuai.</p>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left Panel: Scheme Selection & Filters */}
        <div className="lg:col-span-1 space-y-6">
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 space-y-6">
            <div className="space-y-4">
              <div className="flex items-center gap-2 text-[#006E62]">
                <UserCheck size={18} />
                <h3 className="font-bold text-sm uppercase tracking-wider">Pilih Skema</h3>
              </div>
              <select
                value={selectedSchemeId}
                onChange={(e) => setSelectedSchemeId(e.target.value)}
                className="w-full px-4 py-2.5 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all font-medium"
              >
                <option value="">-- Pilih Skema Gaji --</option>
                {schemes.map(s => (
                  <option key={s.id} value={s.id}>{s.name} ({s.type})</option>
                ))}
              </select>
              {selectedSchemeId && (
                <div className="p-4 bg-blue-50 rounded-lg border border-blue-100">
                  <p className="text-[10px] font-bold text-blue-400 uppercase tracking-widest mb-2">Ringkasan Skema</p>
                  {schemes.find(s => s.id === selectedSchemeId) && (
                    <div className="space-y-1 text-xs">
                      <div className="flex justify-between">
                        <span className="text-blue-600">Gaji Pokok:</span>
                        <span className="font-bold">Rp {schemes.find(s => s.id === selectedSchemeId)?.basic_salary.toLocaleString('id-ID')}</span>
                      </div>
                      <div className="flex justify-between">
                        <span className="text-blue-600">Total Tunjangan:</span>
                        <span className="font-bold">
                          Rp {(
                            (schemes.find(s => s.id === selectedSchemeId)?.position_allowance || 0) +
                            (schemes.find(s => s.id === selectedSchemeId)?.placement_allowance || 0) +
                            (schemes.find(s => s.id === selectedSchemeId)?.other_allowance || 0)
                          ).toLocaleString('id-ID')}
                        </span>
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>

            <div className="pt-6 border-t border-gray-50 space-y-4">
              <div className="flex items-center gap-2 text-gray-500">
                <Filter size={18} />
                <h3 className="font-bold text-sm uppercase tracking-wider">Filter Karyawan</h3>
              </div>
              
              <div className="space-y-3">
                <div className="space-y-1">
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
                    <MapPin size={12} /> Lokasi
                  </label>
                  <select
                    value={filterLocation}
                    onChange={(e) => setFilterLocation(e.target.value)}
                    className="w-full px-3 py-2 bg-gray-50 border-none rounded-lg text-xs focus:ring-2 focus:ring-[#006E62] transition-all"
                  >
                    <option value="">Semua Lokasi</option>
                    {locations.map(l => <option key={l.id} value={l.id}>{l.name}</option>)}
                  </select>
                </div>

                <div className="space-y-1">
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
                    <GraduationCap size={12} /> Golongan
                  </label>
                  <select
                    value={filterGrade}
                    onChange={(e) => setFilterGrade(e.target.value)}
                    className="w-full px-3 py-2 bg-gray-50 border-none rounded-lg text-xs focus:ring-2 focus:ring-[#006E62] transition-all"
                  >
                    <option value="">Semua Golongan</option>
                    {distinctAttrs.grades.map(g => <option key={g} value={g}>{g}</option>)}
                  </select>
                </div>

                <div className="space-y-1">
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
                    <Briefcase size={12} /> Jabatan
                  </label>
                  <select
                    value={filterPosition}
                    onChange={(e) => setFilterPosition(e.target.value)}
                    className="w-full px-3 py-2 bg-gray-50 border-none rounded-lg text-xs focus:ring-2 focus:ring-[#006E62] transition-all"
                  >
                    <option value="">Semua Jabatan</option>
                    {distinctAttrs.positions.map(p => <option key={p} value={p}>{p}</option>)}
                  </select>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Right Panel: Employee List */}
        <div className="lg:col-span-2 space-y-6">
          <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden flex flex-col h-[600px]">
            <div className="p-4 border-b border-gray-50 flex items-center justify-between gap-4">
              <div className="relative flex-1">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
                <input
                  type="text"
                  placeholder="Cari nama atau NIK..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full pl-10 pr-4 py-2 bg-gray-50 border-none rounded-lg text-sm focus:ring-2 focus:ring-[#006E62] transition-all"
                />
              </div>
              <button
                onClick={handleSelectAll}
                className="text-xs font-bold text-[#006E62] hover:underline whitespace-nowrap"
              >
                {selectedAccountIds.length === filteredAccounts.length ? 'Batal Pilih Semua' : 'Pilih Semua'}
              </button>
            </div>

            <div className="flex-1 overflow-y-auto p-2 space-y-1">
              {filteredAccounts.length === 0 ? (
                <div className="flex flex-col items-center justify-center h-full text-gray-400 gap-2">
                  <Search size={48} strokeWidth={1} />
                  <p className="text-sm">Karyawan tidak ditemukan.</p>
                </div>
              ) : (
                filteredAccounts.map(acc => (
                  <div
                    key={acc.id}
                    onClick={() => handleToggleAccount(acc.id)}
                    className={`flex items-center gap-4 p-3 rounded-xl cursor-pointer transition-all border ${
                      selectedAccountIds.includes(acc.id)
                        ? 'bg-emerald-50 border-emerald-200'
                        : 'hover:bg-gray-50 border-transparent'
                    }`}
                  >
                    <div className={`w-5 h-5 rounded border-2 flex items-center justify-center transition-all ${
                      selectedAccountIds.includes(acc.id)
                        ? 'bg-[#006E62] border-[#006E62]'
                        : 'border-gray-200'
                    }`}>
                      {selectedAccountIds.includes(acc.id) && <CheckCircle2 size={14} className="text-white" />}
                    </div>
                    <div className="flex-1 min-w-0">
                      <div className="font-bold text-sm text-gray-900 truncate">{acc.full_name}</div>
                      <div className="flex items-center gap-2 text-[10px] text-gray-500 font-medium">
                        <span>{acc.internal_nik}</span>
                        <span>•</span>
                        <span>{acc.position}</span>
                        <span>•</span>
                        <span>{acc.grade}</span>
                      </div>
                    </div>
                    <div className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">
                      {locations.find(l => l.id === acc.location_id)?.name || '-'}
                    </div>
                  </div>
                ))
              )}
            </div>

            <div className="p-4 border-t border-gray-50 bg-gray-50 flex items-center justify-between">
              <div className="text-sm text-gray-500">
                <span className="font-bold text-[#006E62]">{selectedAccountIds.length}</span> Karyawan terpilih
              </div>
              <button
                onClick={handleAssign}
                disabled={loading || !selectedSchemeId || selectedAccountIds.length === 0}
                className="flex items-center gap-2 px-6 py-2 bg-[#006E62] text-white rounded-lg hover:bg-[#005a50] transition-all text-sm font-bold shadow-md disabled:opacity-50"
              >
                {loading ? (
                  <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></div>
                ) : (
                  <Save size={18} />
                )}
                Terapkan Skema
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default SalarySchemeAssignment;
