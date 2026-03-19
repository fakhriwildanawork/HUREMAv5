
import React from 'react';
import { X, FileBadge, FileText, Paperclip, ExternalLink } from 'lucide-react';
import { AccountContract } from '../../types';
import { googleDriveService } from '../../services/googleDriveService';

interface ContractDetailModalProps {
  contract: AccountContract;
  onClose: () => void;
}

const ContractDetailModal: React.FC<ContractDetailModalProps> = ({ contract, onClose }) => {
  const formatDate = (dateStr?: string | null) => {
    if (!dateStr) return '-';
    return new Date(dateStr).toLocaleDateString('id-ID', {
      day: 'numeric',
      month: 'long',
      year: 'numeric'
    });
  };

  const isImage = (fileId: string) => {
    if (!fileId.includes('|')) return true; // Default to true for old data
    const [, name] = fileId.split('|');
    return /\.(jpg|jpeg|png|webp|gif|svg|bmp)$/i.test(name);
  };

  const isImg = contract.file_id ? isImage(contract.file_id) : false;
  const fileUrl = contract.file_id ? googleDriveService.getFileUrl(contract.file_id, isImg) : null;
  const viewerUrl = contract.file_id ? googleDriveService.getViewerUrl(contract.file_id) : null;

  return (
    <div className="fixed inset-0 z-[250] bg-black/60 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in duration-200">
      <div className="bg-white rounded-md shadow-2xl w-full max-w-2xl overflow-hidden flex flex-col max-h-[90vh] animate-in zoom-in duration-200">
        <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between bg-white shrink-0">
          <div>
            <h3 className="text-base font-bold text-[#006E62]">Detail Kontrak Kerja</h3>
            <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Informasi Lengkap Perjanjian</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors p-1">
            <X size={20} />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-6 space-y-6 bg-gray-50/30">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="space-y-4">
              <div className="bg-white p-4 rounded border border-gray-100 shadow-sm">
                <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Nomor Kontrak</p>
                <p className="text-sm font-bold text-[#006E62]">{contract.contract_number}</p>
              </div>
              <div className="bg-white p-4 rounded border border-gray-100 shadow-sm">
                <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Jenis Kontrak</p>
                <p className="text-sm font-bold text-gray-700">{contract.contract_type}</p>
              </div>
              <div className="bg-white p-4 rounded border border-gray-100 shadow-sm">
                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Tanggal Mulai</p>
                    <p className="text-xs font-bold text-gray-700">{formatDate(contract.start_date)}</p>
                  </div>
                  <div>
                    <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-1">Tanggal Berakhir</p>
                    <p className="text-xs font-bold text-gray-700">{contract.end_date ? formatDate(contract.end_date) : 'TETAP'}</p>
                  </div>
                </div>
              </div>
            </div>

            <div className="bg-white p-4 rounded border border-gray-100 shadow-sm flex flex-col">
              <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest mb-2">Keterangan / Catatan</p>
              <p className="text-xs text-gray-600 leading-relaxed italic flex-1">
                {contract.notes || 'Tidak ada catatan tambahan.'}
              </p>
            </div>
          </div>

          <div className="bg-white p-4 rounded border border-gray-100 shadow-sm">
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-2">
                <Paperclip size={16} className="text-[#006E62]" />
                <h4 className="text-[10px] font-bold uppercase tracking-widest text-gray-400">Lampiran Dokumen</h4>
              </div>
              {viewerUrl && (
                <a 
                  href={viewerUrl} 
                  target="_blank" 
                  rel="noreferrer"
                  className="text-[10px] font-bold text-[#006E62] hover:underline flex items-center gap-1"
                >
                  <ExternalLink size={12} /> BUKA DI TAB BARU
                </a>
              )}
            </div>

            <div className="aspect-video bg-gray-100 rounded border border-dashed border-gray-200 flex items-center justify-center overflow-hidden relative group">
              {contract.file_id ? (
                isImg ? (
                  <img 
                    src={fileUrl!} 
                    alt="Preview Kontrak" 
                    className="w-full h-full object-contain"
                    referrerPolicy="no-referrer"
                  />
                ) : (
                  <div className="flex flex-col items-center gap-3 text-gray-400">
                    <FileText size={48} strokeWidth={1} />
                    <p className="text-[10px] font-bold uppercase tracking-widest">Dokumen PDF / Non-Gambar</p>
                    <a 
                      href={viewerUrl!} 
                      target="_blank" 
                      rel="noreferrer"
                      className="px-6 py-2 bg-[#006E62] text-white rounded text-[10px] font-bold uppercase tracking-widest hover:bg-[#005a50] transition-colors shadow-md"
                    >
                      Lihat Dokumen Lengkap
                    </a>
                  </div>
                )
              ) : (
                <div className="flex flex-col items-center gap-2 text-gray-300">
                  <FileBadge size={48} strokeWidth={1} />
                  <p className="text-[10px] font-bold uppercase tracking-widest">Belum ada lampiran file</p>
                </div>
              )}
            </div>
          </div>
        </div>

        <div className="px-6 py-4 border-t border-gray-100 flex justify-end bg-white shrink-0">
          <button 
            onClick={onClose} 
            className="px-8 py-2 bg-gray-100 text-gray-600 rounded text-xs font-bold uppercase hover:bg-gray-200 transition-colors"
          >
            Tutup
          </button>
        </div>
      </div>
    </div>
  );
};

export default ContractDetailModal;
