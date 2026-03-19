
import React from 'react';
import { X, AlertTriangle, FileText, Paperclip, ExternalLink, Calendar, Info } from 'lucide-react';
import { WarningLog } from '../../types';
import { googleDriveService } from '../../services/googleDriveService';

interface WarningDetailModalProps {
  log: WarningLog;
  onClose: () => void;
}

const WarningDetailModal: React.FC<WarningDetailModalProps> = ({ log, onClose }) => {
  const formatDate = (dateStr?: string | null) => {
    if (!dateStr) return '-';
    return new Date(dateStr).toLocaleDateString('id-ID', {
      day: 'numeric',
      month: 'long',
      year: 'numeric'
    });
  };

  const isImage = (fileId: string) => {
    if (!fileId.includes('|')) return true;
    const [, name] = fileId.split('|');
    return /\.(jpg|jpeg|png|webp|gif|svg|bmp)$/i.test(name);
  };

  const isImg = log.file_id ? isImage(log.file_id) : false;
  const fileUrl = log.file_id ? googleDriveService.getFileUrl(log.file_id, isImg) : null;
  const viewerUrl = log.file_id ? googleDriveService.getViewerUrl(log.file_id) : null;

  const getWarningColor = (type: string) => {
    switch (type) {
      case 'Teguran': return 'text-yellow-600 bg-yellow-50 border-yellow-100';
      case 'SP1': return 'text-orange-600 bg-orange-50 border-orange-100';
      case 'SP2': return 'text-red-500 bg-red-50 border-red-100';
      case 'SP3': return 'text-red-700 bg-red-50 border-red-200';
      default: return 'text-gray-600 bg-gray-50 border-gray-100';
    }
  };

  return (
    <div className="fixed inset-0 z-[250] bg-black/60 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in duration-200">
      <div className="bg-white rounded-md shadow-2xl w-full max-w-2xl overflow-hidden flex flex-col max-h-[90vh] animate-in zoom-in duration-200">
        <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between bg-white shrink-0">
          <div>
            <h3 className="text-base font-bold text-[#006E62]">Detail Kedisiplinan</h3>
            <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Informasi Pelanggaran & Sanksi</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors p-1">
            <X size={20} />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-6 space-y-6 bg-gray-50/30">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="space-y-4">
              <div className={`p-4 rounded border shadow-sm ${getWarningColor(log.warning_type)}`}>
                <div className="flex items-center gap-2 mb-1">
                  <AlertTriangle size={14} />
                  <p className="text-[9px] font-bold uppercase tracking-widest opacity-70">Jenis Sanksi</p>
                </div>
                <p className="text-sm font-bold">{log.warning_type}</p>
              </div>
              <div className="bg-white p-4 rounded border border-gray-100 shadow-sm">
                <div className="flex items-center gap-2 mb-1">
                  <Calendar size={14} className="text-[#006E62]" />
                  <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest">Tanggal Kejadian</p>
                </div>
                <p className="text-sm font-bold text-gray-700">{formatDate(log.issue_date)}</p>
              </div>
            </div>

            <div className="bg-white p-4 rounded border border-gray-100 shadow-sm flex flex-col">
              <div className="flex items-center gap-2 mb-2">
                <Info size={14} className="text-[#006E62]" />
                <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest">Alasan / Kronologi</p>
              </div>
              <p className="text-xs text-gray-600 leading-relaxed italic flex-1">
                {log.reason || 'Tidak ada keterangan tambahan.'}
              </p>
            </div>
          </div>

          <div className="bg-white p-4 rounded border border-gray-100 shadow-sm">
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-2">
                <Paperclip size={16} className="text-[#006E62]" />
                <h4 className="text-[10px] font-bold uppercase tracking-widest text-gray-400">Lampiran Berita Acara</h4>
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
              {log.file_id ? (
                isImg ? (
                  <img 
                    src={fileUrl!} 
                    alt="Preview Sanksi" 
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
                  <FileText size={48} strokeWidth={1} />
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

export default WarningDetailModal;
