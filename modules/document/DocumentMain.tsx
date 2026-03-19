
import React, { useState, useEffect } from 'react';
import { Plus, Search, Files, FileText, Image as ImageIcon, ExternalLink, Trash2, Filter, FolderOpen, Clock, Users } from 'lucide-react';
import Swal from 'sweetalert2';
import { documentService } from '../../services/documentService';
import { DigitalDocument, AuthUser } from '../../types';
import { googleDriveService } from '../../services/googleDriveService';
import DocumentForm from './DocumentForm';
import { CardSkeleton } from '../../components/Common/Skeleton';
import LoadingSpinner from '../../components/Common/LoadingSpinner';

interface DocumentMainProps {
  user?: AuthUser;
}

const DocumentMain: React.FC<DocumentMainProps> = ({ user }) => {
  const [documents, setDocuments] = useState<DigitalDocument[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [showForm, setShowForm] = useState(false);
  const [editingDoc, setEditingDoc] = useState<DigitalDocument | null>(null);

  const isAdmin = user?.role === 'admin' || user?.is_hr_admin || user?.is_performance_admin || user?.is_finance_admin;

  useEffect(() => {
    fetchDocuments();
  }, []);

  const fetchDocuments = async () => {
    try {
      setIsLoading(true);
      const data = await documentService.getAll();
      setDocuments(data);
    } catch (error) {
      Swal.fire('Gagal', 'Gagal memuat data dokumen', 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleDelete = async (id: string) => {
    const result = await Swal.fire({
      title: 'Hapus Dokumen?',
      text: "Metadata dan pengaturan akses akan dihapus dari sistem.",
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
        await documentService.delete(id);
        setDocuments(prev => prev.filter(d => d.id !== id));
        Swal.fire('Terhapus!', 'Dokumen telah dihapus.', 'success');
      } catch (error) {
        Swal.fire('Gagal', 'Gagal menghapus data', 'error');
      } finally {
        setIsSaving(false);
      }
    }
  };

  const filteredDocs = documents.filter(doc => 
    `${doc.name} ${doc.doc_type} ${doc.description}`.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const getFileIcon = (type: string) => {
    const t = type.toLowerCase();
    if (t.includes('gambar') || t.includes('image') || t.includes('foto')) return <ImageIcon className="text-blue-500" size={24} />;
    return <FileText className="text-emerald-500" size={24} />;
  };

  return (
    <div className="space-y-6">
      {isSaving && <LoadingSpinner />}

      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="relative flex-1 max-w-md">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
          <input
            type="text"
            placeholder="Cari Dokumen (Nama, Jenis, Deskripsi)..."
            className="w-full pl-10 pr-4 py-2 bg-gray-50 border border-gray-200 rounded-md focus:outline-none focus:ring-2 focus:ring-[#006E62] focus:border-transparent transition-all text-sm"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </div>
        
        <div className="flex items-center gap-2">
          <button className="p-2 border border-gray-200 rounded-md text-gray-500 hover:bg-gray-50 transition-colors">
            <Filter size={18} />
          </button>
          {isAdmin && (
            <button 
              onClick={() => { setEditingDoc(null); setShowForm(true); }}
              className="flex items-center gap-2 bg-[#006E62] text-white px-4 py-2 rounded-md hover:bg-[#005a50] transition-colors shadow-sm"
            >
              <Plus size={18} />
              <span className="font-medium text-sm">Unggah Dokumen</span>
            </button>
          )}
        </div>
      </div>

      {isLoading ? (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
          {[1, 2, 3].map(i => <CardSkeleton key={i} />)}
        </div>
      ) : filteredDocs.length === 0 ? (
        <div className="flex flex-col items-center justify-center py-20 text-gray-400 bg-gray-50 rounded-md border border-dashed border-gray-200">
          <FolderOpen size={48} strokeWidth={1} className="mb-4" />
          <p className="text-lg font-medium">Belum ada dokumen digital.</p>
          <p className="text-sm">Mulai unggah SOP, Kebijakan, atau Form Internal Anda.</p>
        </div>
      ) : (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
          {filteredDocs.map(doc => (
            <div 
              key={doc.id}
              className="bg-white border border-gray-100 p-5 rounded-md shadow-sm hover:shadow-md transition-all border-l-4 border-l-[#006E62] group"
            >
              <div className="flex justify-between items-start mb-4">
                <div className="p-2 bg-gray-50 rounded-md">
                   {getFileIcon(doc.doc_type)}
                </div>
                <div className="flex gap-1">
                  <a 
                    href={googleDriveService.getFileUrl(doc.file_id).replace('=s1600', '=s0')} 
                    target="_blank" 
                    rel="noreferrer"
                    className="p-1.5 text-gray-400 hover:text-[#006E62] transition-colors"
                  >
                    <ExternalLink size={16} />
                  </a>
                  {isAdmin && (
                    <>
                      <button 
                        onClick={() => { setEditingDoc(doc); setShowForm(true); }}
                        className="p-1.5 text-gray-400 hover:text-[#006E62] transition-colors"
                      >
                        <Files size={16} />
                      </button>
                      <button 
                        onClick={() => handleDelete(doc.id)}
                        className="p-1.5 text-gray-400 hover:text-red-500 transition-colors"
                      >
                        <Trash2 size={16} />
                      </button>
                    </>
                  )}
                </div>
              </div>

              <div className="space-y-1 mb-4">
                <h3 className="font-bold text-gray-800 text-sm group-hover:text-[#006E62] line-clamp-1">{doc.name}</h3>
                <span className="text-[10px] font-bold px-2 py-0.5 rounded uppercase bg-emerald-50 text-[#006E62]">{doc.doc_type}</span>
              </div>

              {doc.description && (
                <p className="text-[11px] text-gray-500 line-clamp-2 mb-4 h-8 italic">"{doc.description}"</p>
              )}

              <div className="flex items-center justify-between pt-4 border-t border-gray-50">
                 <div className="flex items-center gap-1.5 text-[10px] font-bold text-gray-400 uppercase tracking-tighter">
                   <Users size={12} className="text-[#00FFE4]" /> {doc.allowed_account_ids?.length || 0} Akun Akses
                 </div>
                 <div className="flex items-center gap-1.5 text-[10px] font-bold text-gray-400 uppercase tracking-tighter">
                   <Clock size={12} /> {new Date(doc.created_at!).toLocaleDateString('id-ID')}
                 </div>
              </div>
            </div>
          ))}
        </div>
      )}

      {showForm && (
        <DocumentForm 
          onClose={() => { setShowForm(false); setEditingDoc(null); }}
          onSuccess={() => { setShowForm(false); fetchDocuments(); }}
          initialData={editingDoc || undefined}
        />
      )}
    </div>
  );
};

export default DocumentMain;
