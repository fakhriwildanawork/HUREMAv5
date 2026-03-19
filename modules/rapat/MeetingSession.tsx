import React, { useState, useRef, useEffect } from 'react';
import { X, Save, Paperclip, Link as LinkIcon, Plus, Trash2, Loader2, Play, StopCircle, Clock, FileText, Send, AlertCircle, Edit2, Check, ExternalLink, Eye } from 'lucide-react';
import { Meeting, MeetingNote } from '../../types';
import { googleDriveService } from '../../services/googleDriveService';
import { meetingService } from '../../services/meetingService';
import Swal from 'sweetalert2';

interface MeetingSessionProps {
  meeting: Meeting;
  onClose: () => void;
  onEnd: () => void;
}

const MeetingSession: React.FC<MeetingSessionProps> = ({ meeting, onClose, onEnd }) => {
  const [notes, setNotes] = useState<MeetingNote[]>([]);
  const [isLoadingNotes, setIsLoadingNotes] = useState(true);
  
  // Form for new/editing note
  const [isEditing, setIsEditing] = useState(false);
  const [editingNoteId, setEditingNoteId] = useState<string | null>(null);
  const [viewingNote, setViewingNote] = useState<MeetingNote | null>(null);
  const [currentContent, setCurrentContent] = useState('');
  const [currentLinks, setCurrentLinks] = useState<string[]>(['']);
  const [currentAttachments, setCurrentAttachments] = useState<string[]>([]);
  
  const [isUploading, setIsUploading] = useState(false);
  const [elapsedTime, setElapsedTime] = useState(0);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    fetchNotes();
    const startTime = meeting.started_at ? new Date(meeting.started_at).getTime() : Date.now();
    const interval = setInterval(() => {
      setElapsedTime(Math.floor((Date.now() - startTime) / 1000));
    }, 1000);
    return () => clearInterval(interval);
  }, [meeting.id, meeting.started_at]);

  const fetchNotes = async () => {
    try {
      setIsLoadingNotes(true);
      const data = await meetingService.getNotes(meeting.id);
      setNotes(data);
    } catch (error) {
      console.error('Failed to fetch notes:', error);
    } finally {
      setIsLoadingNotes(false);
    }
  };

  const formatTime = (seconds: number) => {
    const h = Math.floor(seconds / 3600);
    const m = Math.floor((seconds % 3600) / 60);
    const s = seconds % 60;
    return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
  };

  const handleAddLink = () => setCurrentLinks([...currentLinks, '']);
  const handleRemoveLink = (index: number) => setCurrentLinks(currentLinks.filter((_, i) => i !== index));
  const handleLinkChange = (index: number, value: string) => {
    const newLinks = [...currentLinks];
    newLinks[index] = value;
    setCurrentLinks(newLinks);
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    try {
      setIsUploading(true);
      for (let i = 0; i < files.length; i++) {
        const fileId = await googleDriveService.uploadFile(files[i]);
        setCurrentAttachments(prev => [...prev, fileId]);
      }
    } catch (error) {
      alert('Gagal mengunggah file.');
    } finally {
      setIsUploading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const handleRemoveAttachment = (index: number) => setCurrentAttachments(currentAttachments.filter((_, i) => i !== index));

  const resetForm = () => {
    setIsEditing(false);
    setEditingNoteId(null);
    setCurrentContent('');
    setCurrentLinks(['']);
    setCurrentAttachments([]);
  };

  const handleSaveNote = async () => {
    if (!currentContent.trim()) return;
    try {
      const links = currentLinks.filter(l => l.trim() !== '');
      if (editingNoteId) {
        await meetingService.updateNote(editingNoteId, currentContent, currentAttachments, links);
      } else {
        await meetingService.addNote(meeting.id, currentContent, currentAttachments, links);
      }
      await fetchNotes();
      resetForm();
    } catch (error) {
      Swal.fire('Gagal', 'Gagal menyimpan notulensi.', 'error');
    }
  };

  const handleEditNote = (note: MeetingNote) => {
    setIsEditing(true);
    setEditingNoteId(note.id);
    setCurrentContent(note.content);
    setCurrentLinks(note.links.length > 0 ? note.links : ['']);
    setCurrentAttachments(note.attachments);
  };

  const handleDeleteNote = async (id: string) => {
    const result = await Swal.fire({
      title: 'Hapus Poin?',
      text: "Poin notulensi ini akan dihapus.",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
      confirmButtonText: 'Ya, Hapus!'
    });

    if (result.isConfirmed) {
      try {
        await meetingService.deleteNote(id);
        await fetchNotes();
      } catch (error) {
        Swal.fire('Gagal', 'Gagal menghapus poin.', 'error');
      }
    }
  };

  const handleEndMeeting = async () => {
    const result = await Swal.fire({
      title: 'Akhiri Rapat?',
      text: "Sesi rapat akan diakhiri dan notulensi akan dikunci.",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonText: 'Ya, Akhiri!'
    });

    if (result.isConfirmed) {
      onEnd();
    }
  };

  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
      <div className="bg-white w-full max-w-6xl rounded-3xl overflow-hidden shadow-2xl flex flex-col animate-in fade-in zoom-in duration-300 h-[95vh]">
        <div className="p-6 border-b border-gray-100 flex items-center justify-between shrink-0 bg-amber-50/30">
          <div className="flex items-center gap-4">
            <div className="w-12 h-12 rounded-2xl bg-amber-100 text-amber-600 flex items-center justify-center animate-pulse">
              <Play size={24} />
            </div>
            <div>
              <h3 className="text-lg font-bold text-gray-800 tracking-tight">{meeting.title}</h3>
              <div className="flex items-center gap-3 mt-1">
                <span className="text-[10px] text-amber-600 font-bold uppercase tracking-widest flex items-center gap-1">
                  <Clock size={12} />
                  Durasi: {formatTime(elapsedTime)}
                </span>
                <span className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">•</span>
                <span className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Sesi Rapat Berlangsung</span>
              </div>
            </div>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-gray-100 rounded-xl transition-colors text-gray-400">
            <X size={24} />
          </button>
        </div>

        <div className="flex-1 flex overflow-hidden">
          {/* Left: Notes List */}
          <div className="flex-1 p-6 overflow-y-auto space-y-6 scrollbar-none border-r border-gray-100 bg-gray-50/30">
            <div className="flex items-center justify-between px-1">
              <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Daftar Notulensi ({notes.length})</h4>
              {!isEditing && (
                <button 
                  onClick={() => setIsEditing(true)}
                  className="text-[10px] font-bold text-amber-600 uppercase flex items-center gap-1 hover:underline"
                >
                  <Plus size={12} /> Tambah Notulensi
                </button>
              )}
            </div>

            {isLoadingNotes ? (
              <div className="flex flex-col items-center justify-center py-20">
                <Loader2 size={32} className="animate-spin text-gray-300 mb-2" />
                <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Memuat Catatan...</p>
              </div>
            ) : (
              <div className="space-y-4">
                {notes.map((note, idx) => (
                  <div key={note.id} className="bg-white border border-gray-100 p-5 rounded-2xl shadow-sm hover:shadow-md transition-all group relative">
                    <div className="absolute top-4 right-4 flex items-center gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                      <button 
                        onClick={() => setViewingNote(note)}
                        className="p-1.5 text-gray-400 hover:text-emerald-600 hover:bg-emerald-50 rounded-lg transition-all"
                        title="Lihat Detail"
                      >
                        <Eye size={14} />
                      </button>
                      <button 
                        onClick={() => handleEditNote(note)}
                        className="p-1.5 text-gray-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg transition-all"
                      >
                        <Edit2 size={14} />
                      </button>
                      <button 
                        onClick={() => handleDeleteNote(note.id)}
                        className="p-1.5 text-gray-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-all"
                      >
                        <Trash2 size={14} />
                      </button>
                    </div>

                    <div className="flex items-start gap-3 mb-3">
                      <div className="w-6 h-6 rounded-lg bg-gray-100 flex items-center justify-center text-[10px] font-bold text-gray-400 shrink-0">
                        {idx + 1}
                      </div>
                      <p className="text-sm text-gray-700 leading-relaxed font-medium whitespace-pre-wrap">{note.content}</p>
                    </div>

                    {(note.attachments.length > 0 || note.links.length > 0) && (
                      <div className="flex flex-wrap gap-2 mt-4 pt-4 border-t border-gray-50">
                        {note.attachments.map((fileId, fIdx) => (
                          <a 
                            key={fIdx}
                            href={`https://drive.google.com/file/d/${fileId}/view`}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="flex items-center gap-1.5 px-2 py-1 bg-gray-50 text-gray-500 rounded text-[9px] font-bold uppercase hover:bg-amber-50 hover:text-amber-600 transition-colors"
                          >
                            <Paperclip size={10} />
                            Bukti_{fIdx + 1}
                          </a>
                        ))}
                        {note.links.map((link, lIdx) => (
                          <a 
                            key={lIdx}
                            href={link}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="flex items-center gap-1.5 px-2 py-1 bg-gray-50 text-gray-500 rounded text-[9px] font-bold uppercase hover:bg-amber-50 hover:text-amber-600 transition-colors"
                          >
                            <LinkIcon size={10} />
                            Link_{lIdx + 1}
                          </a>
                        ))}
                      </div>
                    )}
                    <div className="mt-3 text-[8px] font-bold text-gray-300 uppercase tracking-widest">
                      {new Date(note.created_at).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })} WIB
                    </div>
                  </div>
                ))}

                {notes.length === 0 && !isEditing && (
                  <div className="flex flex-col items-center justify-center py-20 text-gray-400 bg-white rounded-2xl border border-dashed border-gray-200">
                    <FileText size={48} strokeWidth={1} className="mb-4 opacity-20" />
                    <p className="text-[10px] font-bold uppercase tracking-widest">Belum ada notulensi dicatat</p>
                    <button 
                      onClick={() => setIsEditing(true)}
                      className="mt-4 px-4 py-2 bg-amber-600 text-white rounded-xl text-[10px] font-bold uppercase tracking-widest hover:bg-amber-700 transition-all"
                    >
                      Mulai Mencatat
                    </button>
                  </div>
                )}
              </div>
            )}
          </div>

          {/* Right: Editor / Form */}
          <div className={`w-96 p-6 overflow-y-auto space-y-6 scrollbar-none transition-all ${isEditing ? 'translate-x-0 opacity-100' : 'translate-x-full opacity-0 absolute right-0'}`}>
            <div className="flex items-center justify-between px-1">
              <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">
                {editingNoteId ? 'Edit Notulensi' : 'Tambah Notulensi'}
              </h4>
              <button onClick={resetForm} className="text-gray-400 hover:text-gray-600">
                <X size={16} />
              </button>
            </div>

            <div className="space-y-4">
              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Isi Catatan</label>
                <textarea 
                  value={currentContent}
                  onChange={(e) => setCurrentContent(e.target.value)}
                  placeholder="Tuliskan poin pembahasan..."
                  className="w-full h-48 px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-amber-600/20 focus:border-amber-600 transition-all text-sm font-medium text-gray-700 resize-none"
                />
              </div>

              <div className="space-y-3">
                <div className="flex items-center justify-between px-1">
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Lampiran (File)</label>
                  <button 
                    type="button" 
                    onClick={() => fileInputRef.current?.click()} 
                    disabled={isUploading}
                    className="text-[10px] font-bold text-amber-600 uppercase flex items-center gap-1 hover:underline disabled:opacity-50"
                  >
                    <Plus size={12} /> Tambah
                  </button>
                  <input 
                    ref={fileInputRef}
                    type="file"
                    multiple
                    onChange={handleFileUpload}
                    className="hidden"
                  />
                </div>

                {isUploading && (
                  <div className="flex items-center gap-2 px-3 py-2 bg-amber-50 text-amber-600 rounded-xl animate-pulse">
                    <Loader2 size={14} className="animate-spin" />
                    <span className="text-[10px] font-bold uppercase tracking-widest">Mengunggah...</span>
                  </div>
                )}

                <div className="space-y-2">
                  {currentAttachments.map((fileId, idx) => (
                    <div key={idx} className="flex items-center justify-between px-3 py-2 bg-gray-50 border border-gray-100 rounded-xl group">
                      <div className="flex items-center gap-2 overflow-hidden">
                        <Paperclip size={12} className="text-gray-400 shrink-0" />
                        <span className="text-[10px] text-gray-600 truncate">File_{idx + 1}</span>
                      </div>
                      <button 
                        type="button" 
                        onClick={() => handleRemoveAttachment(idx)} 
                        className="p-1.5 text-gray-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-all"
                      >
                        <Trash2 size={14} />
                      </button>
                    </div>
                  ))}
                </div>
              </div>

              <div className="space-y-3">
                <div className="flex items-center justify-between px-1">
                  <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Link Pendukung</label>
                  <button type="button" onClick={handleAddLink} className="text-[10px] font-bold text-amber-600 uppercase flex items-center gap-1 hover:underline">
                    <Plus size={12} /> Tambah
                  </button>
                </div>
                <div className="space-y-2">
                  {currentLinks.map((link, idx) => (
                    <div key={idx} className="flex gap-2">
                      <div className="relative flex-1">
                        <LinkIcon className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-300" size={14} />
                        <input 
                          type="url"
                          value={link}
                          onChange={(e) => handleLinkChange(idx, e.target.value)}
                          placeholder="https://..."
                          className="w-full pl-9 pr-4 py-2 bg-gray-50 border border-gray-100 rounded-lg focus:outline-none focus:ring-2 focus:ring-amber-600/20 text-[10px] font-medium"
                        />
                      </div>
                      {currentLinks.length > 1 && (
                        <button type="button" onClick={() => handleRemoveLink(idx)} className="p-2 text-rose-500 hover:bg-rose-50 rounded-lg transition-colors">
                          <Trash2 size={16} />
                        </button>
                      )}
                    </div>
                  ))}
                </div>
              </div>

              <button 
                onClick={handleSaveNote}
                disabled={!currentContent.trim() || isUploading}
                className="w-full py-3 bg-amber-600 text-white rounded-xl font-bold text-xs uppercase tracking-widest hover:bg-amber-700 transition-all shadow-lg shadow-amber-600/20 flex items-center justify-center gap-2 disabled:opacity-50"
              >
                <Check size={16} />
                {editingNoteId ? 'Perbarui Catatan' : 'Simpan Catatan'}
              </button>
            </div>
          </div>
        </div>

        <div className="p-6 border-t border-gray-100 bg-gray-50/50 flex gap-3 shrink-0">
          <button 
            type="button"
            onClick={onClose}
            className="flex-1 py-3 border border-gray-200 text-gray-600 rounded-xl font-bold text-xs uppercase tracking-widest hover:bg-white transition-all"
          >
            Keluar Sesi
          </button>
          <button 
            onClick={handleEndMeeting}
            className="flex-[2] py-3 bg-rose-600 text-white rounded-xl font-bold text-xs uppercase tracking-widest hover:bg-rose-700 transition-all shadow-lg shadow-rose-600/20 flex items-center justify-center gap-2"
          >
            <StopCircle size={18} />
            Akhiri Rapat & Kunci Notulensi
          </button>
        </div>
      </div>

      {/* Note Detail Modal */}
      {viewingNote && (
        <div className="fixed inset-0 bg-black/40 backdrop-blur-sm z-[60] flex items-center justify-center p-4 animate-in fade-in duration-200">
          <div className="bg-white w-full max-w-2xl rounded-3xl overflow-hidden shadow-2xl flex flex-col animate-in zoom-in duration-300">
            <div className="p-6 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-amber-100 text-amber-600 flex items-center justify-center">
                  <FileText size={20} />
                </div>
                <div>
                  <h3 className="text-sm font-bold text-gray-800 tracking-tight">Detail Notulensi</h3>
                  <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">
                    Dicatat pada {new Date(viewingNote.created_at).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })} WIB
                  </p>
                </div>
              </div>
              <button onClick={() => setViewingNote(null)} className="p-2 hover:bg-gray-100 rounded-xl transition-colors text-gray-400">
                <X size={20} />
              </button>
            </div>
            <div className="p-8 overflow-y-auto max-h-[70vh] space-y-8">
              <div className="space-y-3">
                <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Isi Catatan</h4>
                <div className="p-6 bg-gray-50 rounded-2xl border border-gray-100">
                  <p className="text-sm text-gray-700 leading-relaxed font-medium whitespace-pre-wrap">{viewingNote.content}</p>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-8">
                <div className="space-y-3">
                  <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Lampiran File</h4>
                  <div className="space-y-2">
                    {viewingNote.attachments.length > 0 ? viewingNote.attachments.map((fileId, idx) => (
                      <a 
                        key={idx}
                        href={`https://drive.google.com/file/d/${fileId}/view`}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="flex items-center justify-between p-3 bg-white border border-gray-100 rounded-xl hover:border-amber-200 hover:bg-amber-50/30 transition-all group"
                      >
                        <div className="flex items-center gap-3">
                          <Paperclip size={14} className="text-gray-400 group-hover:text-amber-600" />
                          <span className="text-xs font-bold text-gray-600 group-hover:text-amber-600">Dokumentasi_{idx + 1}</span>
                        </div>
                        <ExternalLink size={14} className="text-gray-300 group-hover:text-amber-600" />
                      </a>
                    )) : (
                      <p className="text-xs text-gray-400 italic px-1">Tidak ada lampiran.</p>
                    )}
                  </div>
                </div>
                <div className="space-y-3">
                  <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Link Pendukung</h4>
                  <div className="space-y-2">
                    {viewingNote.links.length > 0 ? viewingNote.links.map((link, idx) => (
                      <a 
                        key={idx}
                        href={link}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="flex items-center justify-between p-3 bg-white border border-gray-100 rounded-xl hover:border-amber-200 hover:bg-amber-50/30 transition-all group"
                      >
                        <div className="flex items-center gap-3">
                          <LinkIcon size={14} className="text-gray-400 group-hover:text-amber-600" />
                          <span className="text-xs font-bold text-gray-600 group-hover:text-amber-600 truncate max-w-[120px]">{link}</span>
                        </div>
                        <ExternalLink size={14} className="text-gray-300 group-hover:text-amber-600" />
                      </a>
                    )) : (
                      <p className="text-xs text-gray-400 italic px-1">Tidak ada link.</p>
                    )}
                  </div>
                </div>
              </div>
            </div>
            <div className="p-6 border-t border-gray-100 bg-gray-50/50 flex justify-end">
              <button 
                onClick={() => setViewingNote(null)}
                className="px-8 py-2 bg-gray-800 text-white rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-black transition-all"
              >
                Tutup
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default MeetingSession;
