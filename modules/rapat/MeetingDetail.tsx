import React, { useState } from 'react';
import { X, Calendar, Clock, MapPin, Users, FileText, Paperclip, Link as LinkIcon, ExternalLink, CheckCircle2, User, Map as MapIcon, Eye } from 'lucide-react';
import { Meeting, MeetingNote } from '../../types';
import { MapContainer, TileLayer, Marker } from 'react-leaflet';
import L from 'leaflet';

interface MeetingDetailProps {
  meeting: Meeting;
  onClose: () => void;
}

const MeetingDetail: React.FC<MeetingDetailProps> = ({ meeting, onClose }) => {
  const [viewingNote, setViewingNote] = useState<MeetingNote | null>(null);
  const mapPosition: [number, number] = [meeting.latitude || -6.2, meeting.longitude || 106.8];

  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
      <div className="bg-white w-full max-w-4xl rounded-3xl overflow-hidden shadow-2xl flex flex-col animate-in fade-in zoom-in duration-300 max-h-[90vh]">
        <div className="p-6 border-b border-gray-100 flex items-center justify-between shrink-0 bg-emerald-50/30">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-emerald-100 text-emerald-600 flex items-center justify-center">
              <CheckCircle2 size={20} />
            </div>
            <div>
              <h3 className="text-lg font-bold text-gray-800 tracking-tight">{meeting.title}</h3>
              <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Arsip Notulensi Rapat Selesai</p>
            </div>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-gray-100 rounded-xl transition-colors text-gray-400">
            <X size={24} />
          </button>
        </div>

        <div className="p-6 overflow-y-auto space-y-8 scrollbar-none">
          {/* Header Info */}
          <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
            <div className="p-4 bg-gray-50 rounded-2xl border border-gray-100">
              <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Tanggal</p>
              <p className="text-xs font-bold text-gray-700">{new Date(meeting.scheduled_at).toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' })}</p>
            </div>
            <div className="p-4 bg-gray-50 rounded-2xl border border-gray-100">
              <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Waktu</p>
              <p className="text-xs font-bold text-gray-700">{new Date(meeting.scheduled_at).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })} WIB</p>
            </div>
            <div className="p-4 bg-gray-50 rounded-2xl border border-gray-100">
              <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Tipe</p>
              <p className="text-xs font-bold text-emerald-600 uppercase">{meeting.location_type}</p>
            </div>
            <div className="p-4 bg-gray-50 rounded-2xl border border-gray-100">
              <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Status</p>
              <p className="text-xs font-bold text-emerald-600 uppercase">Selesai</p>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
            {/* Main Content */}
            <div className="md:col-span-2 space-y-8">
              <div className="space-y-3">
                <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1 flex items-center gap-2">
                  <FileText size={14} /> Notulensi Rapat
                </h4>
                <div className="space-y-4">
                  {meeting.notes && meeting.notes.length > 0 ? (
                    meeting.notes.map((note, idx) => (
                      <div key={note.id} className="p-6 bg-white border border-gray-100 rounded-2xl shadow-sm space-y-4 group relative">
                        <div className="absolute top-6 right-6 opacity-0 group-hover:opacity-100 transition-opacity">
                          <button 
                            onClick={() => setViewingNote(note)}
                            className="p-2 text-gray-400 hover:text-emerald-600 hover:bg-emerald-50 rounded-xl transition-all"
                            title="Lihat Detail"
                          >
                            <Eye size={16} />
                          </button>
                        </div>
                        <div className="flex items-start gap-3">
                          <div className="w-6 h-6 rounded-lg bg-gray-50 flex items-center justify-center text-[10px] font-bold text-gray-400 shrink-0">
                            {idx + 1}
                          </div>
                          <p className="text-sm text-gray-700 leading-relaxed font-medium whitespace-pre-wrap">{note.content}</p>
                        </div>

                        {(note.attachments.length > 0 || note.links.length > 0) && (
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 pt-4 border-t border-gray-50">
                            {note.attachments.length > 0 && (
                              <div className="space-y-2">
                                <p className="text-[8px] font-bold text-gray-400 uppercase tracking-widest">Lampiran File</p>
                                <div className="space-y-1">
                                  {note.attachments.map((fileId, fIdx) => (
                                    <a 
                                      key={fIdx}
                                      href={`https://drive.google.com/file/d/${fileId}/view`}
                                      target="_blank"
                                      rel="noopener noreferrer"
                                      className="flex items-center justify-between p-2 bg-gray-50 border border-gray-100 rounded-lg hover:bg-emerald-50 transition-colors group"
                                    >
                                      <div className="flex items-center gap-2">
                                        <Paperclip size={10} className="text-gray-400 group-hover:text-emerald-600" />
                                        <span className="text-[9px] font-bold text-gray-600 group-hover:text-emerald-600">Dokumentasi_{fIdx + 1}</span>
                                      </div>
                                      <ExternalLink size={10} className="text-gray-300 group-hover:text-emerald-600" />
                                    </a>
                                  ))}
                                </div>
                              </div>
                            )}
                            {note.links.length > 0 && (
                              <div className="space-y-2">
                                <p className="text-[8px] font-bold text-gray-400 uppercase tracking-widest">Link Pendukung</p>
                                <div className="space-y-1">
                                  {note.links.map((link, lIdx) => (
                                    <a 
                                      key={lIdx}
                                      href={link}
                                      target="_blank"
                                      rel="noopener noreferrer"
                                      className="flex items-center justify-between p-2 bg-gray-50 border border-gray-100 rounded-lg hover:bg-emerald-50 transition-colors group"
                                    >
                                      <div className="flex items-center gap-2">
                                        <LinkIcon size={10} className="text-gray-400 group-hover:text-emerald-600" />
                                        <span className="text-[9px] font-bold text-gray-600 group-hover:text-emerald-600 truncate max-w-[120px]">{link}</span>
                                      </div>
                                      <ExternalLink size={10} className="text-gray-300 group-hover:text-emerald-600" />
                                    </a>
                                  ))}
                                </div>
                              </div>
                            )}
                          </div>
                        )}
                        <div className="text-[8px] font-bold text-gray-300 uppercase tracking-widest text-right">
                          Dicatat pada {new Date(note.created_at).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })} WIB
                        </div>
                      </div>
                    ))
                  ) : (
                    <div className="p-12 bg-white border border-gray-100 rounded-2xl shadow-sm flex flex-col items-center justify-center text-gray-400">
                      <FileText size={48} strokeWidth={1} className="mb-4 opacity-20" />
                      <p className="text-[10px] font-bold uppercase tracking-widest">Tidak ada notulensi yang dicatat.</p>
                    </div>
                  )}
                </div>
              </div>
            </div>

            {/* Sidebar Info */}
            <div className="space-y-8">
              <div className="space-y-3">
                <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Lokasi Rapat</h4>
                <div className="p-4 bg-gray-50 border border-gray-100 rounded-2xl space-y-3">
                  <div className="flex items-start gap-3">
                    <MapPin size={16} className="text-emerald-600 shrink-0 mt-0.5" />
                    <p className="text-xs font-bold text-gray-700 leading-tight">{meeting.location_detail}</p>
                  </div>
                  {meeting.location_type === 'Offline' && meeting.latitude && (
                    <div className="h-[150px] rounded-xl overflow-hidden border border-gray-200 relative z-0">
                      <MapContainer center={mapPosition} zoom={15} scrollWheelZoom={false} style={{ height: '100%', width: '100%' }}>
                        <TileLayer url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png" />
                        <Marker position={mapPosition} />
                      </MapContainer>
                    </div>
                  )}
                </div>
              </div>

              <div className="space-y-3">
                <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Peserta ({meeting.participants?.length || 0})</h4>
                <div className="space-y-2 max-h-[200px] overflow-y-auto scrollbar-thin pr-1">
                  {meeting.participants?.map(acc => (
                    <div key={acc.id} className="flex items-center gap-3 p-2 bg-white border border-gray-50 rounded-xl">
                      <div className="w-7 h-7 rounded-full bg-gray-100 flex items-center justify-center text-gray-500 text-[10px] font-bold">
                        {acc.full_name.charAt(0)}
                      </div>
                      <div className="overflow-hidden">
                        <p className="text-[10px] font-bold text-gray-800 truncate">{acc.full_name}</p>
                        <p className="text-[8px] font-bold text-gray-400 uppercase">{acc.internal_nik}</p>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              <div className="space-y-3">
                <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Notulen ({meeting.notulens?.length || 0})</h4>
                <div className="space-y-2">
                  {meeting.notulens?.map(acc => (
                    <div key={acc.id} className="flex items-center gap-3 p-2 bg-amber-50 border border-amber-100 rounded-xl">
                      <div className="w-7 h-7 rounded-full bg-amber-100 flex items-center justify-center text-amber-600 text-[10px] font-bold">
                        {acc.full_name.charAt(0)}
                      </div>
                      <div className="overflow-hidden">
                        <p className="text-[10px] font-bold text-gray-800 truncate">{acc.full_name}</p>
                        <p className="text-[8px] font-bold text-amber-600 uppercase">Notulen</p>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        </div>

        <div className="p-6 border-t border-gray-100 bg-gray-50/50 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-2 text-[10px] font-bold text-gray-400 uppercase tracking-widest">
            <User size={14} />
            <span>Host: {meeting.creator?.full_name}</span>
          </div>
          <button 
            onClick={onClose}
            className="px-8 py-2 bg-gray-800 text-white rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-black transition-all"
          >
            Tutup
          </button>
        </div>
      </div>

      {/* Note Detail Modal */}
      {viewingNote && (
        <div className="fixed inset-0 bg-black/40 backdrop-blur-sm z-[60] flex items-center justify-center p-4 animate-in fade-in duration-200">
          <div className="bg-white w-full max-w-2xl rounded-3xl overflow-hidden shadow-2xl flex flex-col animate-in zoom-in duration-300">
            <div className="p-6 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-emerald-100 text-emerald-600 flex items-center justify-center">
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
                        className="flex items-center justify-between p-3 bg-white border border-gray-100 rounded-xl hover:border-emerald-200 hover:bg-emerald-50/30 transition-all group"
                      >
                        <div className="flex items-center gap-3">
                          <Paperclip size={14} className="text-gray-400 group-hover:text-emerald-600" />
                          <span className="text-xs font-bold text-gray-600 group-hover:text-emerald-600">Dokumentasi_{idx + 1}</span>
                        </div>
                        <ExternalLink size={14} className="text-gray-300 group-hover:text-emerald-600" />
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
                        className="flex items-center justify-between p-3 bg-white border border-gray-100 rounded-xl hover:border-emerald-200 hover:bg-emerald-50/30 transition-all group"
                      >
                        <div className="flex items-center gap-3">
                          <LinkIcon size={14} className="text-gray-400 group-hover:text-emerald-600" />
                          <span className="text-xs font-bold text-gray-600 group-hover:text-emerald-600 truncate max-w-[120px]">{link}</span>
                        </div>
                        <ExternalLink size={14} className="text-gray-300 group-hover:text-emerald-600" />
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

export default MeetingDetail;
