import React, { useState } from 'react';
import { X, CheckCircle2, MessageSquare, Target, Info } from 'lucide-react';
import { KeyActivityReport } from '../../../types';

interface KeyActivityVerifyFormProps {
  report: KeyActivityReport;
  onClose: () => void;
  onSubmit: (data: { score: number; notes: string }) => void;
}

const KeyActivityVerifyForm: React.FC<KeyActivityVerifyFormProps> = ({ report, onClose, onSubmit }) => {
  const [score, setScore] = useState(100);
  const [notes, setNotes] = useState('');

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    onSubmit({ score, notes });
  };

  return (
    <div className="fixed inset-0 z-[70] flex items-center justify-center bg-black/40 p-4 backdrop-blur-sm">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md overflow-hidden animate-in zoom-in duration-200">
        <div className="px-6 py-4 border-b border-gray-50 flex items-center justify-between bg-gray-50/50">
          <div className="flex items-center gap-2">
            <CheckCircle2 size={18} className="text-[#006E62]" />
            <h3 className="text-sm font-bold text-gray-800 uppercase tracking-tight">Verifikasi Laporan</h3>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors">
            <X size={20} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-6 space-y-5">
          <div className="p-4 bg-emerald-50 rounded-2xl border border-emerald-100">
            <p className="text-[10px] font-bold text-emerald-600 uppercase tracking-widest mb-1">Aktivitas Pegawai</p>
            <p className="text-xs font-bold text-gray-800">{report.activity?.title}</p>
            <p className="text-[10px] text-emerald-600 font-medium mt-1">Oleh: {report.account?.full_name}</p>
          </div>

          <div className="space-y-2">
            <label className="flex items-center gap-2 text-[10px] font-bold text-gray-400 uppercase tracking-widest">
              <Target size={14} />
              Skor Penilaian (1 - 100)
            </label>
            <div className="flex items-center gap-4">
              <input
                type="range"
                min="0"
                max="100"
                step="5"
                className="flex-1 h-2 bg-gray-100 rounded-lg appearance-none cursor-pointer accent-[#006E62]"
                value={score}
                onChange={(e) => setScore(parseInt(e.target.value))}
              />
              <span className="text-lg font-bold text-[#006E62] w-12 text-right">{score}%</span>
            </div>
          </div>

          <div className="space-y-2">
            <label className="flex items-center gap-2 text-[10px] font-bold text-gray-400 uppercase tracking-widest">
              <MessageSquare size={14} />
              Catatan Evaluasi
            </label>
            <textarea
              required
              rows={3}
              placeholder="Berikan masukan atau catatan untuk pegawai..."
              className="w-full px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62] text-xs resize-none"
              value={notes}
              onChange={(e) => setNotes(e.target.value)}
            />
          </div>

          <div className="flex items-center gap-2 p-3 bg-blue-50 rounded-xl text-blue-600">
            <Info size={16} className="shrink-0" />
            <p className="text-[9px] font-medium leading-relaxed">
              Skor ini akan mempengaruhi Performance Score pegawai pada modul Key Activities.
            </p>
          </div>

          <div className="flex gap-3 pt-2">
            <button
              type="button"
              onClick={onClose}
              className="flex-1 px-4 py-2.5 bg-gray-50 text-gray-400 rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-gray-100 transition-all"
            >
              Batal
            </button>
            <button
              type="submit"
              className="flex-1 px-4 py-2.5 bg-[#006E62] text-white rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-[#005a50] transition-all shadow-lg shadow-[#006E62]/20"
            >
              Simpan Verifikasi
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};

export default KeyActivityVerifyForm;
