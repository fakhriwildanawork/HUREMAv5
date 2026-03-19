
import React, { useState } from 'react';
import { ShieldCheck, Lock, User, Loader2, AlertCircle } from 'lucide-react';
import { authService } from '../../services/authService';
import { AuthUser } from '../../types';
import { LOGO_ICON, Client_Name } from '../../assets';

interface LoginProps {
  onLoginSuccess: (user: AuthUser) => void;
}

const Login: React.FC<LoginProps> = ({ onLoginSuccess }) => {
  const [accessCode, setAccessCode] = useState('');
  const [password, setPassword] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setIsLoading(true);

    try {
      const user = await authService.login(accessCode, password);
      
      // Check for desktop access restriction
      const isAdmin = user.role === 'admin' || user.is_hr_admin || user.is_performance_admin || user.is_finance_admin;
      if (window.innerWidth >= 768 && !isAdmin) {
        authService.logout(); // Clear session if it was set
        throw new Error('Akses Desktop hanya untuk Admin. Silakan gunakan aplikasi mobile.');
      }

      onLoginSuccess(user);
    } catch (err: any) {
      setError(err.message || 'Gagal masuk. Periksa kembali kredensial Anda.');
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-white p-4">
      <div className="w-full max-w-md bg-white rounded-3xl border border-gray-100 shadow-[0_20px_50px_rgba(0,110,98,0.1)] overflow-hidden animate-in fade-in zoom-in duration-500">
        <div className="p-8">
          <div className="flex flex-col items-center mb-8">
            <img src={LOGO_ICON} alt="Logo" className="w-16 h-16 object-contain mb-4" />
            <h1 className="text-2xl font-black text-gray-800 tracking-tighter">HUREMA</h1>
            <p className="text-[10px] font-bold text-gray-400 uppercase tracking-[0.2em] mt-1">{Client_Name}</p>
          </div>

          <form onSubmit={handleSubmit} className="space-y-5">
            {error && (
              <div className="p-4 bg-rose-50 border border-rose-100 rounded-xl flex gap-3 items-center animate-shake">
                <AlertCircle className="text-rose-500 shrink-0" size={18} />
                <p className="text-xs font-bold text-rose-600 uppercase leading-tight tracking-tight">{error}</p>
              </div>
            )}

            <div className="space-y-1.5">
              <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Kode Akses</label>
              <div className="relative group">
                <div className="absolute left-4 top-1/2 -translate-y-1/2 text-gray-300 group-focus-within:text-[#006E62] transition-colors">
                  <User size={18} />
                </div>
                <input 
                  required
                  type="text"
                  value={accessCode}
                  onChange={(e) => setAccessCode(e.target.value)}
                  placeholder=""
                  className="w-full pl-12 pr-4 py-3.5 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] transition-all text-sm font-bold text-gray-700"
                />
              </div>
            </div>

            <div className="space-y-1.5">
              <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-1">Password</label>
              <div className="relative group">
                <div className="absolute left-4 top-1/2 -translate-y-1/2 text-gray-300 group-focus-within:text-[#006E62] transition-colors">
                  <Lock size={18} />
                </div>
                <input 
                  required
                  type="password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="••••••••"
                  className="w-full pl-12 pr-4 py-3.5 bg-gray-50 border border-gray-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#006E62]/20 focus:border-[#006E62] transition-all text-sm font-bold text-gray-700"
                />
              </div>
            </div>

            <button 
              type="submit"
              disabled={isLoading}
              className="w-full py-4 bg-[#006E62] hover:bg-[#005a50] text-white rounded-xl font-black text-xs uppercase tracking-[0.2em] shadow-lg shadow-[#006E62]/20 active:scale-[0.98] transition-all flex items-center justify-center gap-3 disabled:opacity-50"
            >
              {isLoading ? (
                <Loader2 className="animate-spin" size={18} />
              ) : (
                <>
                  <ShieldCheck size={18} />
                  MASUK SEKARANG
                </>
              )}
            </button>
          </form>

          <div className="mt-10 pt-6 border-t border-gray-50 text-center">
            <p className="text-[10px] font-bold text-gray-300 uppercase tracking-widest">v1.0.0 • Protected by HUREMA AI Security</p>
          </div>
        </div>
      </div>
    </div>
  );
};

export default Login;
