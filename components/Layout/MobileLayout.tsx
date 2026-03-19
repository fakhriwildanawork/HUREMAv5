
import React from 'react';
import { Bell, Home, Fingerprint, UserCircle, Menu } from 'lucide-react';
import { AuthUser } from '../../types';
import { googleDriveService } from '../../services/googleDriveService';

interface MobileLayoutProps {
  children: React.ReactNode;
  activeTab: string;
  setActiveTab: (tab: any) => void;
  user: AuthUser;
}

const MobileLayout: React.FC<MobileLayoutProps> = ({ children, activeTab, setActiveTab, user }) => {
  return (
    <div className="min-h-screen bg-white flex flex-col md:hidden">
      {/* Top Header */}
      <header className="h-16 flex items-center justify-between px-4 bg-white sticky top-0 z-40 border-b border-gray-50">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 bg-[#006E62] rounded-lg flex items-center justify-center text-white font-bold italic text-sm shadow-sm">H</div>
        </div>
        
        <div className="flex items-center gap-3">
          <span className="text-xs font-semibold text-gray-700 truncate max-w-[120px]">{user.full_name}</span>
          <div className="w-9 h-9 rounded-full bg-gray-100 border border-gray-200 overflow-hidden flex items-center justify-center text-gray-400 font-bold shadow-sm">
            {user.photo_google_id ? (
              <img 
                src={googleDriveService.getFileUrl(user.photo_google_id)} 
                alt="" 
                className="w-full h-full object-cover"
                referrerPolicy="no-referrer"
              />
            ) : (
              <UserCircle size={22} />
            )}
          </div>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="flex-1 overflow-y-auto p-6 scrollbar-none">
        {children}
      </main>

      {/* Bottom Navigation */}
      <nav className="h-20 bg-white border-t border-gray-100 flex items-center justify-around px-6 sticky bottom-0 z-40 pb-safe">
        <button 
          onClick={() => setActiveTab('dashboard')}
          className={`p-3 rounded-2xl transition-all duration-300 ${activeTab === 'dashboard' ? 'bg-[#006E62] text-white shadow-lg shadow-emerald-100 scale-110' : 'text-gray-400'}`}
        >
          <Home size={24} />
        </button>
        
        <button 
          onClick={() => setActiveTab('presence')}
          className={`p-4 rounded-full transition-all duration-300 -mt-10 border-4 border-white ${activeTab === 'presence' ? 'bg-[#006E62] text-white shadow-xl scale-110' : 'bg-white text-[#006E62] shadow-lg'}`}
        >
          <Fingerprint size={28} />
        </button>

        <button 
          onClick={() => setActiveTab('settings')}
          className={`p-3 rounded-2xl transition-all duration-300 ${activeTab === 'settings' ? 'bg-[#006E62] text-white shadow-lg shadow-emerald-100 scale-110' : 'text-gray-400'}`}
        >
          <UserCircle size={24} />
        </button>
      </nav>
    </div>
  );
};

export default MobileLayout;
