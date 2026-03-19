import React from 'react';
import { Menu, UserCircle } from 'lucide-react';
import { AuthUser } from '../../types';
import { googleDriveService } from '../../services/googleDriveService';

interface HeaderProps {
  activeTab: string;
  onMenuClick: () => void;
  user: AuthUser;
}

const Header: React.FC<HeaderProps> = ({ activeTab, onMenuClick, user }) => {
  return (
    <header className="h-16 border-b border-gray-100 flex items-center justify-between px-4 md:px-8 bg-white sticky top-0 z-30">
      <div className="flex items-center gap-4">
        <button className="md:hidden p-2 hover:bg-gray-100 rounded" onClick={onMenuClick}>
          <Menu size={20} />
        </button>
      </div>
      <div className="flex items-center gap-4">
        <div className="flex flex-col items-end hidden sm:flex">
          <span className="text-sm font-semibold text-gray-700">{user.full_name}</span>
        </div>
        <div className="w-10 h-10 rounded-full bg-gray-100 border border-gray-200 overflow-hidden flex items-center justify-center text-gray-400 font-bold shadow-sm">
          {user.photo_google_id ? (
            <img 
              src={googleDriveService.getFileUrl(user.photo_google_id)} 
              alt={user.full_name} 
              className="w-full h-full object-cover"
              referrerPolicy="no-referrer"
            />
          ) : (
            <UserCircle size={24} />
          )}
        </div>
      </div>
    </header>
  );
};

export default Header;