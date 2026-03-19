
import React from 'react';
import ReactDOM from 'react-dom';

const LoadingSpinner: React.FC<{ message?: string }> = ({ message = 'Menyimpan...' }) => {
  return ReactDOM.createPortal(
    <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-black bg-opacity-40 backdrop-blur-sm">
      <div className="bg-white p-6 rounded-lg shadow-xl flex flex-col items-center gap-4">
        <div className="w-12 h-12 border-4 border-[#006E62] border-t-transparent rounded-full animate-spin"></div>
        <p className="text-[#006E62] font-semibold">{message}</p>
      </div>
    </div>,
    document.body
  );
};

export default LoadingSpinner;
