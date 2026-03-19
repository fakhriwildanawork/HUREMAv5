
import React from 'react';

export const CardSkeleton: React.FC = () => (
  <div className="bg-white border border-gray-100 p-4 rounded-md shadow-sm animate-pulse">
    <div className="h-4 bg-gray-200 rounded w-3/4 mb-3"></div>
    <div className="h-3 bg-gray-100 rounded w-full mb-2"></div>
    <div className="h-3 bg-gray-100 rounded w-5/6"></div>
  </div>
);

export const FormSkeleton: React.FC = () => (
  <div className="space-y-4 animate-pulse">
    <div className="grid grid-cols-2 gap-4">
      <div className="h-10 bg-gray-100 rounded"></div>
      <div className="h-10 bg-gray-100 rounded"></div>
    </div>
    <div className="h-24 bg-gray-100 rounded"></div>
  </div>
);
