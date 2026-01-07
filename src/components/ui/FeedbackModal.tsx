// components/ui/FeedbackModal.tsx
import React from 'react';

interface FeedbackModalProps {
  isOpen: boolean;
  type: 'success' | 'error';
  message: string;
  onClose: () => void;
}

export default function FeedbackModal({ isOpen, type, message, onClose }: FeedbackModalProps) {
  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 bg-black/60 flex justify-center items-center z-[60]">
      <div className="bg-white p-8 rounded-xl shadow-2xl max-w-sm w-full text-center">
        <div className={`text-5xl mb-4 ${type === 'success' ? 'text-green-500' : 'text-red-500'}`}>
          {type === 'success' ? '✓' : '✕'}
        </div>
        <h3 className="text-xl font-bold mb-2">
          {type === 'success' ? 'Success!' : 'Something went wrong'}
        </h3>
        <p className="text-gray-600 mb-6">{message}</p>
        <button
          onClick={onClose}
          className={`w-full py-2 rounded-lg font-semibold text-white transition-colors 
            ${type === 'success' ? 'bg-green-500 hover:bg-green-600' : 'bg-red-500 hover:bg-red-600'}`}
        >
          {type === 'success' ? 'Continue' : 'Try Again'}
        </button>
      </div>
    </div>
  );
}