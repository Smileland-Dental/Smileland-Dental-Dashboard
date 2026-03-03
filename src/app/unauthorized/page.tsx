'use client';

import Link from 'next/link';
import { ShieldAlert, ArrowLeft, Lock } from 'lucide-react';

export default function UnauthorizedPage() {
  return (
    <div className="min-h-[80vh] flex items-center justify-center px-4">
      <div className="max-w-md w-full text-center space-y-6 p-8 bg-white rounded-2xl shadow-xl border border-gray-100">
        <div className="relative flex justify-center">
          <div className="p-4 bg-red-50 rounded-full">
            <ShieldAlert className="h-12 w-12 text-red-600" />
          </div>
          <Lock className="absolute -bottom-2 -right-2 h-6 w-6 text-gray-400 bg-white rounded-full p-1 shadow-sm" />
        </div>

        <div className="space-y-2">
          <h1 className="text-3xl font-bold text-gray-900">Restricted Access</h1>
          <p className="text-gray-500">
            Sorry, you don't have the administrative permissions required to view this directory.
          </p>
        </div>

        <div className="pt-4 space-y-3">
          <Link 
            href="/"
            className="flex items-center justify-center gap-2 w-full py-3 bg-indigo-600 text-white font-semibold rounded-xl hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-100"
          >
            <ArrowLeft className="h-4 w-4" />
            Back to Dashboard
          </Link>
          
          <p className="text-xs text-gray-400">
            If you believe this is an error, please contact your System Administrator.
          </p>
        </div>
      </div>
    </div>
  );
}