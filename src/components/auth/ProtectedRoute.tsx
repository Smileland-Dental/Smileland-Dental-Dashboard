'use client';

import { useAuth } from '@/contexts/AuthContext';
import { useRouter } from 'next/navigation';
import { useEffect } from 'react';

interface ProtectedRouteProps {
  children: React.ReactNode;
  allowedRoles: string[];
}

export default function ProtectedRoute({ children, allowedRoles }: ProtectedRouteProps) {
  const { user, loading } = useAuth();
  const router = useRouter();

  useEffect(() => {
    // If auth is finished and there's no user or the role doesn't match, redirec
    if (!loading) {
      if (!user || !allowedRoles.includes(user.role)) {
        // Change this from '/' to '/unauthorized'
        router.push('/unauthorized'); 
      }
    }
  }, [user, loading, allowedRoles, router]);

  // Show a loading state while checking permissions
  if (loading) {
    return (
      <div className="flex h-screen items-center justify-center">
        <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600"></div>
      </div>
    );
  }

  // If user exists and role is allowed, render the page
  if (user && allowedRoles.includes(user.role)) {
    return <>{children}</>;
  }

  // Return null while the useEffect handles the redirect
  return null;
}