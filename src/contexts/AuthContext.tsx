"use client";

import React, { createContext, useContext, useEffect, useState } from 'react';
import { onAuthStateChanged } from 'firebase/auth';
import { doc, onSnapshot } from 'firebase/firestore'; // Swapped getDoc for onSnapshot
import { auth, db } from '@/lib/firebase.config';

interface UserData {
  id: string;
  username: string;
  email: string;
  role: string;
  offices: string[];
  managedEmployeeIds: string[];
}

interface AuthContextType {
  user: UserData | null;
  loading: boolean;
  error: string | null;
}

const AuthContext = createContext<AuthContextType>({
  user: null,
  loading: true,
  error: null,
});

export const AuthProvider = ({ children }: { children: React.ReactNode }) => {
  const [user, setUser] = useState<UserData | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let unsubscribeSnapshot: () => void;

    // Listen for Auth changes
    const unsubscribeAuth = onAuthStateChanged(auth, (firebaseUser) => {
      // 1. Clean up any existing Firestore listener if the user changes or logs out
      if (unsubscribeSnapshot) unsubscribeSnapshot();

      if (firebaseUser) {
        setLoading(true);
        const userDocRef = doc(db, 'users', firebaseUser.uid);

        // 2. Start Live Listener
        unsubscribeSnapshot = onSnapshot(
          userDocRef, 
          (docSnap) => {
            if (docSnap.exists()) {
              const data = docSnap.data();
              setUser({
                id: firebaseUser.uid,
                username: data.username || "Unknown User", // Provide defaults
                email: data.email || firebaseUser.email || "", 
                role: data.role || "Employee",
                offices: data.offices || [],
                managedEmployeeIds: data.managedEmployeeIds || [],
              } as UserData);
            } else {
              setUser(null);
              console.warn("Firestore document does not exist for this user.");
            }
            setLoading(false);
          }, 
          (err) => {
            console.error("Firestore Snapshot Error:", err);
            setError("Failed to sync user data.");
            setLoading(false);
          }
        );
      } else {
        // No user logged in
        setUser(null);
        setLoading(false);
      }
    });

    // Cleanup: Unsubscribe from both Auth and Firestore when component unmounts
    return () => {
      unsubscribeAuth();
      if (unsubscribeSnapshot) unsubscribeSnapshot();
    };
  }, []);

  return (
    <AuthContext.Provider value={{ user, loading, error }}>
      {!loading && children}
    </AuthContext.Provider>
  );
};

export const useAuth = () => useContext(AuthContext);