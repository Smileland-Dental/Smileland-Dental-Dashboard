'use client';

import { useState, useEffect } from 'react';
import { onAuthStateChanged, User } from "firebase/auth";
import { auth } from '@/lib/firebase.config';

export default function Page() {
  const [user, setUser] = useState<User | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    // This listener checks for changes in the user's login state.
    const unsubscribe = onAuthStateChanged(auth, (currentUser) => {
      if (currentUser) {
        // User is signed in
        setUser(currentUser);
      } else {
        // User is signed out
        setUser(null);
      }
      setLoading(false);
    });

    // Cleanup the subscription when the component unmounts
    return () => unsubscribe();
  }, []); // The empty array ensures this effect runs only once on component mount

  if (loading) {
    return <div>Loading...</div>;
  }

  return (
    <div className="mx-2">
      {user ? (
        <div>
          <h1 className="mb-4 text-2xl font-extrabold leading-none tracking-tight text-gray-900 md:text-3xl lg:text-4xl dark:text-white">Welcome to the Smileland Dental Employee Dashboard, 
            <br/><strong>{user.displayName}</strong>!
          </h1>
        </div>
      ) : (
        <div>
          <p>No user is logged in. Please return to the login page.</p>
        </div>
      )}
    </div>
  );
}