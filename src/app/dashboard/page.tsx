/*import Image from 'next/image';
import Link from 'next/link';
import type { Metadata } from "next";

import { getAuth, onAuthStateChanged } from "firebase/auth";
import { auth } from '@/lib/firebase.config'; // Adjust the import path as needed

//const user = auth.currentUser

export const metadata: Metadata = {
  title: "Smileland Dental Dashboard",
}

export default function Page() {
  onAuthStateChanged(auth, (user) => {
    if (user) {
      console.log("User is logged in:", user.displayName);
    } else {
      console.log("No user is logged in.");
    }
  });
  return (
    <div>
        <div>This is the main dashboard.</div>
    </div>
  )
} */
'use client';

import { useState, useEffect } from 'react';
import type { Metadata } from "next";
import { onAuthStateChanged, User } from "firebase/auth";
import { auth } from '@/lib/firebase.config';
import Image from 'next/image';

// Note: Metadata export is for Server Components. 
// For Client Components, you'd typically handle the title differently, 
// but we'll leave it for context.
/*
export const metadata: Metadata = {
  title: "Smileland Dental Dashboard",
}
*/

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