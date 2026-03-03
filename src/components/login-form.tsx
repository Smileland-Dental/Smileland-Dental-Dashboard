'use client'

import { cn } from "@/lib/utils"
import { Button } from "@/components/ui/button"
import { useEffect } from 'react';
import { useRouter } from 'next/navigation';
import { GoogleAuthProvider, onAuthStateChanged, signInWithPopup } from 'firebase/auth';
import { auth, db } from '@/lib/firebase.config'; // Adjust the import path as needed
import { doc, getDoc, setDoc, serverTimestamp } from "firebase/firestore";

export function LoginForm({
  className,
  ...props
}: React.ComponentProps<"form">) {

  const router = useRouter();

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (user) => {
      if (user) {
        const userDocRef = doc(db, "users", user.uid);
        const docSnap = await getDoc(userDocRef);

        if (!docSnap.exists()) {
          await setDoc(userDocRef, {
            username: user.displayName,
            email: user.email,
            role: "Employee",
            offices: [], // Default to no offices, can be updated later
            createdAt: serverTimestamp() // Good practice to store creation time
          });
          console.log("New user document created in Firestore.");
        }

        user.getIdToken().then((token) => {
          // Set the token in a cookie
          document.cookie = `firebaseAuthToken=${token}; path=/`;
          router.push('/dashboard/forms'); // Redirect to a protected page if logged in
        });
      }
    });

    return () => unsubscribe();
    // Change when router is updated
  }, [router]);

  // Function to handle google sign in
  const handleGoogleSignIn = async () => {
    const provider = new GoogleAuthProvider();
    provider.setCustomParameters({
      prompt: 'select_account' // Ensures the user selects an account
    })
    try {
      await signInWithPopup(auth, provider); 
    } catch (error) {
      console.error('Error during Google sign-in:', error); 
    }
  };

  return (
    <form
      className={cn("flex flex-col gap-6", className)}
      // Prevent form submission since we removed the standard inputs
      onSubmit={(e) => e.preventDefault()}
      {...props}
    >
      <div className="flex flex-col items-center gap-2 text-center">
        <h1 className="text-3xl font-bold">Login to your account</h1>
        <p className="text-muted-foreground text-m text-balance">
          Sign in using your Google account to continue
        </p>
      </div>
      <div className="grid gap-6">
        <Button variant="outline" className="w-full h-12 m-0 flex items-center justify-center gap-3" onClick={handleGoogleSignIn}>
          <svg xmlns="http://www.w3.org/2000/svg" className="size-6" x="0px" y="0px" viewBox="0 0 48 48">
            <path fill="#FFC107" d="M43.611,20.083H42V20H24v8h11.303c-1.649,4.657-6.08,8-11.303,8c-6.627,0-12-5.373-12-12c0-6.627,5.373-12,12-12c3.059,0,5.842,1.154,7.961,3.039l5.657-5.657C34.046,6.053,29.268,4,24,4C12.955,4,4,12.955,4,24c0,11.045,8.955,20,20,20c11.045,0,20-8.955,20-20C44,22.659,43.862,21.35,43.611,20.083z"></path><path fill="#FF3D00" d="M6.306,14.691l6.571,4.819C14.655,15.108,18.961,12,24,12c3.059,0,5.842,1.154,7.961,3.039l5.657-5.657C34.046,6.053,29.268,4,24,4C16.318,4,9.656,8.337,6.306,14.691z"></path><path fill="#4CAF50" d="M24,44c5.166,0,9.86-1.977,13.409-5.192l-6.19-5.238C29.211,35.091,26.715,36,24,36c-5.202,0-9.619-3.317-11.283-7.946l-6.522,5.025C9.505,39.556,16.227,44,24,44z"></path><path fill="#1976D2" d="M43.611,20.083H42V20H24v8h11.303c-0.792,2.237-2.231,4.166-4.087,5.571c0.001-0.001,0.002-0.001,0.003-0.002l6.19,5.238C36.971,39.205,44,34,44,24C44,22.659,43.862,21.35,43.611,20.083z"></path>
          </svg>
          <p className="font-bold text-xl">Login with Google</p>
        </Button>
      </div>
      <div className="flex flex-col items-center gap-2 text-center">
        <p className="text-muted-foreground text-xs">
          Remember to enable pop-ups in the<br/> browser settings
        </p>
      </div>
    </form>
  )
}
