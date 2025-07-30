import Image from 'next/image';
import Link from 'next/link';
import type { Metadata } from "next";

import { auth } from '@/lib/firebase.config';

import { getAuth, onAuthStateChanged, User } from "firebase/auth";
export const metadata: Metadata = {
  title: "Smileland Dental Dashboard",
}


export default function Page() {
  return (
    <div>
        <div>This is the main dashboard.</div>
    </div>
  )
}
