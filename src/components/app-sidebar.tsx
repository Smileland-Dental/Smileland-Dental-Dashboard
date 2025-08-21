'use client'

import { Home, File, Settings } from "lucide-react"
import { Button } from "@/components/ui/button"
import Link from 'next/link';

import { auth } from '@/lib/firebase.config'; // Adjust the import path
import { signOut } from 'firebase/auth';
import { useRouter } from 'next/navigation';

import { useState, useEffect } from 'react';
import { onAuthStateChanged, User } from "firebase/auth";

import {
  Sidebar,
  SidebarContent,
  SidebarFooter,
  SidebarGroup,
  SidebarGroupContent,
  SidebarGroupLabel,
  SidebarMenu,
  SidebarMenuButton,
  SidebarMenuItem,
  SidebarTrigger,
} from "@/components/ui/sidebar"

// Menu items.
const items = [
  {
    title: "Overview",
    url: "/dashboard",
    icon: Home,
  },
  {
    title: "Forms",
    url: "/dashboard/forms",
    icon: File,
  },
  /*{
    title: "Settings",
    url: "#",
    icon: Settings,
  },*/
]

export function AppSidebar() {
  const router = useRouter();
  const [user, setUser] = useState<User | null>(null);
  
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
    });

      // Cleanup the subscription when the component unmounts
      return () => unsubscribe();
    }, []); 

  const handleSignOut = async () => {
    try {
      await signOut(auth);
      // Remove the cookie
      document.cookie = 'firebaseAuthToken=; path=/; expires=Thu, 01 Jan 1970 00:00:00 GMT';
      router.push('/');
    } catch (error) {
      console.error('Error signing out:', error);
    }
  };

  return (
    <div>
      {user ? (
      <Sidebar>
        <SidebarContent>
          <SidebarGroup>
            <SidebarGroupLabel>Smileland Dental</SidebarGroupLabel>
            <SidebarGroupContent>
              <SidebarMenu>
                {items.map((item) => (
                  <SidebarMenuItem key={item.title}>
                    <SidebarMenuButton asChild>
                      <Link href={item.url}>
                        <item.icon />
                        <span>{item.title}</span>
                      </Link>
                    </SidebarMenuButton>
                  </SidebarMenuItem>
                ))}
              </SidebarMenu>
            </SidebarGroupContent>
          </SidebarGroup>
        </SidebarContent>
        <SidebarFooter>
          <p className="text-xs">Current User: <strong>{user.email}</strong></p>
          <Button onClick={handleSignOut}>
                <p>Sign Out</p>
          </Button>
        </SidebarFooter>
      </Sidebar>
      ) : (
        <div>
          <p>No user is logged in.</p>
        </div>
      )}
    </div>
  )
}