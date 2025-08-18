'use client'

import { Home, File, Settings } from "lucide-react"
import { Button } from "@/components/ui/button"
import Link from 'next/link';

import { auth } from '@/lib/firebase.config'; // Adjust the import path
import { signOut } from 'firebase/auth';
import { useRouter } from 'next/navigation';

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

import { getAuth, onAuthStateChanged, User } from "firebase/auth";

onAuthStateChanged(auth, (user: User | null) => {
  if (user) {
    // User is signed in
    const userName = user.displayName;
    if (userName) {
      console.log("Logged-in user's name:", userName);
      // You can now use userName in your application
    } else {
      console.log("Logged-in user does not have a display name.");
    }
  } else {
    // User is signed out
    console.log("No user is currently logged in.");
  }
});

export function AppSidebar() {
  const router = useRouter();

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
    <Sidebar>
      <SidebarContent>
        <SidebarGroup>
          <SidebarGroupLabel>Application</SidebarGroupLabel>
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
        <Button onClick={handleSignOut}>
              <p>Sign Out</p>
        </Button>
      </SidebarFooter>
    </Sidebar>
  )
}