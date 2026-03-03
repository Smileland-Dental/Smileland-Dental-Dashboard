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

import { useAuth } from "@/contexts/AuthContext";

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
  {
    title: "Absence Form",
    url: "/dashboard/forms/absence-request",
    icon: File,
  },
  {
    title: "Approval",
    url: "/dashboard/tools/approval",
    icon: File,
    allowedRoles: ['Manager', 'HR', 'Director'], // Only show to Managers, HR, and Directors
  },
  {
    title: "HR Table",
    url: "/dashboard/hr-table",
    icon: File,
    allowedRoles: ['HR', 'Director'],
  },
  {
    title: "Company Directory",
    url: "/dashboard/company-directory",
    icon: File,
    allowedRoles: ['HR', 'Director'],
  },
  /*{
    title: "Settings",
    url: "#",
    icon: Settings,
  },*/
]

export function AppSidebar() {
  const { user, loading } = useAuth();

  const router = useRouter();
  
  // Allow the user to log out based on a button on the sidebar
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
      <Sidebar variant="floating">
        <SidebarContent>
          <SidebarGroup>
            <SidebarGroupLabel>Smileland Dental</SidebarGroupLabel>
            <SidebarGroupContent>
            <SidebarMenu>
              {items
                .filter((item) => {
                  // If the item has no role restrictions, show it
                  if (!item.allowedRoles) return true;
                  
                  // Otherwise, only show if user's role is in the allowed list
                  return user && item.allowedRoles.includes(user.role);
                })
                .map((item) => (
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
          <p>No User</p>
        </div>
      )}
    </div>
  )
}
