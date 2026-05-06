'use client'

import {  Home, 
          File, 
          Settings, 
          ScrollText,
          SquareCheckBig, 
          ChevronRight, 
          BookUser, 
          CalendarX,
          UserCog,
          Sheet
} from "lucide-react" 
import { Button } from "@/components/ui/button"
import Link from 'next/link';

import { auth } from '@/lib/firebase.config';
import { signOut } from 'firebase/auth';
import { useRouter } from 'next/navigation';

import {
  Sidebar,
  SidebarContent,
  SidebarFooter,
  SidebarHeader,
  SidebarGroup,
  SidebarGroupContent,
  SidebarMenu,
  SidebarMenuButton,
  SidebarMenuItem,
  SidebarMenuSub,
  SidebarMenuSubButton,
  SidebarMenuSubItem,
} from "@/components/ui/sidebar"

// IMPORTANT: Import these for collapsibility
import {
  Collapsible,
  CollapsibleContent,
  CollapsibleTrigger,
} from "@/components/ui/collapsible"

import { useAuth } from "@/contexts/AuthContext";

const mainItems = [
  { title: "Overview", url: "/dashboard", icon: Home },
  { title: "Forms", url: "/dashboard/forms", icon: ScrollText },
  { title: "Absence Form", url: "/dashboard/forms/absence-request", icon: CalendarX },
];

const toolItems = [
  { title: "Approval", url: "/dashboard/tools/approval", icon: SquareCheckBig, allowedRoles: ['Manager', 'HR', 'Director'] },
  { title: "HR Absence Table", url: "/dashboard/hr-table", icon: Sheet, allowedRoles: ['HR', 'Director'] },
  { title: "Company Directory", url: "/dashboard/company-directory", icon: BookUser, allowedRoles: ['HR', 'Director'] },
  { title: "User Management", url: "/dashboard/tools/administrator", icon: UserCog, allowedRoles: ['HR', 'Director'] },
];

export function AppSidebar() {
  const { user } = useAuth();
  const router = useRouter();

  const handleSignOut = async () => {
    try {
      await signOut(auth);
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
          <SidebarHeader className="flex items-center justify-between p-4">
            <span className="text-2xl font-black font-mono tracking-tight">
              Smileland Dental
            </span>
          </SidebarHeader>

          <SidebarContent>
            <SidebarGroup>
              <SidebarGroupContent>
                <SidebarMenu>
                  {/* Static Main Items */}
                  {mainItems.map((item) => (
                    <SidebarMenuItem key={item.title}>
                      <SidebarMenuButton asChild>
                        <Link href={item.url}>
                          <item.icon />
                          <span>{item.title}</span>
                        </Link>
                      </SidebarMenuButton>
                    </SidebarMenuItem>
                  ))}

                  {/* COLLAPSIBLE TOOLS SECTION */}
                  <Collapsible asChild className="group/collapsible">
                    <SidebarMenuItem>
                      <CollapsibleTrigger asChild>
                        <SidebarMenuButton tooltip="Tools">
                          <Settings />
                          <span>Tools</span>
                          <ChevronRight className="ml-auto transition-transform duration-200 group-data-[state=open]/collapsible:rotate-90" />
                        </SidebarMenuButton>
                      </CollapsibleTrigger>
                      
                      <CollapsibleContent>
                        <SidebarMenuSub>
                          {toolItems
                            .filter((item) => !item.allowedRoles || item.allowedRoles.includes(user?.role))
                            .map((item) => (
                              <SidebarMenuSubItem key={item.title}>
                                <SidebarMenuSubButton asChild>
                                  <Link href={item.url}>
                                    <item.icon />
                                    <span>{item.title}</span>
                                  </Link>
                                </SidebarMenuSubButton>
                              </SidebarMenuSubItem>
                            ))}
                        </SidebarMenuSub>
                      </CollapsibleContent>
                    </SidebarMenuItem>
                  </Collapsible>

                </SidebarMenu>
              </SidebarGroupContent>
            </SidebarGroup>
          </SidebarContent>

          <SidebarFooter className="p-4 border-t">
            <p className="text-xs">Current User: <strong>{user?.email}</strong></p>
            <Button className="w-full gap-2" onClick={handleSignOut}>
              Sign Out
            </Button>
          </SidebarFooter>
        </Sidebar>
      ) : (
        <div className="p-4 text-center">
          <p className="text-sm text-gray-500">No active session</p>
        </div>
      )}
    </div>
  )
}