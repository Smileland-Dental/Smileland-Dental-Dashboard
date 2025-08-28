import { SidebarProvider, SidebarTrigger } from "@/components/ui/sidebar"
import { AppSidebar } from "@/components/app-sidebar"
import { Metadata } from "next"

// Metadata for the /dashboard page
export const metadata: Metadata = {
  title: "Smileland Dental Dashboard",
  description: "The Employee Dashboard for Smileland Dental"
}

// This is to adjust the main layout of the dashboard pages. The 
// SidebarProvider wraps the sidebar and main content area.
// If you need to adjust how the sidebar is positioned, adjust here.
export default function Layout({ children }: { children: React.ReactNode }) {
  return (
    <SidebarProvider>
      <AppSidebar />
      <main className='w-full'>
        <SidebarTrigger/>
        {children}
      </main>
    </SidebarProvider>
  )
}