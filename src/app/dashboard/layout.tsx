import { SidebarProvider, SidebarTrigger } from "@/components/ui/sidebar"
import { AppSidebar } from "@/components/app-sidebar"
import { Metadata } from "next"

export const metadata: Metadata = {
  title: "Smileland Dental Dashboard",
  description: "The Employee Dashboard for Smileland Dental"
}

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