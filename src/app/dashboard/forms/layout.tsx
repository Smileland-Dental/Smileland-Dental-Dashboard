import { Metadata } from "next"

// Metadata for the /dashboard/forms page
export const metadata: Metadata = {
  title: "Smileland Dental Forms",
  description: "The Employee Dashboard for Smileland Dental"
}


// Render /dashboard/forms/page.tsx
export default function Layout({ children }: { children: React.ReactNode }) {
  return (
    children
  )
}