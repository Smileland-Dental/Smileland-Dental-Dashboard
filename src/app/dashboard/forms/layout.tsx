import { Metadata } from "next"

export const metadata: Metadata = {
  title: "Smileland Dental Forms",
  description: "The Employee Dashboard for Smileland Dental"
}

export default function Layout({ children }: { children: React.ReactNode }) {
  return (
    children
  )
}