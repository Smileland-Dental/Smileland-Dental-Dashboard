import type { Metadata } from "next";
//import { Geist, Geist_Mono } from "next/font/google";
import "./globals.css";

/*const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});
*/


export const metadata: Metadata = {
  title: "Smileland Dental Login",
  description: "The Employee Dashboard for Smileland Dental",
  icons: {
    icon: [
      { url: 'favicon.ico' },
      { url: '/favicon.ico', type: 'image/png' },
    ],
    apple: [
      { url: '/icon.png'}
    ],
  }
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      {/*<link rel="icon" href="/favicon.ico" sizes="any" />*/}
      <body
        //className={`${geistSans.variable} ${geistMono.variable} antialiased`}
      >
        {children}
      </body>
    </html>
  );
}
