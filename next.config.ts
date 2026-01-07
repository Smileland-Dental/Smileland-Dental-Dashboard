import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  /* config options here */
  devIndicators: false,
  /*rewrites: async () => {
   return [
     {
        source: '/absence-request',
        destination: '/forms/absence-request/form.html',
     },
    ]
  },*/
  images: {
    remotePatterns: [
      {
        protocol: 'https',
        hostname: 'firebasestorage.googleapis.com',
        port: '',
        pathname: '/v0/b/smileland-dental-dashboard.firebasestorage.app/o/**'
      }
    ]
  },
}

export default nextConfig;
