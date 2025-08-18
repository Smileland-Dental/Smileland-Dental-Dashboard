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
};

export default nextConfig;
