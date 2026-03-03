// tailwind.config.js
/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    // This is the correct, standard path for Next.js 
    // to find classes in pages, components, and app directories.
    "./app/**/*.{js,ts,jsx,tsx,mdx}",
    "./pages/**/*.{js,ts,jsx,tsx,mdx}",
    "./components/**/*.{js,ts,jsx,tsx,mdx}",
    // If your project files are in a "src" folder, add this line too.
    "./src/**/*.{js,ts,jsx,tsx,mdx}",
  ],
  theme: {
    extend: {},
  },
  plugins: [],
};
