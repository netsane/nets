/** @type {import('tailwindcss').Config} */
// This file configures Tailwind CSS for your project.
module.exports = {
  // CRITICAL: This content array tells Tailwind where your component files are 
  // so it can scan them for class names and generate the necessary CSS.
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}", // Scans all files in src/ with these extensions (standard for React/Vite)
  ],
  theme: {
    // You can customize colors, fonts, spacing, etc., here
    extend: {},
  },
  plugins: [],
}
