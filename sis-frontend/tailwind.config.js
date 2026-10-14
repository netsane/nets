/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    "./index.html",
    // This line is crucial for React/Vite projects
    "./src/**/*.{js,ts,jsx,tsx}", 
  ],
  theme: {
    extend: {
      fontFamily: {
        // Ensuring Inter font is used as requested
        sans: ['Inter', 'sans-serif'],
      },
    },
  },
  plugins: [],
}