// This file tells your bundler (like Vite) which PostCSS plugins to use.
module.exports = {
  plugins: {
    // 1. Tailwind CSS must be the first plugin
    'tailwindcss': {},
    // 2. Autoprefixer adds vendor prefixes (like -webkit- and -moz-) to your CSS
    // to ensure compatibility across different browsers.
    'autoprefixer': {},
  },
}
