/** @type {import('tailwindcss').Config} */
export default {
  content: ["./index.html", "./src/**/*.{ts,tsx}"],
  theme: {
    extend: {
      colors: {
        brand: {
          50: "#eef4ff",
          100: "#dce8ff",
          500: "#3763f4",
          600: "#2c4fd6",
          700: "#233fb0",
        },
      },
    },
  },
  plugins: [],
};
