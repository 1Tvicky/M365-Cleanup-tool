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
      keyframes: {
        "indeterminate-bar": {
          "0%": { transform: "translateX(-100%)" },
          "100%": { transform: "translateX(300%)" },
        },
      },
      animation: {
        "indeterminate-bar": "indeterminate-bar 1.1s ease-in-out infinite",
      },
    },
  },
  plugins: [],
};
