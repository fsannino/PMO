import type { Config } from "tailwindcss";

export default {
  content: ["./src/**/*.{ts,tsx}"],
  theme: {
    extend: {
      colors: {
        brand: {
          50: "#eef6ff",
          100: "#d9eaff",
          500: "#2f6feb",
          600: "#1f56c8",
          700: "#1a47a3",
          900: "#102a5e",
        },
      },
      fontFamily: {
        sans: ["ui-sans-serif", "system-ui", "Inter", "Arial", "sans-serif"],
      },
    },
  },
  plugins: [],
} satisfies Config;
