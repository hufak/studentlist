/** @type {import('tailwindcss').Config} */
export default {
  content: ["./index.html", "./src/**/*.{ts,tsx}"],
  theme: {
    extend: {
      colors: {
        bg: "#f4efe9",
        bgAccent: "#f0dcc5",
        surface: "#ffffff",
        text: "#1e1b16",
        muted: "#5a4c3b",
        accent: "#c8553d",
        "accent-strong": "#8f3524",
        border: "#e3d2c0",
      },
      boxShadow: {
        card: "0 24px 60px rgba(54, 35, 16, 0.18)",
        panel: "0 12px 30px rgba(54, 35, 16, 0.08)",
      },
      fontFamily: {
        sans: ["Space Grotesk", "system-ui", "sans-serif"],
        mono: ["IBM Plex Mono", "ui-monospace", "SFMono-Regular", "monospace"],
      },
    },
  },
  plugins: [],
};
