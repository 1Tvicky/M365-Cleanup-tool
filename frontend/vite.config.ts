import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

export default defineConfig({
  plugins: [react()],
  server: {
    // Without this, Vite binds only to IPv6 loopback (::1) on this machine — a browser resolving
    // "localhost" to 127.0.0.1 (IPv4) then gets ERR_CONNECTION_REFUSED even though the dev server
    // is actually running. host: true binds every interface, IPv4 and IPv6 alike.
    host: true,
    proxy: {
      "/api": "http://localhost:4000",
    },
  },
});
