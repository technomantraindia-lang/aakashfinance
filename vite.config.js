import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

const apiProxy = {
  target: "https://erp.aakashfinance.com",
  changeOrigin: true,
  secure: true
};

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5175,
    proxy: { "/api": apiProxy },
    allowedHosts: [".onrender.com"]
  },
  preview: {
    host: "0.0.0.0",
    port: 4174,
    proxy: { "/api": apiProxy },
    allowedHosts: [".onrender.com"]
  }
});
