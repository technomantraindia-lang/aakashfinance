import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

const apiProxy = {
  target: "https://app-webappication.vercel.app",
  changeOrigin: true,
  secure: true,
  configure(proxy) {
    proxy.on("proxyReq", (proxyRequest) => {
      proxyRequest.removeHeader("origin");
      proxyRequest.removeHeader("referer");
      proxyRequest.setHeader("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/141.0.0.0 Safari/537.36");
    });
  }
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
