import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { loadEnv } from "vite";

function createApiProxy(target) {
  return {
    target,
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
}

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), "");
  const apiProxy = createApiProxy(env.VITE_API_PROXY_TARGET || "http://localhost:4000");

  return {
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
  };
});
