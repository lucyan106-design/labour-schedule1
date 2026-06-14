import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  build: {
    outDir: "dist",
    // Single chunk for Capacitor (avoids dynamic import issues in WebView)
    rollupOptions: {
      output: {
        manualChunks: undefined
      }
    },
    chunkSizeWarningLimit: 5000
  },
  // Needed so GPS and camera permissions work in Android WebView
  server: {
    headers: {
      "Permissions-Policy": "geolocation=(), camera=()"
    }
  }
});
