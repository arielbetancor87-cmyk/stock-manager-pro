import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { resolve } from "path";

export default defineConfig({
  plugins: [react()],
  root: "public",
  build: {
    outDir: "../dist",
    emptyOutDir: true,
    commonjsOptions: { transformMixedEsModules: true }
  },
  optimizeDeps: {
    include: ["react", "react-dom", "xlsx"]
  }
});
