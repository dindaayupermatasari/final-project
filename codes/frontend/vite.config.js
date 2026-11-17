import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Selama development, permintaan /api diteruskan ke backend FastAPI (uvicorn api.app:app --port 8000).
export default defineConfig({
  plugins: [react()],
  server: { proxy: { "/api": "http://localhost:8000" } },
});
