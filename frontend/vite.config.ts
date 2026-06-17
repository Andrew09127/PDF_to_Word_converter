import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Dev-сервер фронта на 127.0.0.1:5173. Запросы к API проксируем на локальный
// FastAPI (main.py, 127.0.0.1:8000), поэтому в коде используем относительные пути
// — и в деве, и при раздаче из FastAPI статикой всё работает без правок.
export default defineConfig({
  plugins: [react()],
  server: {
    host: "127.0.0.1",
    port: 5173,
    proxy: {
      "/convert": "http://127.0.0.1:8000",
      "/health": "http://127.0.0.1:8000",
    },
  },
  build: {
    outDir: "dist",
  },
});
