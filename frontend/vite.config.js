// vite.config.js
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import svgr from 'vite-plugin-svgr';
import { fileURLToPath } from 'url';
import path from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export default defineConfig({
  plugins: [react(), svgr()],
  resolve: {
    alias: {
      '@': path.resolve(__dirname, 'src'),
      // or: '@': fileURLToPath(new URL('./src', import.meta.url)),
    },
  },
  server: { host: true, port: 3000, strictPort: true },
  preview: { port: 3000, strictPort: true },
});
