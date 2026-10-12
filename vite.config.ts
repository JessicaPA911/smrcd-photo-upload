import { defineConfig } from 'vite';

export default defineConfig({
  base: '/smrcd-photo-upload/',
  server: { port: 5173, strictPort: true }
});
