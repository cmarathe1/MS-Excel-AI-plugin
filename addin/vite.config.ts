import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { resolve } from 'node:path';

export default defineConfig({
  plugins: [react()],
  build: {
    outDir: 'dist',
    rollupOptions: {
      // Single shared-runtime page: the taskpane hosts chat UI and registers
      // the custom functions (see manifest Script -> Shared.Url).
      input: {
        taskpane: resolve(__dirname, 'taskpane.html'),
      },
    },
  },
  server: {
    port: 3000,
  },
});
