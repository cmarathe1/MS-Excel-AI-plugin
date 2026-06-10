import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { resolve } from 'node:path';

export default defineConfig({
  plugins: [react()],
  build: {
    outDir: 'dist',
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, 'taskpane.html'),
        functions: resolve(__dirname, 'functions.html'),
      },
      output: {
        // The manifest references the functions script by a fixed name.
        entryFileNames: (chunk) =>
          chunk.name === 'functions' ? 'assets/functions.js' : 'assets/[name]-[hash].js',
      },
    },
  },
  server: {
    port: 3000,
  },
});
