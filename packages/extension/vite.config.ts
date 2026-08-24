/// <reference types="vitest/config" />
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { fileURLToPath } from 'node:url'

// The panel and the service worker. The content script is built separately by
// vite.content.config.ts, because MV3 content scripts are not ES modules and
// cannot be emitted by the same rollup pass.
//
// Filenames are unhashed: manifest.json names background.js literally, and a
// content hash would change it on every build.
export default defineConfig({
  plugins: [react()],
  // Extension pages are loaded from chrome-extension://<id>/, a real origin
  // root, so absolute asset paths are correct here — unlike the task pane,
  // which is served from a path and needs './'.
  base: '/',
  resolve: {
    alias: {
      '@': fileURLToPath(new URL('./src', import.meta.url)),
    },
  },
  build: {
    target: 'es2022',
    sourcemap: false,
    outDir: 'dist',
    emptyOutDir: true,
    chunkSizeWarningLimit: 4000,
    rollupOptions: {
      input: {
        sidepanel: fileURLToPath(new URL('./sidepanel.html', import.meta.url)),
        background: fileURLToPath(new URL('./src/background.ts', import.meta.url)),
      },
      output: {
        entryFileNames: '[name].js',
        chunkFileNames: 'chunks/[name]-[hash].js',
        assetFileNames: 'assets/[name]-[hash][extname]',
      },
    },
  },
  test: {
    environment: 'jsdom',
    globals: true,
    include: ['src/**/*.{test,spec}.{ts,tsx}'],
  },
})
