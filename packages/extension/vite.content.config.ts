import { defineConfig } from 'vite'
import { fileURLToPath } from 'node:url'

// The content script, built on its own.
//
// MV3 content scripts are classic scripts, not modules: no `import`, no
// top-level await, no code splitting. Library mode with a single IIFE output
// is the only shape Chrome will accept, and `emptyOutDir: false` keeps it from
// deleting the main build that ran a moment earlier.
export default defineConfig({
  build: {
    target: 'es2022',
    outDir: 'dist',
    emptyOutDir: false,
    sourcemap: false,
    lib: {
      entry: fileURLToPath(new URL('./src/content/page.ts', import.meta.url)),
      name: 'OpenOfficeLLMPageReader',
      formats: ['iife'],
      // Must match CONTENT_SCRIPT in browser/pageAdapter.ts.
      fileName: () => 'content.js',
    },
  },
})
