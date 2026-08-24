import { defineConfig } from 'tsup'

export default defineConfig({
  entry: ['src/index.ts'],
  // CJS for the SEA binary: Node's SEA embedder runs the blob through CJS,
  // and ESM entry points fail with "Cannot use import statement outside a
  // module" inside the SEA loader. The dev workflow (tsx, vitest) imports
  // the source .ts directly, so this only affects the built output.
  format: ['cjs'],
  platform: 'node',
  target: 'es2022',
  outDir: 'dist',
  clean: true,
  sourcemap: true,
  // Bundle EVERYTHING into the output. Node SEA's require() only supports
  // built-in modules — any `require('hono')` etc. in the bundle fails with
  // ERR_UNKNOWN_BUILTIN_MODULE. The only thing that must stay external is
  // win-dpapi (native addon, loaded via process.dlopen at runtime).
  noExternal: [/.*/],
  external: ['win-dpapi'],
  banner: { js: '#!/usr/bin/env node' },
})
