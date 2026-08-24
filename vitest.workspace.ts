import { defineWorkspace } from 'vitest/config'
import react from '@vitejs/plugin-react'
import { fileURLToPath } from 'node:url'

// Vitest workspace — each package owns its test environment. The browser-side
// packages need jsdom + jest-dom matchers + the React plugin; host and shared
// run under node. This file takes precedence over vitest.config.ts (we keep
// only one).
//
// `ui` and the shells share one project rather than getting one each: they
// need identical environments, and a single project lets a shell's tests
// import the UI package's test helpers without a cross-project resolve.

export default defineWorkspace([
  {
    plugins: [react()],
    test: {
      name: 'browser',
      environment: 'jsdom',
      globals: true,
      setupFiles: ['./packages/ui/src/test/setup.ts'],
      css: true,
      include: [
        'packages/ui/src/**/*.test.{ts,tsx}',
        'packages/addin/src/**/*.test.{ts,tsx}',
        'packages/extension/src/**/*.test.{ts,tsx}',
      ],
      coverage: {
        provider: 'v8',
        reporter: ['text', 'html'],
        include: ['packages/ui/src/**', 'packages/addin/src/**', 'packages/extension/src/**'],
        exclude: ['**/*.test.{ts,tsx}', '**/types.ts'],
      },
    },
    resolve: {
      alias: {
        '@': fileURLToPath(new URL('./packages/ui/src', import.meta.url)),
      },
    },
  },
  {
    test: {
      name: 'node',
      environment: 'node',
      globals: true,
      include: ['packages/host/**/*.test.ts', 'packages/shared/**/*.test.ts'],
      coverage: {
        provider: 'v8',
        reporter: ['text', 'html'],
        include: ['packages/host/src/**', 'packages/shared/src/**'],
        exclude: [
          'packages/host/src/**/*.test.ts',
          'packages/shared/src/**/*.test.ts',
          'packages/shared/src/**/types.ts',
        ],
      },
    },
  },
])
