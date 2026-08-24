// A shell for tests.
//
// Components used to be tested by `vi.mock`-ing the Office bootstrap module,
// which meant every test file re-declared its own idea of what a host looked
// like and none of them exercised the bridge. Registering a real Shell instead
// costs one line per test and catches the case that actually bites: a UI change
// that starts reaching for something no shell provides.

import type { DetectedHost, ToolDefinition } from '@openofficellm/shared'
import type { HostAdapter } from '../host/adapter'
import { registerShell, type Shell, type ToolOutcome } from '../host/bridge'

export interface FakeShellOptions {
  host?: DetectedHost
  surface?: 'office' | 'browser'
  adapter?: HostAdapter | null
  documentKey?: string
  dark?: boolean
  tools?: ToolDefinition[]
  writeTools?: readonly string[]
  exec?: (name: string, argsJson: string) => Promise<ToolOutcome>
}

export function makeFakeShell(opts: FakeShellOptions = {}): Shell {
  const host = opts.host ?? 'word'
  const writes = new Set(opts.writeTools ?? [])
  return {
    surface: opts.surface ?? (host === 'browser' ? 'browser' : 'office'),
    getHost: () => host,
    getAdapter: () => opts.adapter ?? null,
    getDocumentKey: () => opts.documentKey ?? 'doc:test',
    isDark: () => opts.dark ?? false,
    toolCatalog: (_h, allowWrites) =>
      (opts.tools ?? []).filter((t) => allowWrites || !writes.has(t.name)),
    isWriteTool: (name) => writes.has(name),
    executeDocumentTool: (name, argsJson) =>
      opts.exec
        ? opts.exec(name, argsJson)
        : Promise.resolve({ content: `called ${name}`, isError: false }),
  }
}

/** Register a fake shell and return it. Call in `beforeEach`. */
export function installFakeShell(opts: FakeShellOptions = {}): Shell {
  const s = makeFakeShell(opts)
  registerShell(s)
  return s
}
