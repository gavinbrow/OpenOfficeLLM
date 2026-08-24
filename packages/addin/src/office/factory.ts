// Returns the adapter for the detected host (P4.2).

import type { HostAdapter } from '@openofficellm/ui'
import { getHost } from './bootstrap'
import { WordAdapter } from './word'
import { ExcelAdapter } from './excel'

let cached: HostAdapter | null = null

/**
 * The adapter for the current host, or null outside Office.
 *
 * Cached because the adapters are stateless — each operation opens its own
 * Word.run/Excel.run — so a single instance is safe and avoids re-creating one
 * per tool call in an agentic loop.
 */
export function getAdapter(): HostAdapter | null {
  const host = getHost()
  if (host === 'none') return null
  if (cached && cached.host === host) return cached
  cached = host === 'word' ? new WordAdapter() : new ExcelAdapter()
  return cached
}

/** Test seam: force a specific adapter. */
export function __setAdapterForTest(adapter: HostAdapter | null): void {
  cached = adapter
}
