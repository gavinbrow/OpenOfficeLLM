// The browser implementation of the UI's Shell contract.
//
// Compare packages/addin/src/office/shell.ts: same shape, different host. That
// symmetry is the whole point of the bridge — everything below this file is
// browser-specific, everything above it is shared.

import type { Shell } from '@openofficellm/ui'
import { getHost, getAdapter, getDocumentKey, isDark } from './bootstrap'
import { toolCatalog, isWriteTool, executeDocumentTool } from './tools'

export const browserShell: Shell = {
  surface: 'browser',
  getHost,
  getAdapter,
  getDocumentKey,
  isDark,
  toolCatalog,
  isWriteTool,
  executeDocumentTool,
  // No `describeFormatting`: this host has no formatting vocabulary, and
  // nothing can stage a proposal on it anyway. The optional member exists for
  // exactly this case.
}
