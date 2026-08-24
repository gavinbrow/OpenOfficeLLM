// The Office implementation of the UI's Shell contract.
//
// Everything Office-specific that the chat UI needs is assembled here and
// handed over in one object, so `packages/ui` never imports Office.js and the
// Chrome extension can supply its own equivalent without touching a single
// component.

import type { Shell } from '@openofficellm/ui'
import { getHost, getDocumentKey, getTheme } from './bootstrap'
import { getAdapter } from './factory'
import { toolCatalog, isWriteTool, executeDocumentTool } from './tools'
import { describeFormatting } from './wordFormat'

export const officeShell: Shell = {
  surface: 'office',
  getHost,
  getAdapter,
  getDocumentKey,
  isDark: () => getTheme().isDark,
  toolCatalog,
  isWriteTool,
  executeDocumentTool,
  describeFormatting,
}
