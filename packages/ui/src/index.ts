// The public surface of the UI package — everything a shell needs and nothing
// it does not. Shells import from here rather than reaching into `src/`, so
// internal reorganisation stays internal.

export { App } from './App'
export { ErrorBoundary } from './components/ErrorBoundary'

export {
  registerShell,
  shell,
  getSurface,
  getHost,
  getAdapter,
  getDocumentKey,
  isDark,
  settingsHost,
  __resetShellForTest,
  type Shell,
  type ToolExecContext,
  type ToolOutcome,
} from './host/bridge'

export {
  SCOPES_FOR_HOST,
  estimateTokens,
  type HostAdapter,
  type SnapshotPayload,
  type ApplyResult,
  type SearchHit,
} from './host/adapter'

export { configureApi, apiUrl, getAuthToken } from './api/client'
export { useChatStore } from './store/chatStore'
export { useSettingsStore } from './store/settingsStore'
export { useUiStore } from './store/uiStore'
export { useProposalStore } from './store/proposalStore'
export { applyTheme } from './theme'
