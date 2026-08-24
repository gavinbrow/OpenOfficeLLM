// MCP store: server list, live status, and per-tool consent.
//
// Consent is never held here as the source of truth — the host enforces it on
// every call from config.json. This store only mirrors it for the settings UI,
// so a stale pane cannot grant a tool anything.

import { create } from 'zustand'
import type { McpServerConfig, McpServerInfo } from '@openofficellm/shared'
import { getMcpServers, putMcpServer, deleteMcpServer, setMcpConsent } from '../api/client'

export interface McpState {
  servers: McpServerInfo[]
  loading: boolean
  error: string | null
  load: (refresh?: boolean) => Promise<void>
  save: (server: McpServerConfig) => Promise<boolean>
  remove: (id: string) => Promise<void>
  setToolEnabled: (serverId: string, tool: string, enabled: boolean) => Promise<void>
}

export const useMcpStore = create<McpState>((set) => ({
  servers: [],
  loading: false,
  error: null,

  load: async (refresh = false) => {
    set({ loading: true, error: null })
    try {
      const { servers } = await getMcpServers(refresh)
      set({ servers, loading: false })
    } catch (e) {
      set({ loading: false, error: (e as Error).message ?? 'Could not load MCP servers.' })
    }
  },

  save: async (server) => {
    set({ error: null })
    try {
      const { servers } = await putMcpServer(server)
      set({ servers })
      return true
    } catch (e) {
      set({ error: (e as Error).message ?? 'Could not save the server.' })
      return false
    }
  },

  remove: async (id) => {
    try {
      const { servers } = await deleteMcpServer(id)
      set({ servers })
    } catch (e) {
      set({ error: (e as Error).message ?? 'Could not remove the server.' })
    }
  },

  setToolEnabled: async (serverId, tool, enabled) => {
    try {
      const { servers } = await setMcpConsent(serverId, tool, enabled)
      set({ servers })
    } catch (e) {
      set({ error: (e as Error).message ?? 'Could not change consent.' })
    }
  },
}))
