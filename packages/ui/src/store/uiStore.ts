// UI store: theme, sidebar/settings visibility, service-down state, toasts.

import { create } from 'zustand'
import { applyUpdate } from '../api/client'
import { loadPersisted, savePersisted } from './persist'

export type Theme = 'light' | 'dark' | 'office'

export interface Toast {
  id: number
  kind: 'info' | 'success' | 'error' | 'warn'
  message: string
  timeoutMs?: number
}

export interface UiState {
  theme: Theme
  sidebarOpen: boolean
  settingsOpen: boolean
  serviceDown: boolean
  serviceVersion: string | null
  updateAvailable: boolean
  latestVersion: string | null
  updateCountdown: number | null
  toasts: Toast[]
  /** Incremented to signal "open the model selector" — ModelSelector watches
   *  this and opens itself when the value changes. Decouples error recovery
   *  from the selector's internal open state. */
  modelSelectorHint: number
  setTheme: (t: Theme) => void
  toggleSidebar: () => void
  setSidebarOpen: (open: boolean) => void
  openSettings: () => void
  closeSettings: () => void
  setServiceDown: (down: boolean, version?: string | null) => void
  setUpdateAvailable: (available: boolean, version: string | null) => void
  startUpdateCountdown: () => void
  cancelUpdateCountdown: () => void
  tickUpdateCountdown: () => void
  toast: (t: Omit<Toast, 'id'>) => void
  dismissToast: (id: number) => void
  openModelSelectorHint: () => void
}

let countdownInterval: ReturnType<typeof setInterval> | null = null

let toastId = 0

export const useUiStore = create<UiState>((set) => ({
  theme: loadPersisted<Theme>('ui.theme', 'office'),
  sidebarOpen: loadPersisted<boolean>('ui.sidebarOpen', false),
  settingsOpen: false,
  serviceDown: false,
  serviceVersion: null,
  updateAvailable: false,
  latestVersion: null,
  updateCountdown: null,
  toasts: [],
  modelSelectorHint: 0,
  setTheme: (t) => {
    savePersisted('ui.theme', t)
    set({ theme: t })
  },
  toggleSidebar: () =>
    set((s) => {
      const next = !s.sidebarOpen
      savePersisted('ui.sidebarOpen', next)
      return { sidebarOpen: next }
    }),
  setSidebarOpen: (open) => {
    savePersisted('ui.sidebarOpen', open)
    set({ sidebarOpen: open })
  },
  openSettings: () => set({ settingsOpen: true }),
  closeSettings: () => set({ settingsOpen: false }),
  setServiceDown: (down, version = null) => set({ serviceDown: down, serviceVersion: version }),
  setUpdateAvailable: (available, version) =>
    set({ updateAvailable: available, latestVersion: version }),
  startUpdateCountdown: () => {
    if (countdownInterval) clearInterval(countdownInterval)
    set({ updateCountdown: 5 })
    countdownInterval = setInterval(() => {
      useUiStore.getState().tickUpdateCountdown()
    }, 1000)
  },
  cancelUpdateCountdown: () => {
    if (countdownInterval) {
      clearInterval(countdownInterval)
      countdownInterval = null
    }
    set({ updateCountdown: null })
  },
  tickUpdateCountdown: () => {
    const current = useUiStore.getState().updateCountdown
    if (current === null) return
    if (current <= 1) {
      if (countdownInterval) {
        clearInterval(countdownInterval)
        countdownInterval = null
      }
      set({ updateCountdown: null })
      void applyUpdate().catch(() => {})
      return
    }
    set({ updateCountdown: current - 1 })
  },
  toast: (t) => {
    const id = ++toastId
    set((s) => ({ toasts: [...s.toasts, { ...t, id }] }))
    const ttl = t.timeoutMs ?? 5000
    if (ttl > 0) setTimeout(() => dismissById(id), ttl)
  },
  dismissToast: (id) => dismissById(id),
  openModelSelectorHint: () => set((s) => ({ modelSelectorHint: s.modelSelectorHint + 1 })),
}))

function dismissById(id: number) {
  useUiStore.setState((s) => ({ toasts: s.toasts.filter((t) => t.id !== id) }))
}
