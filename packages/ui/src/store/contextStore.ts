// Context store: what document context is included with the next turn.

import { create } from 'zustand'
import type { ContextScope, DocumentContext } from '@openofficellm/shared'

export interface ContextItem {
  id: string
  label: string
  scope: ContextScope
  tokenEstimate: number
  /** Snapshot of the context payload at add time (for resending). */
  context: DocumentContext
}

export interface ContextState {
  items: ContextItem[]
  add: (item: ContextItem) => void
  remove: (id: string) => void
  clear: () => void
  totalTokens: () => number
  toDocumentContext: () => DocumentContext | undefined
}

export const useContextStore = create<ContextState>((set, get) => ({
  items: [],
  add: (item) =>
    set((s) => ({
      items: s.items.some((i) => i.id === item.id)
        ? s.items.map((i) => (i.id === item.id ? item : i))
        : [...s.items, item],
    })),
  remove: (id) => set((s) => ({ items: s.items.filter((i) => i.id !== id) })),
  clear: () => set({ items: [] }),
  totalTokens: () => get().items.reduce((sum, i) => sum + (i.tokenEstimate || 0), 0),
  toDocumentContext: () => {
    const items = get().items
    if (items.length === 0) return undefined
    // Merge into a single DocumentContext; the host treats `text` as the
    // concatenation and uses `scope` only as a hint.
    const last = items[items.length - 1]
    const text = items.map((i) => i.context.text).join('\n\n---\n\n')
    return {
      host: last.context.host,
      scope: last.scope,
      text,
      tokenEstimate: get().totalTokens(),
    }
  },
}))
