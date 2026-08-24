// Edits the model has proposed but that have not touched the document (P4.18).
//
// Only populated in `propose` mode. Direct and agentic apply as they go, so a
// non-empty list here always means "waiting on the user".

import { create } from 'zustand'
import type { Edit } from '@openofficellm/shared'
import { getAdapter } from '../host/bridge'
import type { ApplyResult, HostAdapter } from '../host/adapter'

/** Apply a batch, or say plainly that this host cannot be written to.
 *
 *  `applyEdits` is optional on the adapter contract — the browser adapter
 *  omits it, because a web page is not ours to rewrite. Nothing should ever
 *  stage a proposal on such a host, but if something does, the user gets a
 *  refusal rather than a TypeError swallowed into "could not apply". */
async function applyOrRefuse(adapter: HostAdapter, edits: Edit[]): Promise<ApplyResult> {
  if (!adapter.applyEdits) {
    return { ok: false, summary: 'This host is read-only — nothing was changed.' }
  }
  return adapter.applyEdits(edits)
}

export interface Proposal {
  id: string
  /** One-line description shown in the review card. */
  description: string
  edit: Edit
  /** The conversation turn that produced it, so stale proposals from an
   *  earlier turn can be cleared when a new one starts. */
  turnId: string
}

export interface ProposalState {
  proposals: Proposal[]
  applying: boolean
  lastResult: { ok: boolean; message: string } | null

  add: (proposal: Omit<Proposal, 'id'>) => void
  discard: (id: string) => void
  discardAll: () => void
  applyAll: () => Promise<void>
  applyOne: (id: string) => Promise<void>
  clearResult: () => void
}

let counter = 0
function genId(): string {
  counter += 1
  return `prop_${Date.now().toString(36)}_${counter}`
}

export const useProposalStore = create<ProposalState>((set, get) => ({
  proposals: [],
  applying: false,
  lastResult: null,

  add: (proposal) => set((s) => ({ proposals: [...s.proposals, { ...proposal, id: genId() }] })),

  discard: (id) => set((s) => ({ proposals: s.proposals.filter((p) => p.id !== id) })),

  discardAll: () => set({ proposals: [], lastResult: null }),

  applyAll: async () => {
    const { proposals } = get()
    if (proposals.length === 0 || get().applying) return
    const adapter = getAdapter()
    if (!adapter) {
      set({ lastResult: { ok: false, message: 'No document is open.' } })
      return
    }
    set({ applying: true })
    try {
      // One applyEdits call for the whole batch: Word and Excel each group a
      // single run into one undo record, so accepting five proposals still
      // reverts with one Ctrl+Z.
      const result = await applyOrRefuse(
        adapter,
        proposals.map((p) => p.edit),
      )
      set({
        applying: false,
        proposals: result.ok ? [] : get().proposals,
        lastResult: { ok: result.ok, message: result.summary },
      })
    } catch (e) {
      set({
        applying: false,
        lastResult: { ok: false, message: (e as Error).message ?? 'Could not apply the changes.' },
      })
    }
  },

  applyOne: async (id) => {
    const proposal = get().proposals.find((p) => p.id === id)
    if (!proposal || get().applying) return
    const adapter = getAdapter()
    if (!adapter) {
      set({ lastResult: { ok: false, message: 'No document is open.' } })
      return
    }
    set({ applying: true })
    try {
      const result = await applyOrRefuse(adapter, [proposal.edit])
      set((s) => ({
        applying: false,
        proposals: result.ok ? s.proposals.filter((p) => p.id !== id) : s.proposals,
        lastResult: { ok: result.ok, message: result.summary },
      }))
    } catch (e) {
      set({
        applying: false,
        lastResult: { ok: false, message: (e as Error).message ?? 'Could not apply the change.' },
      })
    }
  },

  clearResult: () => set({ lastResult: null }),
}))
