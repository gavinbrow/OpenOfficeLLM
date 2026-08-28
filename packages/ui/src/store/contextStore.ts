// Context store: what document context is included with the next turn.
//
// Items are either live-document context (kind 'document' or undefined — the
// older shape, before attachments existed) or user-attached files
// ('text-attachment' / 'image-attachment'). Live-document items are merged into
// a single DocumentContext and sent on ChatRequest.context; attachment items
// are sent as AttachmentRef[] on ChatRequest.attachments and the host holds the
// bytes. The split is why toDocumentContext() filters out attachments and
// toAttachments() filters out documents.

import { create } from 'zustand'
import type {
  AttachmentKind,
  AttachmentRef,
  ContextScope,
  DocumentContext,
} from '@openofficellm/shared'
import { deleteAttachment } from '../api/client'

export interface ContextItem {
  id: string
  label: string
  scope: ContextScope
  tokenEstimate: number
  /** Snapshot of the context payload at add time (for resending). */
  context: DocumentContext
  /** 'document' for live-document scope chips, 'text-attachment' or
   *  'image-attachment' for dropped files. Absent means document — the older
   *  shape, before attachments were introduced. */
  kind?: 'document' | 'text-attachment' | 'image-attachment'
  /** Host-assigned attachment id, present only for attachment items. */
  attachmentId?: string
  /** Object URL for image thumbnail display (created on the client). Revoked
   *  in removeAttachment(). */
  thumbUrl?: string
  /** MIME type of the source file, for attachment items. Used by
   *  toAttachments() so the host gets the original content type even though
   *  DocumentContext.text is always a string. */
  mimeType?: string
}

export interface ContextState {
  items: ContextItem[]
  add: (item: ContextItem) => void
  remove: (id: string) => void
  clear: () => void
  totalTokens: () => number
  toDocumentContext: () => DocumentContext | undefined
  /** Named entry point for attachment items. Same dedup behaviour as add(),
   *  but a distinct name so call sites read clearly
   *  (`contextStore.addAttachment(...)` vs `contextStore.add(...)`). */
  addAttachment: (item: ContextItem) => void
  /** Remove an attachment item: drop it from the store, revoke any object URL,
   *  and best-effort delete it on the host. The host delete is fire-and-forget
   *  so a network failure doesn't leave the chip stranded in the UI. */
  removeAttachment: (id: string) => Promise<void>
  /** Attachment refs for ChatRequest.attachments — attachment items only;
   *  live-document items are excluded (they go on `context`). */
  toAttachments: () => AttachmentRef[]
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
    // Only live-document items belong in the DocumentContext. Attachments
    // travel on ChatRequest.attachments, so including them here would double-
    // count their text once the host also folds the attachment text in.
    const items = get().items.filter((i) => i.kind === 'document' || i.kind === undefined)
    if (items.length === 0) return undefined
    // Merge into a single DocumentContext; the host treats `text` as the
    // concatenation and uses `scope` only as a hint.
    const last = items[items.length - 1]
    const text = items.map((i) => i.context.text).join('\n\n---\n\n')
    return {
      host: last.context.host,
      scope: last.scope,
      text,
      tokenEstimate: items.reduce((sum, i) => sum + (i.tokenEstimate || 0), 0),
    }
  },
  addAttachment: (item) =>
    set((s) => ({
      // Same dedup as add(): replace an existing item with the same id, else
      // append. Kept separate from add() purely for call-site readability.
      items: s.items.some((i) => i.id === item.id)
        ? s.items.map((i) => (i.id === item.id ? item : i))
        : [...s.items, item],
    })),
  removeAttachment: async (id) => {
    const item = get().items.find((i) => i.id === id)
    // Only attachment items should be removed through this path; a document
    // chip routed here by mistake would fire a bogus host DELETE against an
    // id that is not an attachment blob, which is at best a wasted round-trip
    // and at worst confuses a future host that reuses the id. Document items
    // are removed through `remove` instead.
    if (!item || (item.kind !== 'text-attachment' && item.kind !== 'image-attachment')) return
    // Revoke the object URL so the browser can free the underlying blob. A
    // failed revoke is harmless (the URL may already be gone) — swallow it.
    if (item.thumbUrl) {
      try {
        URL.revokeObjectURL(item.thumbUrl)
      } catch {
        /* ignore — best-effort cleanup */
      }
    }
    set((s) => ({ items: s.items.filter((i) => i.id !== id) }))
    // Best-effort host delete; don't block the UI on failure. A dropped network
    // here just leaves an orphaned blob on the host, which is preferable to a
    // chip the user thought they removed snapping back into the store.
    void deleteAttachment(id).catch(() => {})
  },
  toAttachments: () => {
    return get()
      .items.filter((i) => i.kind === 'text-attachment' || i.kind === 'image-attachment')
      .map((i) => ({
        id: i.attachmentId!,
        fileName: i.label,
        kind: (i.kind === 'image-attachment' ? 'image' : 'text') as AttachmentKind,
        // Fall back to a generic binary type if the caller never set one; the
        // host treats text attachments by their extracted text regardless.
        mimeType: i.mimeType ?? 'application/octet-stream',
        tokenEstimate: i.tokenEstimate,
      }))
  },
}))
