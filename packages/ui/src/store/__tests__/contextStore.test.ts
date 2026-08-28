// Context store: which document context and attachments are attached to the
// next turn, and the split between them. Live-document items go on
// ChatRequest.context; attachment items go on ChatRequest.attachments and
// the host holds the bytes. The store's removeAttachment also fires a
// best-effort host delete and revokes any object URL, so both are mocked.

import { describe, it, expect, beforeEach, vi } from 'vitest'
import { useContextStore, type ContextItem } from '../contextStore'

// deleteAttachment hits the network; mock the api client barrel so the store
// never makes a real call. Only deleteAttachment is used by contextStore, but
// the mock replaces the whole module so other re-exports stay inert. The real
// deleteAttachment returns Promise<void>, so the default mock resolves to
// undefined — contextStore calls `.catch()` on the result, so each mock must
// return a promise.
vi.mock('../../api/client', () => ({
  deleteAttachment: vi.fn().mockResolvedValue(undefined),
}))

import { deleteAttachment } from '../../api/client'

function docItem(id: string, text = 'body'): ContextItem {
  return {
    id,
    label: id,
    scope: 'document',
    tokenEstimate: 10,
    context: { host: 'word', scope: 'document', text },
    kind: 'document',
  }
}

function textItem(id: string, label = 'notes.txt'): ContextItem {
  return {
    id,
    label,
    scope: 'none',
    tokenEstimate: 5,
    context: { host: 'none', scope: 'none', text: 'extracted', fileName: label, isAttachment: true },
    kind: 'text-attachment',
    attachmentId: id,
    mimeType: 'text/plain',
  }
}

function imageItem(id: string, label = 'photo.png', thumbUrl?: string): ContextItem {
  return {
    id,
    label,
    scope: 'none',
    tokenEstimate: 0,
    context: { host: 'none', scope: 'none', text: '', fileName: label, isAttachment: true },
    kind: 'image-attachment',
    attachmentId: id,
    mimeType: 'image/png',
    thumbUrl,
  }
}

function resetStore() {
  useContextStore.setState({ items: [] })
}

describe('contextStore attachments', () => {
  beforeEach(() => {
    resetStore()
    vi.clearAllMocks()
  })

  it('addAttachment adds an item with kind text-attachment', () => {
    useContextStore.getState().addAttachment(textItem('a1'))
    expect(useContextStore.getState().items).toHaveLength(1)
    expect(useContextStore.getState().items[0].kind).toBe('text-attachment')
  })

  it('addAttachment deduplicates by id (replaces rather than appends)', () => {
    useContextStore.getState().addAttachment(textItem('a1', 'first.txt'))
    useContextStore.getState().addAttachment(textItem('a1', 'second.txt'))
    expect(useContextStore.getState().items).toHaveLength(1)
    expect(useContextStore.getState().items[0].label).toBe('second.txt')
  })

  it('toAttachments returns AttachmentRef[] for attachment items only', () => {
    useContextStore.getState().add(docItem('d1', 'doc body'))
    useContextStore.getState().addAttachment(textItem('a1', 'notes.txt'))
    useContextStore.getState().addAttachment(imageItem('a2', 'photo.png'))
    const refs = useContextStore.getState().toAttachments()
    expect(refs).toHaveLength(2)
    expect(refs.map((r) => r.id).sort()).toEqual(['a1', 'a2'])
    const img = refs.find((r) => r.id === 'a2')!
    expect(img.kind).toBe('image')
    expect(img.mimeType).toBe('image/png')
    const txt = refs.find((r) => r.id === 'a1')!
    expect(txt.kind).toBe('text')
    expect(txt.fileName).toBe('notes.txt')
  })

  it('toDocumentContext excludes attachment items', () => {
    useContextStore.getState().add(docItem('d1', 'doc body'))
    useContextStore.getState().addAttachment(textItem('a1', 'notes.txt'))
    const ctx = useContextStore.getState().toDocumentContext()
    expect(ctx).toBeDefined()
    expect(ctx!.text).toBe('doc body')
  })

  it('toDocumentContext is undefined when only attachments are present', () => {
    useContextStore.getState().addAttachment(textItem('a1'))
    expect(useContextStore.getState().toDocumentContext()).toBeUndefined()
  })

  it('toAttachments is empty when only document items are present', () => {
    useContextStore.getState().add(docItem('d1'))
    expect(useContextStore.getState().toAttachments()).toEqual([])
  })

  it('removeAttachment removes the item and calls deleteAttachment on the host', async () => {
    useContextStore.getState().addAttachment(textItem('a1'))
    await useContextStore.getState().removeAttachment('a1')
    expect(useContextStore.getState().items).toHaveLength(0)
    expect(deleteAttachment).toHaveBeenCalledWith('a1')
  })

  it('removeAttachment revokes the thumbUrl object URL', async () => {
    const revoke = vi.fn()
    // jsdom may not define revokeObjectURL; provide a stub for this test.
    const orig = URL.revokeObjectURL
    ;(URL as { revokeObjectURL: unknown }).revokeObjectURL = revoke
    try {
      useContextStore.getState().addAttachment(imageItem('a1', 'photo.png', 'blob:uuid-1'))
      await useContextStore.getState().removeAttachment('a1')
      expect(revoke).toHaveBeenCalledWith('blob:uuid-1')
    } finally {
      ;(URL as { revokeObjectURL: unknown }).revokeObjectURL = orig
    }
  })

  it('removeAttachment still completes when deleteAttachment rejects', async () => {
    vi.mocked(deleteAttachment).mockRejectedValueOnce(new Error('network'))
    useContextStore.getState().addAttachment(textItem('a1'))
    await expect(useContextStore.getState().removeAttachment('a1')).resolves.toBeUndefined()
    expect(useContextStore.getState().items).toHaveLength(0)
  })

  it('removeAttachment does not fire a host delete for an unknown id', async () => {
    // An unknown id is not in the store, so there is no item to remove and no
    // attachment blob to delete on the host. Firing a DELETE anyway would be a
    // wasted round-trip against an id that is not an attachment — and a
    // document chip routed here by mistake would have produced a bogus
    // request. The early return guards both.
    await expect(useContextStore.getState().removeAttachment('nope')).resolves.toBeUndefined()
    expect(deleteAttachment).not.toHaveBeenCalled()
  })

  it('removeAttachment does not fire a host delete for a document item', async () => {
    useContextStore.getState().add(docItem('d1', 'doc body'))
    await expect(useContextStore.getState().removeAttachment('d1')).resolves.toBeUndefined()
    // The document item is still present — removeAttachment is not its path.
    expect(useContextStore.getState().items).toHaveLength(1)
    expect(deleteAttachment).not.toHaveBeenCalled()
  })
})