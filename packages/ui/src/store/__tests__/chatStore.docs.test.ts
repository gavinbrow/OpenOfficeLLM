// Each document gets its own chat.
//
// The task pane's localStorage is keyed to the add-in's origin, not to the
// document, so without `adoptDocument` opening a second document in Word
// reopens the first document's conversation.

import { describe, it, expect, beforeEach, vi } from 'vitest'
import type { Conversation } from '../chatStore'

const KEY = 'openofficellm:chat.conversations'
const ACTIVE = 'openofficellm:chat.activeId'

function conv(id: string, docKey: string | undefined, updatedAt: number): Conversation {
  return {
    id,
    title: id,
    messages: [{ id: `${id}_m`, role: 'user', content: 'hi' }],
    createdAt: updatedAt,
    updatedAt,
    ...(docKey ? { docKey } : {}),
  }
}

async function storeWith(conversations: Conversation[], activeId: string | null) {
  localStorage.setItem(KEY, JSON.stringify(conversations))
  localStorage.setItem(ACTIVE, JSON.stringify(activeId))
  // The store hydrates from localStorage at module load, so it has to be
  // re-imported after the seed is written. Resetting modules also clears the
  // registered shell, hence the re-install: `newChat` tags conversations with
  // whatever `getDocumentKey` returns, and that comes from the shell.
  vi.resetModules()
  const { installFakeShell } = await import('../../test/fakeShell')
  installFakeShell({ host: 'word', documentKey: 'doc:/current.docx' })
  const { useChatStore } = await import('../chatStore')
  return useChatStore
}

describe('adoptDocument', () => {
  beforeEach(() => {
    localStorage.clear()
  })

  it('starts fresh when the restored chat belongs to a different document', async () => {
    const store = await storeWith([conv('c1', 'doc:/a.docx', 100)], 'c1')
    store.getState().adoptDocument('doc:/b.docx')
    expect(store.getState().activeId).toBeNull()
    // The other document's chat is not destroyed, just not shown.
    expect(store.getState().conversations).toHaveLength(1)
  })

  it('keeps the restored chat when it belongs to this document', async () => {
    const store = await storeWith([conv('c1', 'doc:/a.docx', 100)], 'c1')
    store.getState().adoptDocument('doc:/a.docx')
    expect(store.getState().activeId).toBe('c1')
  })

  it('reopens the most recent chat for a document returned to', async () => {
    const store = await storeWith(
      [
        conv('older', 'doc:/a.docx', 100),
        conv('newer', 'doc:/a.docx', 300),
        conv('other', 'doc:/b.docx', 400),
      ],
      'other',
    )
    store.getState().adoptDocument('doc:/a.docx')
    expect(store.getState().activeId).toBe('newer')
  })

  it('never auto-adopts a chat saved before chats were scoped per document', async () => {
    const store = await storeWith([conv('legacy', undefined, 100)], 'legacy')
    store.getState().adoptDocument('doc:/a.docx')
    expect(store.getState().activeId).toBeNull()
    // Still reachable from the sidebar.
    expect(store.getState().conversations.map((c) => c.id)).toEqual(['legacy'])
  })

  it('persists the choice so a reload does not resurrect the other document', async () => {
    const store = await storeWith([conv('c1', 'doc:/a.docx', 100)], 'c1')
    store.getState().adoptDocument('doc:/b.docx')
    expect(JSON.parse(localStorage.getItem(ACTIVE) ?? 'undefined')).toBeNull()
  })

  it('tags a newly created chat with the current document', async () => {
    const store = await storeWith([], null)
    store.getState().newChat()
    const created = store.getState().conversations[0]
    expect(created.docKey).toBe('doc:/current.docx')
    // And is therefore not adopted by a different document.
    store.getState().adoptDocument('doc:/somewhere-else.docx')
    expect(store.getState().activeId).toBeNull()
  })
})
