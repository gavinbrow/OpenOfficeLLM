// Type-level checks for the attachment protocol additions.
//
// The shared package's other tests (settings.test.ts) use plain runtime
// assertions; this file mirrors that style. Each `it` constructs a value of
// the relevant type and asserts on its shape. If the types regress, tsc fails
// to compile this file and the test run breaks — so these are compile-time
// assertions wearing runtime clothes.

import { describe, it, expect } from 'vitest'
import type {
  AttachmentRef,
  ChatMessage,
  ChatRequest,
  ContentBlock,
  DocumentContext,
} from '../index'

describe('ContentBlock', () => {
  it('has a text variant carrying a string', () => {
    const block: ContentBlock = { type: 'text', text: 'hi' }
    expect(block.type).toBe('text')
    expect(block.text).toBe('hi')
  })

  it('has an image variant carrying a mimeType and base64 data', () => {
    const block: ContentBlock = { type: 'image', mimeType: 'image/png', data: 'AAA' }
    expect(block.type).toBe('image')
    expect(block.mimeType).toBe('image/png')
    expect(block.data).toBe('AAA')
  })

  it('discriminates text from image on the type tag', () => {
    const blocks: ContentBlock[] = [
      { type: 'text', text: 'a' },
      { type: 'image', mimeType: 'image/png', data: 'AAA' },
    ]
    const texts = blocks.filter((b): b is Extract<ContentBlock, { type: 'text' }> => b.type === 'text')
    expect(texts).toHaveLength(1)
    expect(texts[0].text).toBe('a')
  })
})

describe('AttachmentRef', () => {
  it('carries id, fileName, kind, mimeType and tokenEstimate', () => {
    const ref: AttachmentRef = {
      id: 'att-1',
      fileName: 'notes.txt',
      kind: 'text',
      mimeType: 'text/plain',
      tokenEstimate: 12,
    }
    expect(ref.id).toBe('att-1')
    expect(ref.fileName).toBe('notes.txt')
    expect(ref.kind).toBe('text')
    expect(ref.mimeType).toBe('text/plain')
    expect(ref.tokenEstimate).toBe(12)
  })

  it('accepts kind "image" for vision attachments', () => {
    const ref: AttachmentRef = {
      id: 'att-2',
      fileName: 'photo.png',
      kind: 'image',
      mimeType: 'image/png',
      tokenEstimate: 0,
    }
    expect(ref.kind).toBe('image')
  })
})

describe('ChatMessage.content', () => {
  it('accepts a plain string', () => {
    const msg: ChatMessage = { role: 'user', content: 'hello' }
    expect(msg.content).toBe('hello')
  })

  it('accepts an array of ContentBlock for multimodal turns', () => {
    const msg: ChatMessage = {
      role: 'user',
      content: [
        { type: 'image', mimeType: 'image/png', data: 'AAA' },
        { type: 'text', text: 'what is this?' },
      ],
    }
    expect(Array.isArray(msg.content)).toBe(true)
    expect((msg.content as ContentBlock[])).toHaveLength(2)
  })
})

describe('ChatRequest.attachments', () => {
  it('is optional — a request with no attachments compiles', () => {
    const req: ChatRequest = {
      messages: [{ role: 'user', content: 'hi' }],
      model: 'ollama/test',
      mode: 'propose',
    }
    expect(req.attachments).toBeUndefined()
  })

  it('accepts an AttachmentRef[] when present', () => {
    const req: ChatRequest = {
      messages: [{ role: 'user', content: 'hi' }],
      model: 'ollama/test',
      mode: 'propose',
      attachments: [
        { id: 'a1', fileName: 'a.txt', kind: 'text', mimeType: 'text/plain', tokenEstimate: 1 },
        { id: 'a2', fileName: 'b.png', kind: 'image', mimeType: 'image/png', tokenEstimate: 0 },
      ],
    }
    expect(req.attachments).toHaveLength(2)
  })
})

describe('DocumentContext attachment fields', () => {
  it('compiles without fileName and isAttachment (back-compat with live-document context)', () => {
    const ctx: DocumentContext = { host: 'word', scope: 'document', text: 'body' }
    expect(ctx.fileName).toBeUndefined()
    expect(ctx.isAttachment).toBeUndefined()
  })

  it('accepts fileName and isAttachment for attachment-derived context', () => {
    const ctx: DocumentContext = {
      host: 'none',
      scope: 'none',
      text: 'extracted',
      fileName: 'upload.txt',
      isAttachment: true,
    }
    expect(ctx.fileName).toBe('upload.txt')
    expect(ctx.isAttachment).toBe(true)
  })
})