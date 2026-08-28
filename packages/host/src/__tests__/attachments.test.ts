// On-disk attachment store lifecycle.
//
// store.ts keeps an in-memory Map keyed by id and writes files under
// ATTACHMENTS_DIR, which paths.ts derives from %APPDATA% (Windows) at module
// load. Redirecting %APPDATA% to a temp dir and re-importing the module under
// test — the same trick secrets.test.ts uses — gives every test a throwaway
// store on a real filesystem.

import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'

const TMP_DIR = path.join(os.tmpdir(), `ool-attach-test-${process.pid}-${Date.now()}`)

async function withTempStore<T>(fn: (mod: typeof import('../attachments/store.js')) => Promise<T>): Promise<T> {
  const origAppData = process.env.APPDATA
  const origHome = process.env.HOME
  const tmpAppData = path.join(TMP_DIR, 'appdata')
  fs.mkdirSync(tmpAppData, { recursive: true })
  process.env.APPDATA = tmpAppData
  process.env.HOME = TMP_DIR
  // ATTACHMENTS_DIR is computed at module load, so the module has to be
  // re-imported after the env is redirected. Resetting also clears the
  // module-level `store` Map so tests don't see each other's attachments.
  vi.resetModules()
  try {
    const mod = await import('../attachments/store.js')
    return await fn(mod)
  } finally {
    process.env.APPDATA = origAppData
    process.env.HOME = origHome
    try {
      fs.rmSync(TMP_DIR, { recursive: true, force: true })
    } catch {
      // ignore
    }
  }
}

describe('attachments store', () => {
  beforeEach(() => {
    fs.rmSync(TMP_DIR, { recursive: true, force: true })
  })
  afterEach(() => {
    fs.rmSync(TMP_DIR, { recursive: true, force: true })
  })

  it('saveAttachment writes a file and returns metadata with an id', async () => {
    await withTempStore(async ({ saveAttachment, getMeta }) => {
      const result = saveAttachment({
        name: 'notes.txt',
        type: 'text/plain',
        buffer: Buffer.from('hello world', 'utf-8'),
      })
      expect(result.id).toBeTruthy()
      expect(result.fileName).toBe('notes.txt')
      expect(result.mimeType).toBe('text/plain')
      expect(result.kind).toBe('text')
      const meta = getMeta(result.id)
      expect(meta).not.toBeNull()
      expect(meta!.id).toBe(result.id)
      expect(meta!.fileName).toBe('notes.txt')
      expect(fs.existsSync(meta!.path)).toBe(true)
    })
  })

  it('getMeta returns null for an unknown id', async () => {
    await withTempStore(async ({ getMeta }) => {
      expect(getMeta('does-not-exist')).toBeNull()
    })
  })

  it('getBytes returns the buffer and mime type for a saved attachment', async () => {
    await withTempStore(async ({ saveAttachment, getBytes }) => {
      const result = saveAttachment({
        name: 'data.txt',
        type: 'text/plain',
        buffer: Buffer.from('payload', 'utf-8'),
      })
      const bytes = getBytes(result.id)
      expect(bytes).not.toBeNull()
      expect(bytes!.buffer.toString('utf-8')).toBe('payload')
      expect(bytes!.mimeType).toBe('text/plain')
    })
  })

  it('getBytes returns null for an unknown id', async () => {
    await withTempStore(async ({ getBytes }) => {
      expect(getBytes('missing')).toBeNull()
    })
  })

  it('classifies image/* uploads as kind image', async () => {
    await withTempStore(async ({ saveAttachment, getMeta }) => {
      const result = saveAttachment({
        name: 'photo.png',
        type: 'image/png',
        buffer: Buffer.from([0x89, 0x50, 0x4e, 0x47]),
      })
      expect(result.kind).toBe('image')
      expect(getMeta(result.id)!.kind).toBe('image')
    })
  })

  it('falls back to application/octet-stream when the browser sent no mime type', async () => {
    await withTempStore(async ({ saveAttachment, getMeta }) => {
      const result = saveAttachment({
        name: 'README.md',
        type: '',
        buffer: Buffer.from('# hi', 'utf-8'),
      })
      expect(result.mimeType).toBe('application/octet-stream')
      expect(getMeta(result.id)!.mimeType).toBe('application/octet-stream')
    })
  })

  it('deleteAttachment removes the file and meta', async () => {
    await withTempStore(async ({ saveAttachment, getMeta, getBytes, deleteAttachment }) => {
      const result = saveAttachment({
        name: 'temp.txt',
        type: 'text/plain',
        buffer: Buffer.from('bye', 'utf-8'),
      })
      const storedPath = getMeta(result.id)!.path
      const storedDir = path.dirname(storedPath)
      expect(fs.existsSync(storedPath)).toBe(true)
      expect(deleteAttachment(result.id)).toBe(true)
      expect(getMeta(result.id)).toBeNull()
      expect(getBytes(result.id)).toBeNull()
      // The per-id directory is gone too.
      expect(fs.existsSync(storedDir)).toBe(false)
    })
  })

  it('deleteAttachment returns false for a missing id', async () => {
    await withTempStore(async ({ deleteAttachment }) => {
      expect(deleteAttachment('nope')).toBe(false)
    })
  })

  it('cleanup wipes the directory and clears the in-memory map', async () => {
    await withTempStore(async ({ saveAttachment, getMeta, cleanup }) => {
      const a = saveAttachment({ name: 'a.txt', type: 'text/plain', buffer: Buffer.from('a') })
      const b = saveAttachment({ name: 'b.txt', type: 'text/plain', buffer: Buffer.from('b') })
      expect(getMeta(a.id)).not.toBeNull()
      expect(getMeta(b.id)).not.toBeNull()
      cleanup()
      expect(getMeta(a.id)).toBeNull()
      expect(getMeta(b.id)).toBeNull()
    })
  })

  it('assigns a fresh id per save (no dedup by content or name)', async () => {
    await withTempStore(async ({ saveAttachment }) => {
      const a = saveAttachment({ name: 'same.txt', type: 'text/plain', buffer: Buffer.from('x') })
      const b = saveAttachment({ name: 'same.txt', type: 'text/plain', buffer: Buffer.from('x') })
      expect(a.id).not.toBe(b.id)
    })
  })

  it('preserves only the basename in the stored path and fileName', async () => {
    await withTempStore(async ({ saveAttachment, getMeta }) => {
      const result = saveAttachment({
        name: 'sub/dir/escape.txt',
        type: 'text/plain',
        buffer: Buffer.from('x'),
      })
      // path.basename strips the directory — a name with a path separator
      // cannot write outside its per-id dir.
      expect(result.fileName).toBe('escape.txt')
      expect(getMeta(result.id)!.fileName).toBe('escape.txt')
      expect(getMeta(result.id)!.path).not.toContain('sub')
    })
  })
})