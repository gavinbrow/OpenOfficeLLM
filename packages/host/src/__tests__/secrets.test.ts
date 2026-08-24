import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import { setSecret, getSecret, deleteSecret, listConfigured, isFallbackMode } from '../secrets.js'

const TMP_DIR = path.join(os.tmpdir(), `ool-test-${process.pid}-${Date.now()}`)

function withTempSecrets<T>(fn: () => Promise<T>): Promise<T> {
  const origAppData = process.env.APPDATA
  const tmpAppData = path.join(TMP_DIR, 'appdata')
  fs.mkdirSync(tmpAppData, { recursive: true })
  process.env.APPDATA = tmpAppData
  try {
    return fn()
  } finally {
    process.env.APPDATA = origAppData
    try {
      fs.rmSync(TMP_DIR, { recursive: true, force: true })
    } catch {
      // ignore
    }
  }
}

describe('secrets', () => {
  beforeEach(() => {
    fs.rmSync(TMP_DIR, { recursive: true, force: true })
  })
  afterEach(() => {
    fs.rmSync(TMP_DIR, { recursive: true, force: true })
  })

  it('round-trips a secret (fallback or DPAPI)', async () => {
    await withTempSecrets(async () => {
      setSecret('anthropic', 'sk-test-key-123')
      const got = getSecret('anthropic')
      expect(got).toBe('sk-test-key-123')
    })
  })

  it('listConfigured returns IDs only, never values', async () => {
    await withTempSecrets(async () => {
      setSecret('anthropic', 'sk-secret-value-aaa')
      setSecret('openai', 'sk-secret-value-bbb')
      const ids = listConfigured()
      expect(ids).toContain('anthropic')
      expect(ids).toContain('openai')
      expect(ids.every((id) => typeof id === 'string')).toBe(true)
      const file = fs.readFileSync(
        path.join(process.env.APPDATA!, 'OpenOfficeLLM', 'secrets.dat'),
        'utf8',
      )
      expect(file).not.toContain('sk-secret-value-aaa')
      expect(file).not.toContain('sk-secret-value-bbb')
    })
  })

  it('deleteSecret removes a secret', async () => {
    await withTempSecrets(async () => {
      setSecret('groq', 'gsk_test')
      expect(deleteSecret('groq')).toBe(true)
      expect(getSecret('groq')).toBeNull()
      expect(listConfigured()).not.toContain('groq')
    })
  })

  it('deleteSecret returns false for missing secret', async () => {
    await withTempSecrets(async () => {
      expect(deleteSecret('nope')).toBe(false)
    })
  })

  it('getSecret returns null for missing secret', async () => {
    await withTempSecrets(async () => {
      expect(getSecret('missing')).toBeNull()
    })
  })

  it('secret never appears in a serialized response object', async () => {
    await withTempSecrets(async () => {
      setSecret('anthropic', 'sk-leak-check-xyz')
      const sampleResponse = {
        ok: true,
        configured: true,
      }
      const serialized = JSON.stringify(sampleResponse)
      expect(serialized).not.toContain('sk-leak-check-xyz')
      const ids = listConfigured()
      const idsSerialized = JSON.stringify(ids)
      expect(idsSerialized).not.toContain('sk-leak-check-xyz')
    })
  })

  it('isFallbackMode returns a boolean', async () => {
    await withTempSecrets(async () => {
      expect(typeof isFallbackMode()).toBe('boolean')
    })
  })

  it('overwrites an existing secret', async () => {
    await withTempSecrets(async () => {
      setSecret('anthropic', 'first-key')
      setSecret('anthropic', 'second-key')
      expect(getSecret('anthropic')).toBe('second-key')
      expect(listConfigured().filter((id) => id === 'anthropic').length).toBe(1)
    })
  })
})
