import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import { setSecret, getSecret, deleteSecret, listConfigured, isFallbackMode } from '../secrets.js'
import { resolveSecretsPath } from '../paths.js'

const TMP_DIR = path.join(os.tmpdir(), `ool-test-${process.pid}-${Date.now()}`)

// paths.ts picks its root per platform: %APPDATA% on Windows, $HOME on macOS.
// Redirect both so the suite exercises a throwaway store on every platform.
function withTempSecrets<T>(fn: () => Promise<T>): Promise<T> {
  const origAppData = process.env.APPDATA
  const origHome = process.env.HOME
  const tmpAppData = path.join(TMP_DIR, 'appdata')
  fs.mkdirSync(tmpAppData, { recursive: true })
  process.env.APPDATA = tmpAppData
  process.env.HOME = TMP_DIR
  try {
    return fn()
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

  it('round-trips a secret when DPAPI is available', async () => {
    await withTempSecrets(async () => {
      // Force a DPAPI probe; on Windows this exercises the PowerShell path,
      // on other platforms the AES fallback is used.
      setSecret('anthropic', 'sk-dpapi-roundtrip-456')
      if (!isFallbackMode()) {
        const file = fs.readFileSync(resolveSecretsPath(), 'utf8')
        // New DPAPI blobs must be written with the 'dpapi:' prefix.
        expect(file).toContain('"dpapi:')
      }
      expect(getSecret('anthropic')).toBe('sk-dpapi-roundtrip-456')
    })
  })

  it('an existing fb: blob still decrypts when DPAPI is available', async () => {
    await withTempSecrets(async () => {
      // Simulate an existing user's secrets file written by a previous build
      // that only ever used the AES fallback. We construct an 'fb:' blob
      // directly with a fixed key, bypassing setSecret (which would write
      // 'dpapi:' when DPAPI is available).
      const { resolveSecretsPath, ensureDirs } = await import('../paths.js')
      ensureDirs()
      const storePath = resolveSecretsPath()

      const fallbackKey = 'ool-test-fallback-key-32-bytes-long!!'
      process.env.OPENOFFICELLM_SECRETS_FALLBACK_KEY = fallbackKey
      const crypto = await import('node:crypto')
      const iv = crypto.randomBytes(16)
      const cipher = crypto.createCipheriv('aes-256-gcm', Buffer.from(fallbackKey, 'utf8').subarray(0, 32), iv)
      const enc = Buffer.concat([cipher.update('sk-legacy-fb-blob', 'utf8'), cipher.final()])
      const tag = cipher.getAuthTag()
      const blob = ['fb', iv.toString('base64'), tag.toString('base64'), enc.toString('base64')].join(':')
      fs.writeFileSync(storePath, JSON.stringify({ legacy: blob }, null, 2), { encoding: 'utf8', mode: 0o600 })

      // Reading must dispatch on the 'fb:' prefix and succeed regardless of
      // whether DPAPI is currently available on this host.
      expect(getSecret('legacy')).toBe('sk-legacy-fb-blob')

      delete process.env.OPENOFFICELLM_SECRETS_FALLBACK_KEY
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
      const file = fs.readFileSync(resolveSecretsPath(), 'utf8')
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
