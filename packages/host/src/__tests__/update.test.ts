import { describe, it, expect, vi, beforeEach, afterEach, afterAll } from 'vitest'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { checkForUpdate, getCachedUpdate, skipVersion, applyUpdate } from '../update.js'
import { resolveUpdateStatePath } from '../paths.js'
import { logger } from '../logging.js'
import { HOST_VERSION } from '../paths.js'

const TMP_DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'ool-update-'))
const HOME = path.join(TMP_DIR, 'home')

const ORIG_HOME = process.env.HOME
const ORIG_APPDATA = process.env.APPDATA
const ORIG_HOMEDIR = os.homedir

function fakeRelease(tagName: string): Response {
  return {
    ok: true,
    status: 200,
    json: async () => ({
      tag_name: tagName,
      published_at: '2026-01-01T00:00:00Z',
      body: 'Release notes here',
      assets: [
        {
          name: 'OpenOfficeLLM-0.2.0-macOS.zip',
          browser_download_url: 'https://example.com/mac.zip',
        },
        {
          name: 'OpenOfficeLLM-0.2.0-win.zip',
          browser_download_url: 'https://example.com/win.zip',
        },
      ],
    }),
  } as Response
}

function readStateFile(): Record<string, unknown> | null {
  try {
    return JSON.parse(fs.readFileSync(resolveUpdateStatePath(), 'utf8'))
  } catch {
    return null
  }
}

beforeEach(() => {
  fs.rmSync(TMP_DIR, { recursive: true, force: true })
  fs.mkdirSync(HOME, { recursive: true })
  process.env.HOME = HOME
  if (process.platform === 'win32') {
    process.env.APPDATA = path.join(TMP_DIR, 'appdata')
    fs.mkdirSync(process.env.APPDATA, { recursive: true })
  }
  vi.spyOn(os, 'homedir').mockReturnValue(HOME)
  vi.spyOn(logger, 'warn').mockImplementation(() => undefined)
  vi.spyOn(logger, 'info').mockImplementation(() => undefined)
  vi.spyOn(logger, 'error').mockImplementation(() => undefined)
})

afterEach(() => {
  vi.restoreAllMocks()
  process.env.HOME = ORIG_HOME
  process.env.APPDATA = ORIG_APPDATA
  fs.rmSync(TMP_DIR, { recursive: true, force: true })
})

afterAll(() => {
  os.homedir = ORIG_HOMEDIR
})

describe('checkForUpdate', () => {
  it('detects a newer release and writes the state file', async () => {
    const fetchSpy = vi.spyOn(globalThis, 'fetch').mockResolvedValue(fakeRelease('v0.2.0'))
    const res = await checkForUpdate(true)
    expect(res.updateAvailable).toBe(true)
    expect(res.latestVersion).toBe('0.2.0')
    expect(res.currentVersion).toBe(HOST_VERSION)
    expect(fetchSpy).toHaveBeenCalledTimes(1)
    const state = readStateFile()
    expect(state).not.toBeNull()
    expect(state!.latestVersion).toBe('0.2.0')
    expect(state!.downloadUrl).toBe(
      process.platform === 'win32' ? 'https://example.com/win.zip' : 'https://example.com/mac.zip',
    )
  })

  it('uses the cache when force is false and skips a second fetch', async () => {
    const fetchSpy = vi.spyOn(globalThis, 'fetch').mockResolvedValue(fakeRelease('v0.2.0'))
    await checkForUpdate(true)
    expect(fetchSpy).toHaveBeenCalledTimes(1)
    fetchSpy.mockClear()
    const cached = await checkForUpdate()
    expect(fetchSpy).not.toHaveBeenCalled()
    expect(cached.updateAvailable).toBe(true)
    expect(cached.latestVersion).toBe('0.2.0')
  })

  it('force=true bypasses the cache and fetches again', async () => {
    const fetchSpy = vi.spyOn(globalThis, 'fetch').mockResolvedValue(fakeRelease('v0.2.0'))
    await checkForUpdate(true)
    expect(fetchSpy).toHaveBeenCalledTimes(1)
    fetchSpy.mockClear()
    await checkForUpdate(true)
    expect(fetchSpy).toHaveBeenCalledTimes(1)
  })

  it('reports updateAvailable false when already up to date', async () => {
    vi.spyOn(globalThis, 'fetch').mockResolvedValue(fakeRelease('v0.1.0'))
    const res = await checkForUpdate(true)
    expect(res.updateAvailable).toBe(false)
    expect(res.latestVersion).toBe('0.1.0')
  })

  it('returns updateAvailable false and empty latestVersion on network failure', async () => {
    vi.spyOn(globalThis, 'fetch').mockRejectedValue(new Error('boom'))
    const res = await checkForUpdate(true)
    expect(res.updateAvailable).toBe(false)
    expect(res.latestVersion).toBe('')
    expect(res.currentVersion).toBe(HOST_VERSION)
  })
})

describe('getCachedUpdate', () => {
  it('returns null when no check has run', () => {
    expect(getCachedUpdate()).toBeNull()
  })

  it('returns the cached update after a successful check', async () => {
    vi.spyOn(globalThis, 'fetch').mockResolvedValue(fakeRelease('v0.2.0'))
    await checkForUpdate(true)
    const cached = getCachedUpdate()
    expect(cached).not.toBeNull()
    expect(cached!.updateAvailable).toBe(true)
    expect(cached!.latestVersion).toBe('0.2.0')
  })

  it('returns null when the cached version equals the current version', async () => {
    vi.spyOn(globalThis, 'fetch').mockResolvedValue(fakeRelease('v0.1.0'))
    await checkForUpdate(true)
    expect(getCachedUpdate()).toBeNull()
  })
})

describe('skipVersion', () => {
  it('marks the latest version as skipped so getCachedUpdate returns null', async () => {
    vi.spyOn(globalThis, 'fetch').mockResolvedValue(fakeRelease('v0.2.0'))
    await checkForUpdate(true)
    expect(getCachedUpdate()).not.toBeNull()
    skipVersion('0.2.0')
    expect(getCachedUpdate()).toBeNull()
    const state = readStateFile()
    expect(state!.skippedVersion).toBe('0.2.0')
  })
})

describe('applyUpdate', () => {
  it('fails when no update has been staged', async () => {
    const res = await applyUpdate()
    expect(res.ok).toBe(false)
    expect(res.message).toBe('No update staged')
  })
})
