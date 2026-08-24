import fs from 'node:fs'
import path from 'node:path'
import crypto from 'node:crypto'
import os from 'node:os'
import process from 'node:process'
import { resolveSecretsPath, ensureDirs } from './paths.js'
import { logger } from './logging.js'

// `require` is available in CJS (the tsup bundle for the SEA binary) and in
// ESM via tsx (which provides it as an interop). We avoid `createRequire` +
// `import.meta.url` because the CJS bundle has no `import.meta`.
const requireModule: NodeRequire =
  typeof require !== 'undefined' ? require : (eval('require') as NodeRequire)

type DpapiModule = {
  protectData: (
    data: Uint8Array,
    entropy: Uint8Array | null,
    scope: 'CurrentUser' | 'LocalMachine',
  ) => Uint8Array
  unprotectData: (
    data: Uint8Array,
    entropy: Uint8Array | null,
    scope: 'CurrentUser' | 'LocalMachine',
  ) => Uint8Array
}

let dpapi: DpapiModule | null | undefined
let dpapiLoadAttempted = false
let fallbackMode = false

const FALLBACK_ENV_KEY_ENV = 'OPENOFFICELLM_SECRETS_FALLBACK_KEY'

/**
 * Load win-dpapi. Two paths:
 *
 * 1. Packaged (SEA / installer): the native addon ships as `win-dpapi.node`
 *    next to the host binary. `require()` cannot resolve it inside a SEA, so we
 *    load it explicitly via `process.dlopen` with an absolute path.
 * 2. Dev: normal `require('win-dpapi')`, which resolves through node_modules.
 *
 * If neither path works the caller falls back to AES-256-GCM — see the
 * encrypt/decryptFallback pair below. That keeps the app working when the
 * addon is absent, at the cost of weaker secret protection.
 */
function loadDpapi(): DpapiModule | null {
  if (dpapiLoadAttempted) return dpapi ?? null
  dpapiLoadAttempted = true

  // Packaged layout: <install-dir>/win-dpapi.node next to the host binary.
  const exeDir = path.dirname(process.execPath)
  const addonPath = path.join(exeDir, 'win-dpapi.node')
  if (fs.existsSync(addonPath)) {
    try {
      const mod: { exports: DpapiModule } = { exports: {} as DpapiModule }
      // process.dlopen needs the module object shape that require() would
      // produce — it writes to mod.exports.
      process.dlopen(mod, addonPath)
      if (
        typeof mod.exports.protectData !== 'function' ||
        typeof mod.exports.unprotectData !== 'function'
      ) {
        throw new Error('win-dpapi.node missing expected functions')
      }
      dpapi = mod.exports
      logger.info({ msg: 'secrets: win-dpapi loaded (packaged)', path: addonPath })
      return mod.exports
    } catch (e) {
      logger.warn({
        msg: 'win-dpapi.node found but failed to load; falling back',
        path: addonPath,
        error: String((e as Error).message ?? e),
      })
      dpapi = null
      fallbackMode = true
      return null
    }
  }

  // Dev layout: normal require through node_modules.
  try {
    const mod = requireModule('win-dpapi') as DpapiModule
    if (typeof mod.protectData !== 'function' || typeof mod.unprotectData !== 'function') {
      throw new Error('win-dpapi missing expected functions')
    }
    dpapi = mod
    logger.info({ msg: 'secrets: win-dpapi loaded' })
    return mod
  } catch (e) {
    dpapi = null
    fallbackMode = true
    logger.warn({
      msg: 'win-dpapi unavailable; falling back to env-key encrypted file. API keys will NOT be DPAPI-protected.',
      error: String((e as Error).message ?? e),
    })
    return null
  }
}

function getFallbackKey(): Buffer {
  const fromEnv = process.env[FALLBACK_ENV_KEY_ENV]
  if (fromEnv && fromEnv.length >= 32) {
    return Buffer.from(fromEnv, 'utf8').subarray(0, 32)
  }
  const machine = os.hostname() || 'fallback-host'
  const user = os.userInfo().username || 'fallback-user'
  const seed = `openofficellm::${machine}::${user}`
  const key = crypto.createHash('sha256').update(seed).digest()
  return key
}

function encryptDpapi(plaintext: string): string {
  const mod = loadDpapi()
  if (mod) {
    const buf = Buffer.from(plaintext, 'utf8')
    const enc = mod.protectData(buf, null, 'CurrentUser')
    return Buffer.from(enc).toString('base64')
  }
  return encryptFallback(plaintext)
}

function decryptDpapi(base64: string): string {
  const mod = loadDpapi()
  if (mod) {
    const buf = Buffer.from(base64, 'base64')
    const dec = mod.unprotectData(buf, null, 'CurrentUser')
    return Buffer.from(dec).toString('utf8')
  }
  return decryptFallback(base64)
}

function encryptFallback(plaintext: string): string {
  const key = getFallbackKey()
  const iv = crypto.randomBytes(16)
  const cipher = crypto.createCipheriv('aes-256-gcm', key, iv)
  const enc = Buffer.concat([cipher.update(plaintext, 'utf8'), cipher.final()])
  const tag = cipher.getAuthTag()
  return ['fb', iv.toString('base64'), tag.toString('base64'), enc.toString('base64')].join(':')
}

function decryptFallback(blob: string): string {
  const parts = blob.split(':')
  if (parts.length !== 4 || parts[0] !== 'fb') {
    throw new Error('malformed fallback secret blob')
  }
  const key = getFallbackKey()
  const iv = Buffer.from(parts[1], 'base64')
  const tag = Buffer.from(parts[2], 'base64')
  const enc = Buffer.from(parts[3], 'base64')
  const decipher = crypto.createDecipheriv('aes-256-gcm', key, iv)
  decipher.setAuthTag(tag)
  const dec = Buffer.concat([decipher.update(enc), decipher.final()])
  return dec.toString('utf8')
}

type SecretMap = Record<string, string>

function secretsPath(): string {
  return resolveSecretsPath()
}

function readStore(): SecretMap {
  try {
    const raw = fs.readFileSync(secretsPath(), 'utf8').trim()
    if (!raw) return {}
    const parsed = JSON.parse(raw) as SecretMap
    if (typeof parsed !== 'object' || parsed === null) return {}
    return parsed
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code === 'ENOENT') return {}
    logger.warn({ msg: 'secrets store unreadable, starting fresh', error: String(e) })
    return {}
  }
}

function writeStore(map: SecretMap): void {
  ensureDirs()
  const p = secretsPath()
  const tmp = p + '.tmp.' + process.pid
  fs.writeFileSync(tmp, JSON.stringify(map, null, 2), { encoding: 'utf8', mode: 0o600 })
  try {
    fs.renameSync(tmp, p)
  } catch (e) {
    try {
      fs.copyFileSync(tmp, p)
      fs.unlinkSync(tmp)
    } catch {
      try {
        fs.unlinkSync(tmp)
      } catch {
        // ignore
      }
      throw e
    }
  }
}

export function setSecret(providerId: string, key: string): void {
  const id = providerId.trim()
  if (!id) throw new Error('providerId required')
  if (!key) throw new Error('key required')
  loadDpapi()
  const map = readStore()
  map[id] = encryptDpapi(key)
  writeStore(map)
}

export function getSecret(providerId: string): string | null {
  const id = providerId.trim()
  if (!id) return null
  const map = readStore()
  const blob = map[id]
  if (!blob) return null
  try {
    return decryptDpapi(blob)
  } catch (e) {
    logger.error({ msg: 'secret decrypt failed', providerId: id, error: String(e) })
    return null
  }
}

export function deleteSecret(providerId: string): boolean {
  const id = providerId.trim()
  if (!id) return false
  const map = readStore()
  if (!(id in map)) return false
  delete map[id]
  writeStore(map)
  return true
}

export function listConfigured(): string[] {
  const map = readStore()
  return Object.keys(map).sort()
}

export function isFallbackMode(): boolean {
  if (!dpapiLoadAttempted) loadDpapi()
  return fallbackMode
}

export function secretsStoreExists(): boolean {
  try {
    fs.accessSync(secretsPath())
    return true
  } catch {
    return false
  }
}

export function redactSecretsFromObject<T>(obj: T): T {
  return obj
}
