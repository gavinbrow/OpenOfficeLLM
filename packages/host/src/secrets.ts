import fs from 'node:fs'
import crypto from 'node:crypto'
import os from 'node:os'
import process from 'node:process'
import { execFileSync } from 'node:child_process'
import { resolveSecretsPath, ensureDirs } from './paths.js'
import { logger } from './logging.js'

const FALLBACK_ENV_KEY_ENV = 'OPENOFFICELLM_SECRETS_FALLBACK_KEY'

// DPAPI is Windows-only and is invoked through Windows PowerShell (.NET
// System.Security.Cryptography.ProtectedData). The payload is delivered via
// stdin only — never on the command line — because process command lines are
// readable by other users on Windows.
const DPAPI_PREFIX = 'dpapi:'
const FB_PREFIX = 'fb:'

let dpapiProbed = false
let dpapiAvailable = false
let fallbackMode = false

/**
 * Probe DPAPI availability exactly once on Windows by protecting a short
 * constant string. On any non-Windows platform, or if the probe throws, we
 * permanently switch to the AES fallback. Subsequent calls return the cached
 * result.
 */
function probeDpapi(): void {
  if (dpapiProbed) return
  dpapiProbed = true

  if (process.platform !== 'win32') {
    dpapiAvailable = false
    fallbackMode = true
    logger.info({ msg: 'secrets: non-Windows host; using AES fallback' })
    return
  }

  try {
    // The probe uses a fixed, non-secret string. The result is discarded.
    const probeB64 = Buffer.from('openofficellm-dpapi-probe', 'utf8').toString('base64')
    dpapiRun('protect', probeB64)
    dpapiAvailable = true
    fallbackMode = false
    logger.info({ msg: 'secrets: DPAPI available via PowerShell' })
  } catch (e) {
    dpapiAvailable = false
    fallbackMode = true
    logger.warn({
      msg: 'DPAPI probe failed; falling back to env-key encrypted file. API keys will NOT be DPAPI-protected.',
      error: String((e as Error).message ?? e),
    })
  }
}

/**
 * Run a DPAPI protect/unprotect operation via Windows PowerShell. The base64
 * payload is embedded in the script text that is fed to PowerShell over stdin;
 * it never appears on the command line.
 */
function dpapiRun(op: 'protect' | 'unprotect', base64Payload: string): string {
  const method = op === 'protect' ? 'Protect' : 'Unprotect'
  const script =
    `$ErrorActionPreference='Stop'\n` +
    `Add-Type -AssemblyName System.Security\n` +
    `$b=[Convert]::FromBase64String('${base64Payload}')\n` +
    `$r=[System.Security.Cryptography.ProtectedData]::${method}($b,$null,[System.Security.Cryptography.DataProtectionScope]::CurrentUser)\n` +
    `[Convert]::ToBase64String($r)\n`

  const out = execFileSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', '-'], {
    input: script,
    encoding: 'utf8',
    windowsHide: true,
    timeout: 15000,
    maxBuffer: 8 * 1024 * 1024,
  })
  return out.trim()
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
  const buf = Buffer.from(plaintext, 'utf8').toString('base64')
  const enc = dpapiRun('protect', buf)
  return DPAPI_PREFIX + enc
}

function decryptDpapi(base64: string): string {
  const dec = dpapiRun('unprotect', base64)
  return Buffer.from(dec, 'base64').toString('utf8')
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

/**
 * Dispatch decryption based on the stored blob's prefix, never on whether
 * DPAPI currently works — so an existing 'fb:' secret still decrypts when
 * DPAPI is available, and a bare-base64 legacy DPAPI blob still decrypts
 * after the 'dpapi:' prefix was introduced.
 */
function decryptBlob(blob: string): string {
  if (blob.startsWith(FB_PREFIX)) {
    return decryptFallback(blob)
  }
  if (blob.startsWith(DPAPI_PREFIX)) {
    return decryptDpapi(blob.slice(DPAPI_PREFIX.length))
  }
  // Legacy bare-base64 DPAPI blob (pre-prefix). Try DPAPI unprotect.
  return decryptDpapi(blob)
}

export function setSecret(providerId: string, key: string): void {
  const id = providerId.trim()
  if (!id) throw new Error('providerId required')
  if (!key) throw new Error('key required')
  probeDpapi()
  const map = readStore()
  map[id] = dpapiAvailable ? encryptDpapi(key) : encryptFallback(key)
  writeStore(map)
}

export function getSecret(providerId: string): string | null {
  const id = providerId.trim()
  if (!id) return null
  const map = readStore()
  const blob = map[id]
  if (!blob) return null
  try {
    return decryptBlob(blob)
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
  if (!dpapiProbed) probeDpapi()
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