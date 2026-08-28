import os from 'node:os'
import path from 'node:path'
import process from 'node:process'

/**
 * Where configuration, secrets, certs, logs and chat history live.
 *
 * Per-platform because each has a real convention, and a dot-folder in a Mac
 * home directory is the kind of small rudeness that compounds: Time Machine,
 * Migration Assistant and most backup tools know about Application Support and
 * nothing about `~/.openofficellm`.
 *
 * Windows keeps `%APPDATA%\OpenOfficeLLM` exactly as it was. Moving it would
 * strand every existing install's config, secrets and chat history behind a
 * path nothing looks at any more.
 */
function appDataDir(): string {
  if (process.platform === 'win32') {
    const appdata = process.env.APPDATA
    if (appdata && appdata.length > 0) return path.join(appdata, 'OpenOfficeLLM')
    return path.join(os.homedir(), 'AppData', 'Roaming', 'OpenOfficeLLM')
  }
  if (process.platform === 'darwin') {
    return path.join(os.homedir(), 'Library', 'Application Support', 'OpenOfficeLLM')
  }
  // Linux and the BSDs: XDG, falling back to the spec's own default.
  const xdg = process.env.XDG_CONFIG_HOME
  const base = xdg && xdg.length > 0 ? xdg : path.join(os.homedir(), '.config')
  return path.join(base, 'openofficellm')
}

function appDir(): string {
  return appDataDir()
}

export const APP_DIR = appDir()
export const CONFIG_PATH = path.join(APP_DIR, 'config.json')
export const LOG_DIR = path.join(APP_DIR, 'logs')
export const LOG_PATH = path.join(LOG_DIR, 'host.log')
export const LOCK_PATH = path.join(APP_DIR, 'host.lock')
export const CERT_DIR = path.join(APP_DIR, 'certs')
export const SECRETS_PATH = path.join(APP_DIR, 'secrets.dat')
export const CACHE_DIR = path.join(APP_DIR, 'cache')
export const OFFICEJS_CACHE_DIR = path.join(CACHE_DIR, 'officejs')
export const ATTACHMENTS_DIR = path.join(CACHE_DIR, 'attachments')
export const TESSERACT_CACHE_DIR = path.join(CACHE_DIR, 'tesseract')
export const MANIFEST_DIR = path.join(APP_DIR, 'manifest')
export const SKILLS_DIR = path.join(APP_DIR, 'skills')
export const UPDATE_STATE_PATH = path.join(APP_DIR, 'update-state.json')
export const UPDATE_STAGING_DIR = path.join(CACHE_DIR, 'update-pending')

export const DEFAULT_PORT = 7317
export const HOST_VERSION = '0.1.5'
export const HOST_INTERFACE = '127.0.0.1' as const

/** GitHub repo that hosts releases. The update feed is the releases/latest API. */
export const UPDATE_FEED = { owner: 'gavinbrow', repo: 'OpenOfficeLLM' } as const

import fs from 'node:fs'

export function ensureDirs(): void {
  const dir = appDir()
  for (const d of [
    dir,
    path.join(dir, 'logs'),
    path.join(dir, 'certs'),
    path.join(dir, 'cache'),
    path.join(dir, 'cache', 'officejs'),
    path.join(dir, 'cache', 'attachments'),
    path.join(dir, 'cache', 'tesseract'),
    path.join(dir, 'manifest'),
    path.join(dir, 'skills'),
    path.join(dir, 'cache', 'update-pending'),
  ]) {
    fs.mkdirSync(d, { recursive: true })
  }
}

export function resolveAppDir(): string {
  return appDir()
}

export function resolveSecretsPath(): string {
  return path.join(appDir(), 'secrets.dat')
}

export function resolveConfigPath(): string {
  return path.join(appDir(), 'config.json')
}

export function resolveLogDir(): string {
  return path.join(appDir(), 'logs')
}

export function resolveLogPath(): string {
  return path.join(resolveLogDir(), 'host.log')
}

export function resolveLockPath(): string {
  return path.join(appDir(), 'host.lock')
}

export function resolveCertDir(): string {
  return path.join(appDir(), 'certs')
}

export function resolveOfficejsCacheDir(): string {
  return path.join(appDir(), 'cache', 'officejs')
}

export function resolveAttachmentsDir(): string {
  return path.join(appDir(), 'cache', 'attachments')
}

export function resolveTesseractCacheDir(): string {
  return path.join(appDir(), 'cache', 'tesseract')
}

export function resolveManifestDir(): string {
  return path.join(appDir(), 'manifest')
}

export function resolveSkillsDir(): string {
  return path.join(appDir(), 'skills')
}

export function resolveUpdateStatePath(): string {
  return path.join(appDir(), 'update-state.json')
}

export function resolveUpdateStagingDir(): string {
  return path.join(appDir(), 'cache', 'update-pending')
}
