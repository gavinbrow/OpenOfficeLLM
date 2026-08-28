import fs from 'node:fs'
import path from 'node:path'
import process from 'node:process'
import * as childProcess from 'node:child_process'
import {
  HOST_VERSION,
  UPDATE_FEED,
  ensureDirs,
  resolveUpdateStatePath,
  resolveUpdateStagingDir,
} from './paths.js'
import { logger } from './logging.js'
import type { UpdateCheckResponse, UpdateApplyResponse } from '@openofficellm/shared'

interface UpdateState {
  lastCheckedAt: string | null
  latestVersion: string | null
  downloadUrl: string | null
  publishedAt: string | null
  releaseNotes: string | null
  skippedVersion: string | null
}

export const CHECK_INTERVAL_MS = 6 * 60 * 60 * 1000
const RELEASE_NOTES_MAX = 500

function defaultState(): UpdateState {
  return {
    lastCheckedAt: null,
    latestVersion: null,
    downloadUrl: null,
    publishedAt: null,
    releaseNotes: null,
    skippedVersion: null,
  }
}

function readState(): UpdateState {
  const p = resolveUpdateStatePath()
  try {
    const raw = fs.readFileSync(p, 'utf8').trim()
    if (!raw) return defaultState()
    const parsed = JSON.parse(raw) as Partial<UpdateState>
    return { ...defaultState(), ...parsed }
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code === 'ENOENT') return defaultState()
    logger.warn({ msg: 'update state unreadable, starting fresh', error: String(e) })
    return defaultState()
  }
}

function writeState(state: UpdateState): void {
  ensureDirs()
  const p = resolveUpdateStatePath()
  const tmp = p + '.tmp.' + process.pid
  fs.writeFileSync(tmp, JSON.stringify(state, null, 2), { encoding: 'utf8', mode: 0o600 })
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

function compareVersions(a: string, b: string): number {
  const pa = a.split('.').map((n) => Number(n) || 0)
  const pb = b.split('.').map((n) => Number(n) || 0)
  const len = Math.max(pa.length, pb.length)
  for (let i = 0; i < len; i++) {
    const da = pa[i] ?? 0
    const db = pb[i] ?? 0
    if (da > db) return 1
    if (da < db) return -1
  }
  return 0
}

interface GitHubAsset {
  name: string
  browser_download_url: string
}

interface GitHubRelease {
  tag_name: string
  published_at?: string
  body?: string
  assets?: GitHubAsset[]
}

function matchAsset(assets: GitHubAsset[]): string | null {
  const isMac = process.platform === 'darwin'
  const isWin = process.platform === 'win32'
  for (const a of assets) {
    if (isMac && a.name.endsWith('.zip') && /macOS/i.test(a.name)) return a.browser_download_url
    if (isWin && a.name.endsWith('.zip') && /win/i.test(a.name)) return a.browser_download_url
  }
  return null
}

export async function checkForUpdate(force?: boolean): Promise<UpdateCheckResponse> {
  const state = readState()

  if (!force && state.lastCheckedAt && state.latestVersion) {
    const age = Date.now() - new Date(state.lastCheckedAt).getTime()
    if (age < CHECK_INTERVAL_MS) {
      const updateAvailable =
        compareVersions(state.latestVersion, HOST_VERSION) > 0 &&
        state.latestVersion !== state.skippedVersion
      return {
        updateAvailable,
        currentVersion: HOST_VERSION,
        latestVersion: state.latestVersion,
        publishedAt: state.publishedAt ?? undefined,
        releaseNotes: state.releaseNotes ?? undefined,
      }
    }
  }

  const url = `https://api.github.com/repos/${UPDATE_FEED.owner}/${UPDATE_FEED.repo}/releases/latest`
  const headers = {
    Accept: 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28',
    'User-Agent': 'OpenOfficeLLM-Host',
  }

  try {
    const ctrl = new AbortController()
    const timer = setTimeout(() => ctrl.abort(), 10_000)
    let res: Response
    try {
      res = await fetch(url, { headers, signal: ctrl.signal })
    } finally {
      clearTimeout(timer)
    }
    if (!res.ok) {
      logger.warn({ msg: 'update check non-ok response', status: res.status })
      return { updateAvailable: false, currentVersion: HOST_VERSION, latestVersion: '' }
    }
    const release = (await res.json()) as GitHubRelease
    const tag = release.tag_name ?? ''
    const latest = tag.startsWith('v') ? tag.slice(1) : tag
    if (!latest) {
      return { updateAvailable: false, currentVersion: HOST_VERSION, latestVersion: '' }
    }
    const downloadUrl = release.assets ? matchAsset(release.assets) : null
    const publishedAt = release.published_at ?? null
    const releaseNotes = release.body ? release.body.slice(0, RELEASE_NOTES_MAX) : null

    const nextState: UpdateState = {
      lastCheckedAt: new Date().toISOString(),
      latestVersion: latest,
      downloadUrl,
      publishedAt,
      releaseNotes,
      skippedVersion: state.skippedVersion,
    }
    writeState(nextState)

    const updateAvailable = compareVersions(latest, HOST_VERSION) > 0
    return {
      updateAvailable,
      currentVersion: HOST_VERSION,
      latestVersion: latest,
      publishedAt: publishedAt ?? undefined,
      releaseNotes: releaseNotes ?? undefined,
    }
  } catch (e) {
    logger.warn({ msg: 'update check failed', error: String(e) })
    return { updateAvailable: false, currentVersion: HOST_VERSION, latestVersion: '' }
  }
}

export function getCachedUpdate(): UpdateCheckResponse | null {
  const state = readState()
  if (!state.latestVersion) return null
  if (state.latestVersion === state.skippedVersion) return null
  if (compareVersions(state.latestVersion, HOST_VERSION) <= 0) return null
  return {
    updateAvailable: true,
    currentVersion: HOST_VERSION,
    latestVersion: state.latestVersion,
    publishedAt: state.publishedAt ?? undefined,
    releaseNotes: state.releaseNotes ?? undefined,
  }
}

export function skipVersion(version: string): void {
  const state = readState()
  writeState({ ...state, skippedVersion: version })
}

const MACOS_HELPER = `#!/bin/sh
# Auto-generated by OpenOfficeLLM host. Waits for the host to exit, then
# swaps the .app bundle and relaunches.
set -e
STAGING_DIR="$(cd "$(dirname "$0")" && pwd)"
APP_DIR="/Applications/OpenOfficeLLM.app"

# Wait for the host to exit (give it time to shut down gracefully).
sleep 3

# Unzip the downloaded archive.
cd "$STAGING_DIR"
unzip -o -q update.zip

# If the zip contained a .app, sync it over the installed app.
if [ -d "OpenOfficeLLM.app" ]; then
  # Preserve any user-added files (none expected, but be safe).
  rsync -a --delete "OpenOfficeLLM.app/" "$APP_DIR/"
  # Re-sign ad-hoc so Gatekeeper/launchd accept the new binary.
  codesign --force --deep --sign - "$APP_DIR"
  # Strip quarantine, or the update undoes the user's own approval.
  #
  # Gatekeeper only evaluates bundles carrying com.apple.quarantine. An app
  # installed from the downloaded DMG has it, and the user cleared the
  # resulting "Apple could not verify..." dialog once by hand. rsync syncs the
  # bundle's CONTENTS into an existing directory, so $APP_DIR keeps its own
  # quarantine xattr, while codesign above gives the bundle a new cdhash --
  # which voids the approval that was pinned to the old one. Quarantined again
  # and unrecognised again, the very next launch re-runs the malware dialog.
  #
  # Until the app is Developer ID-signed and notarized (Docs/TODO.md), that
  # dialog is unavoidable on first install but must not come back on every
  # update. Removing the attribute here is what makes "you only have to do
  # this once" true.
  xattr -dr com.apple.quarantine "$APP_DIR" 2>/dev/null || true
fi

# Relaunch via open — the launcher re-runs --install then starts the host.
open "$APP_DIR"
`

const WINDOWS_HELPER = `@echo off
REM Auto-generated by OpenOfficeLLM host. Waits for the host to exit, then
REM swaps the binary and web bundle and relaunches.
setlocal
set STAGING_DIR=%~dp0
set APP_DIR=%LOCALAPPDATA%\\Programs\\OpenOfficeLLM

REM Wait for the host to exit. 3s gives the host time to flush its response
REM and shut down gracefully after the SIGTERM.
timeout /t 3 /nobreak >nul

REM Unzip (PowerShell's Expand-Archive is always available).
powershell -NoProfile -Command "Expand-Archive -Force -Path '%STAGING_DIR%update.zip' -DestinationPath '%STAGING_DIR%extracted'"
if errorlevel 1 (
  echo Failed to extract update zip >> "%STAGING_DIR%update-error.log"
  exit /b 1
)

REM Swap the binary. move /y fails if the file is locked, but the host
REM has exited by now (timeout 3 above).
if exist "%APP_DIR%\\host.exe" move /y "%APP_DIR%\\host.exe" "%APP_DIR%\\host.exe.old" >nul 2>&1
move /y "%STAGING_DIR%extracted\\host.exe" "%APP_DIR%\\host.exe"
if errorlevel 1 (
  echo Failed to swap host.exe >> "%STAGING_DIR%update-error.log"
  exit /b 1
)
if exist "%STAGING_DIR%extracted\\web" (
  if exist "%APP_DIR%\\web" rmdir /s /q "%APP_DIR%\\web"
  xcopy /e /i /q "%STAGING_DIR%extracted\\web" "%APP_DIR%\\web" >nul
)
if exist "%STAGING_DIR%extracted\\version.txt" copy /y "%STAGING_DIR%extracted\\version.txt" "%APP_DIR%\\version.txt" >nul

REM Re-provision (idempotent: trust CA, write manifest, register add-in).
"%APP_DIR%\\host.exe" --install

REM Relaunch the host.
start "" "%APP_DIR%\\host.exe"

REM Clean up.
rd /s /q "%STAGING_DIR%" 2>nul
`

function writeHelperScript(stagingDir: string): { helperPath: string; helperArgs: string[] } {
  if (process.platform === 'win32') {
    const helperPath = path.join(stagingDir, 'apply-update.cmd')
    fs.writeFileSync(helperPath, WINDOWS_HELPER, { encoding: 'utf8' })
    return { helperPath, helperArgs: [] }
  }
  const helperPath = path.join(stagingDir, 'apply-update.sh')
  fs.writeFileSync(helperPath, MACOS_HELPER, { encoding: 'utf8', mode: 0o755 })
  return { helperPath, helperArgs: [] }
}

export async function applyUpdate(): Promise<UpdateApplyResponse> {
  const state = readState()
  if (!state.downloadUrl) {
    return { ok: false, message: 'No update staged' }
  }

  ensureDirs()
  const stagingDir = resolveUpdateStagingDir()
  const zipPath = path.join(stagingDir, 'update.zip')

  try {
    const ctrl = new AbortController()
    const timer = setTimeout(() => ctrl.abort(), 120_000)
    let res: Response
    try {
      res = await fetch(state.downloadUrl, { signal: ctrl.signal })
    } finally {
      clearTimeout(timer)
    }
    if (!res.ok || !res.body) {
      logger.warn({ msg: 'update download failed', status: res.status })
      return { ok: false, message: 'Download failed' }
    }
    const buf = Buffer.from(await res.arrayBuffer())
    fs.writeFileSync(zipPath, buf, { mode: 0o600 })
  } catch (e) {
    logger.error({ msg: 'update download error', error: String(e) })
    return { ok: false, message: 'Download failed' }
  }

  const { helperPath, helperArgs } = writeHelperScript(stagingDir)

  try {
    // The helper must survive the host's imminent SIGTERM. Node's
    // `spawn(..., { detached: true, shell: true })` on Windows launches the
    // helper via `cmd.exe`, but that cmd.exe is still in the host's process
    // tree — when the host is killed, cmd.exe and its child die with it,
    // and the update never applies.
    //
    // `cmd /c start /b` creates a truly independent process: `start` opens a
    // new process group and `/b` keeps it windowless. The helper runs
    // outside the host's lifetime, which is the whole point — it waits for
    // the host to exit, then swaps the binary.
    //
    // On macOS, `nohup` + `&` in a `sh -c` achieves the same independence
    // from the parent's process group, and `detached` + `unref` ensures
    // Node won't wait for it.
    if (process.platform === 'win32') {
      const child = childProcess.spawn(
        'cmd.exe',
        ['/c', 'start', '/b', '', helperPath, ...helperArgs],
        { detached: true, stdio: 'ignore', windowsVerbatimArguments: true },
      )
      child.unref()
    } else {
      const child = childProcess.spawn(
        'nohup',
        ['sh', '-c', `"${helperPath}" ${helperArgs.join(' ')}`],
        { detached: true, stdio: 'ignore' },
      )
      child.unref()
    }
  } catch (e) {
    logger.error({ msg: 'failed to spawn update helper', error: String(e) })
    return { ok: false, message: 'Failed to start helper' }
  }

  logger.info({ msg: 'update helper spawned', stagingDir })
  return { ok: true }
}
