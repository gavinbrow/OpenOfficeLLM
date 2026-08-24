// Windows registration for the add-in: the WEF sideload key and logon autostart.
//
// Everything here is per-user (HKCU) and needs no elevation, which is a hard
// product requirement — see P6.10.

import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import process from 'node:process'
import { runPowershell, escapePs } from './trust.js'
import { logger } from './logging.js'

const WEF_DEVELOPER_KEY = 'HKCU:\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer'
const RUN_KEY = 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Run'
const RUN_VALUE_NAME = 'OpenOfficeLLM Host'

export interface RegistrationState {
  registered: boolean
  /** Every value currently under the Developer key, path -> data. */
  entries: Record<string, string>
  /** Entries whose manifest file no longer exists on disk. */
  stale: string[]
}

function requireWindows(): void {
  if (os.platform() !== 'win32') {
    throw new Error('Add-in registration is only supported on Windows')
  }
}

/**
 * Read the WEF Developer key.
 *
 * Office treats this key as a flat map of manifest path -> manifest path; the
 * value *name* is what it actually reads. It is registry-native to have the
 * name and data be identical here, odd as that looks.
 */
export async function readRegistration(manifest?: string): Promise<RegistrationState> {
  requireWindows()
  const script = `
$ErrorActionPreference = 'SilentlyContinue'
$key = ${escapePs(WEF_DEVELOPER_KEY)}
if (-not (Test-Path $key)) { '{}' ; exit 0 }
$props = Get-ItemProperty -Path $key
$out = @{}
foreach ($p in $props.PSObject.Properties) {
  if ($p.Name -like 'PS*') { continue }
  $out[$p.Name] = [string]$p.Value
}
$out | ConvertTo-Json -Compress -Depth 3
`.trim()
  const { code, stdout } = await runPowershell(script)
  let entries: Record<string, string> = {}
  if (code === 0) {
    const raw = stdout.trim()
    if (raw && raw !== '{}') {
      try {
        const parsed: unknown = JSON.parse(raw)
        if (parsed && typeof parsed === 'object') {
          entries = parsed as Record<string, string>
        }
      } catch {
        logger.warn({ msg: 'could not parse WEF Developer key', raw })
      }
    }
  }
  const stale = Object.keys(entries).filter((p) => {
    try {
      return !fs.statSync(p).isFile()
    } catch {
      return true
    }
  })
  const registered = manifest
    ? Object.keys(entries).some(
        (p) => path.resolve(p).toLowerCase() === path.resolve(manifest).toLowerCase(),
      )
    : Object.keys(entries).length > 0
  return { registered, entries, stale }
}

/**
 * Point the WEF Developer key at our manifest.
 *
 * Also drops any value whose manifest file is gone. A dangling entry is not
 * harmless: Office walks every value in this key while building the ribbon,
 * and entries left behind by uninstalled or moved add-ins are a known way to
 * end up with none of them loading.
 */
export async function registerAddin(manifest: string): Promise<{ changed: boolean }> {
  requireWindows()
  const abs = path.resolve(manifest)
  if (!fs.existsSync(abs)) {
    throw new Error(`Manifest not found at ${abs} — start the host once to generate it.`)
  }
  const before = await readRegistration(abs)
  const script = `
$ErrorActionPreference = 'Stop'
$key = ${escapePs(WEF_DEVELOPER_KEY)}
$manifest = ${escapePs(abs)}
if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
${before.stale.map((s) => `Remove-ItemProperty -Path $key -Name ${escapePs(s)} -ErrorAction SilentlyContinue`).join('\n')}
New-ItemProperty -Path $key -Name $manifest -Value $manifest -PropertyType String -Force | Out-Null
Write-Output 'OK'
`.trim()
  const { code, stdout, stderr } = await runPowershell(script)
  if (code !== 0 || stdout.trim().split(/\r?\n/).pop() !== 'OK') {
    throw new Error(`Failed to register add-in (code ${code}): ${stderr.trim() || stdout.trim()}`)
  }
  if (before.stale.length > 0) {
    logger.info({ msg: 'removed stale WEF Developer entries', removed: before.stale })
  }
  return { changed: !before.registered || before.stale.length > 0 }
}

export async function unregisterAddin(manifest: string): Promise<boolean> {
  requireWindows()
  const abs = path.resolve(manifest)
  const script = `
$ErrorActionPreference = 'SilentlyContinue'
$key = ${escapePs(WEF_DEVELOPER_KEY)}
$manifest = ${escapePs(abs)}
if (-not (Test-Path $key)) { Write-Output 'NOTFOUND'; exit 0 }
$existing = Get-ItemProperty -Path $key -Name $manifest -ErrorAction SilentlyContinue
if (-not $existing) { Write-Output 'NOTFOUND'; exit 0 }
Remove-ItemProperty -Path $key -Name $manifest
Write-Output 'REMOVED'
`.trim()
  const { code, stdout } = await runPowershell(script)
  if (code !== 0) return false
  return stdout.trim().split(/\r?\n/).pop() === 'REMOVED'
}

/** Command line that relaunches this host exactly as it is running now. */
function autostartCommand(): string {
  // process.execPath is node.exe during development and the packaged binary
  // once P6.7 lands. In the dev case we must also pass the script path, which
  // argv[1] carries; in the packaged case argv[1] is absent and execPath alone
  // is the whole command.
  const exe = process.execPath
  const script = process.argv[1]
  const quoted = (s: string) => `"${s}"`
  if (script && path.basename(exe).toLowerCase().startsWith('node')) {
    return `${quoted(exe)} ${quoted(path.resolve(script))}`
  }
  return quoted(exe)
}

export async function installAutostart(): Promise<string> {
  requireWindows()
  const command = autostartCommand()
  const script = `
$ErrorActionPreference = 'Stop'
New-ItemProperty -Path ${escapePs(RUN_KEY)} -Name ${escapePs(RUN_VALUE_NAME)} \`
  -Value ${escapePs(command)} -PropertyType String -Force | Out-Null
Write-Output 'OK'
`.trim()
  const { code, stdout, stderr } = await runPowershell(script)
  if (code !== 0 || stdout.trim().split(/\r?\n/).pop() !== 'OK') {
    throw new Error(`Failed to install autostart (code ${code}): ${stderr.trim() || stdout.trim()}`)
  }
  return command
}

export async function removeAutostart(): Promise<boolean> {
  requireWindows()
  const script = `
$ErrorActionPreference = 'SilentlyContinue'
Remove-ItemProperty -Path ${escapePs(RUN_KEY)} -Name ${escapePs(RUN_VALUE_NAME)}
if ($?) { Write-Output 'REMOVED' } else { Write-Output 'NOTFOUND' }
`.trim()
  const { code, stdout } = await runPowershell(script)
  if (code !== 0) return false
  return stdout.trim().split(/\r?\n/).pop() === 'REMOVED'
}

export async function readAutostart(): Promise<string | null> {
  requireWindows()
  // Get-ItemPropertyValue rather than (Get-ItemProperty ...).Name — the value
  // name contains a space, which makes the property-access form awkward to
  // quote correctly through two layers of escaping.
  const script = `
$ErrorActionPreference = 'SilentlyContinue'
$v = Get-ItemPropertyValue -Path ${escapePs(RUN_KEY)} -Name ${escapePs(RUN_VALUE_NAME)}
if ($v) { Write-Output $v }
`.trim()
  const { code, stdout } = await runPowershell(script)
  if (code !== 0) return null
  const v = stdout.trim()
  return v.length > 0 ? v : null
}
