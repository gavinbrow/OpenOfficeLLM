// Add-in registration and logon autostart for Office.
//
// Everything here is per-user and needs no elevation, which is a hard product
// requirement — see P6.10. The two supported platforms disagree on the
// mechanism: Windows keeps both the sideload registration and autostart in the
// HKCU registry (WEF\Developer and the Run key), while macOS sideloads by
// copying the manifest into each Office app's wef folder and autostarts
// through a per-user LaunchAgent. Each function below checks os.platform() and
// dispatches; the Windows paths are the PowerShell scripts they always were.

import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import process from 'node:process'
import { runPowershell, escapePs, runCommand } from './trust.js'
import { logger } from './logging.js'
import { MANIFEST_FILENAME } from './manifest.js'
import {
  LAUNCH_AGENT_LABEL,
  writeLaunchAgentPlist,
  readLaunchAgentCommand,
  removeLaunchAgentPlist,
} from './autostart.js'

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

// ─── macOS registration (wef sideload folders) ────────────────────────────
//
// macOS has no registry; Office's developer sideloading there means copying
// the manifest into each app's container:
//
//   ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
//   ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
//
// The wef folder is only created by an app that has run at least once, so its
// presence doubles as the "is this app installed" check.

/** Bundle ids of the Office apps we sideload into. */
const DARWIN_OFFICE_APPS = ['com.microsoft.Word', 'com.microsoft.Excel']

/** WEF sideload folder for one Office app under a given home directory. */
function wefDirForApp(homeDir: string, app: string): string {
  return path.join(homeDir, 'Library', 'Containers', app, 'Data', 'Documents', 'wef')
}

/** WEF sideload folders for Word and Excel under a given home directory. */
export function wefDirs(homeDir: string = os.homedir()): Record<string, string> {
  const entries: Record<string, string> = {}
  for (const app of DARWIN_OFFICE_APPS) {
    const wefPath = wefDirForApp(homeDir, app)
    entries[wefPath] = wefPath
  }
  return entries
}

function wefManifestPath(wefPath: string, filename: string): string {
  return path.join(wefPath, filename)
}

function fileMatches(sourcePath: string, targetPath: string): boolean {
  try {
    return fs.readFileSync(sourcePath).equals(fs.readFileSync(targetPath))
  } catch {
    return false
  }
}

export async function readRegistrationDarwin(manifest?: string): Promise<RegistrationState> {
  const entries: Record<string, string> = {}
  const stale: string[] = []
  for (const wefPath of Object.keys(wefDirs())) {
    const wefFile = wefManifestPath(wefPath, MANIFEST_FILENAME)
    if (fs.existsSync(wefFile)) {
      entries[wefPath] = wefPath
      // An entry is stale when its own file is gone (impossible here, since
      // we just checked it) or the manifest it was copied from no longer
      // exists — a sideloaded copy of a deleted manifest is a dangling
      // registration, exactly like a Windows entry whose manifest file has
      // gone. Entries are only listed when present, so staleness on macOS
      // reduces to "the copy exists but the source manifest does not".
      if (manifest && !fs.existsSync(manifest)) stale.push(wefPath)
    }
  }
  const registered = manifest
    ? Object.keys(entries).some((wefPath) =>
        fileMatches(manifest, wefManifestPath(wefPath, MANIFEST_FILENAME)),
      )
    : Object.keys(entries).length > 0
  return { registered, entries, stale }
}

function existingOfficeContainer(home: string, app: string): boolean {
  return fs.existsSync(path.join(home, 'Library', 'Containers', app))
}

export async function registerAddinDarwin(manifest: string): Promise<{ changed: boolean }> {
  const abs = path.resolve(manifest)
  if (!fs.existsSync(abs)) {
    throw new Error(`Manifest not found at ${abs} — start the host once to generate it.`)
  }
  const home = os.homedir()
  const present = DARWIN_OFFICE_APPS.filter((app) => existingOfficeContainer(home, app))
  const missing = DARWIN_OFFICE_APPS.filter((app) => !existingOfficeContainer(home, app))
  if (present.length === 0) {
    throw new Error(
      'Neither Word nor Excel was found on this Mac. Install Office, start Word or Excel ' +
        'once so its wef folder exists, then re-run --install.',
    )
  }
  if (missing.length > 0) {
    logger.warn({
      msg: 'Office app not installed, skipping its wef folder',
      missing,
    })
  }
  let changed = false
  for (const app of present) {
    const wefPath = wefDirForApp(home, app)
    const target = wefManifestPath(wefPath, MANIFEST_FILENAME)
    if (!fs.existsSync(target) || !fileMatches(abs, target)) {
      fs.mkdirSync(wefPath, { recursive: true })
      fs.copyFileSync(abs, target)
      changed = true
    }
  }
  return { changed }
}

export async function unregisterAddinDarwin(_manifest: string): Promise<boolean> {
  let removed = false
  for (const wefPath of Object.keys(wefDirs())) {
    const wefFile = wefManifestPath(wefPath, MANIFEST_FILENAME)
    if (fs.existsSync(wefFile)) {
      try {
        fs.unlinkSync(wefFile)
        removed = true
      } catch {
        // ignore — best effort; the copy stays but nothing breaks
      }
    }
  }
  return removed
}

// ─── macOS autostart (LaunchAgent) ────────────────────────────────────────
//
// A per-user LaunchAgent is the launchd-idiomatic sibling of the Windows Run
// key: loaded at logon into the user's GUI session, restartable, and removed
// without touching anything system-wide.

/** Child-process runner used for launchctl. Tests replace this so the real
 *  launchctl is never invoked; `resetDarwinRunner` restores it. */
let darwinRunner: typeof runCommand = runCommand

/** @internal — test seam. */
export function setDarwinRunner(runner: typeof runCommand): void {
  darwinRunner = runner
}

/** @internal — test seam. */
export function resetDarwinRunner(): void {
  darwinRunner = runCommand
}

/** The user's GUI launchd domain. `id -u` runs the current user; launchctl
 *  wants `gui/<uid>`, not `user/<uid>`. */
async function launchdGuiDomain(): Promise<string> {
  const { code, stdout } = await darwinRunner('id', ['-u'])
  return `gui/${code === 0 ? stdout.trim() : (process.getuid?.() ?? '')}`
}

/**
 * argv for the LaunchAgent, the full command launchd will exec.
 *
 * ProgramArguments is an exec-array, not a command line, so the Windows
 * quoting in autostartCommand() must not be carried over — quotes are
 * literal there and a quoted path fails to exec. In dev the entry is
 * `<node> <script> --no-browser`; packaged it is `<binary> --no-browser`.
 */
function darwinLaunchArguments(): string[] {
  const exe = process.execPath
  const script = process.argv[1]
  if (script && path.basename(exe).toLowerCase().startsWith('node')) {
    return [exe, path.resolve(script), '--no-browser']
  }
  return [exe, '--no-browser']
}

export async function installAutostartDarwin(): Promise<string> {
  const args = darwinLaunchArguments()
  const plistPath = writeLaunchAgentPlist(os.homedir(), args)
  const domain = await launchdGuiDomain()
  // Bootstrap into the user's GUI domain. exit code 5 means the service is
  // already loaded, which is exactly what a second --install produces — that
  // is idempotence, not failure.
  const { code, stderr } = await darwinRunner('launchctl', ['bootstrap', domain, plistPath])
  if (code === 0 || code === 5 || /already loaded/i.test(stderr)) {
    if (code !== 0) {
      logger.info({ msg: 'autostart service was already loaded', code })
    }
    return args.join(' ')
  }
  logger.warn({
    msg: 'LaunchAgent bootstrap failed; run it manually',
    command: `launchctl bootstrap ${domain} ${plistPath}`,
    code,
    stderr: stderr.trim(),
  })
  return args.join(' ')
}

export async function removeAutostartDarwin(): Promise<boolean> {
  // bootout errors when the service is not loaded; either way the plist is
  // what we are actually removing.
  const domain = await launchdGuiDomain()
  const { code } = await darwinRunner('launchctl', ['bootout', `${domain}/${LAUNCH_AGENT_LABEL}`])
  return removeLaunchAgentPlist(os.homedir()) || code === 0
}

export async function readAutostartDarwin(): Promise<string | null> {
  return readLaunchAgentCommand(os.homedir())
}

// ─── platform dispatch ────────────────────────────────────────────────────

/**
 * Read the WEF Developer key.
 *
 * Office treats this key as a flat map of manifest path -> manifest path; the
 * value *name* is what it actually reads. It is registry-native to have the
 * name and data be identical here, odd as that looks.
 */
export async function readRegistration(manifest?: string): Promise<RegistrationState> {
  if (os.platform() === 'darwin') return readRegistrationDarwin(manifest)
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
  if (os.platform() === 'darwin') return registerAddinDarwin(manifest)
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
  if (os.platform() === 'darwin') return unregisterAddinDarwin(manifest)
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
export function autostartCommand(): string {
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
  if (os.platform() === 'darwin') return installAutostartDarwin()
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
  if (os.platform() === 'darwin') return removeAutostartDarwin()
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
  if (os.platform() === 'darwin') return readAutostartDarwin()
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
