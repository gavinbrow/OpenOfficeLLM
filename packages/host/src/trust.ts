import { spawn } from 'node:child_process'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import { logger } from './logging.js'
import { resolveCertDir } from './paths.js'

const PS_ENCODING = 'utf8' as const

function powershell(script: string): Promise<{ code: number; stdout: string; stderr: string }> {
  return new Promise((resolve) => {
    const child = spawn(
      os.platform() === 'win32' ? 'powershell.exe' : 'pwsh',
      ['-NoProfile', '-NonInteractive', '-Command', script],
      { windowsHide: true },
    )
    let stdout = ''
    let stderr = ''
    child.stdout.setEncoding(PS_ENCODING)
    child.stderr.setEncoding(PS_ENCODING)
    child.stdout.on('data', (c) => (stdout += c))
    child.stderr.on('data', (c) => (stderr += c))
    child.on('close', (code) => resolve({ code: code ?? -1, stdout, stderr }))
    child.on('error', (e) => resolve({ code: -1, stdout, stderr: String(e) }))
  })
}

/** Run a command directly, for the platforms that are not driven through
 *  PowerShell. Same result shape as `powershell` so callers read alike. */
function run(
  cmd: string,
  args: string[],
): Promise<{ code: number; stdout: string; stderr: string }> {
  return new Promise((resolve) => {
    const child = spawn(cmd, args, { windowsHide: true })
    let stdout = ''
    let stderr = ''
    child.stdout.setEncoding('utf8')
    child.stderr.setEncoding('utf8')
    child.stdout.on('data', (c) => (stdout += c))
    child.stderr.on('data', (c) => (stderr += c))
    child.on('close', (code) => resolve({ code: code ?? -1, stdout, stderr }))
    child.on('error', (e) => resolve({ code: -1, stdout, stderr: String(e) }))
  })
}

function escapePsSingle(value: string): string {
  return "'" + value.replace(/'/g, "''") + "'"
}

export async function isTrusted(thumbprint: string): Promise<boolean> {
  if (os.platform() === 'darwin') return isTrustedDarwin(thumbprint)
  if (os.platform() !== 'win32') return false
  const script = `
$ErrorActionPreference = 'Stop'
$tp = ${escapePsSingle(thumbprint)}
$found = Get-ChildItem Cert:\\CurrentUser\\Root -ErrorAction SilentlyContinue |
  Where-Object { $_.Thumbprint -eq $tp }
if ($found) { 'yes' } else { 'no' }
`.trim()
  const { code, stdout, stderr } = await powershell(script)
  if (code !== 0) {
    logger.warn({ msg: 'trust lookup failed', code, stderr: stderr.trim() })
    return false
  }
  return stdout.trim() === 'yes'
}

export async function trustLocalCa(caPem: string): Promise<string | null> {
  if (os.platform() === 'darwin') return trustLocalCaDarwin(caPem)
  if (os.platform() !== 'win32') {
    throw new Error(
      'Automatic CA trust is implemented for Windows and macOS. On Linux, add the CA in ' +
        `${resolveCertDir()} to your browser's certificate store manually.`,
    )
  }
  // Write the PEM to a temp file — X509Certificate2 needs a file path. The
  // caller (index.ts) passes the PEM content, not a path, so we materialize
  // it here and clean up after the PowerShell call.
  const tmpDir = os.tmpdir()
  const tmpPath = path.join(tmpDir, `openofficellm-ca-${Date.now()}.crt`)
  fs.writeFileSync(tmpPath, caPem, { mode: 0o600 })
  try {
    const script = `
$ErrorActionPreference = 'Stop'
$caPath = ${escapePsSingle(tmpPath)}
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $caPath
$tp = $cert.Thumbprint
$existing = Get-ChildItem Cert:\\CurrentUser\\Root -ErrorAction SilentlyContinue |
  Where-Object { $_.Thumbprint -eq $tp }
if ($existing) {
  Write-Output ('ALREADY:' + $tp)
} else {
  $store = New-Object System.Security.Cryptography.X509Certificates.X509Store 'Root','CurrentUser'
  $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
  try { $store.Add($cert) } finally { $store.Close() }
  Write-Output ('ADDED:' + $cert.Thumbprint)
}
`.trim()
    const { code, stdout, stderr } = await powershell(script)
    if (code !== 0) {
      throw new Error(`X509Store trust failed (code ${code}): ${stderr.trim() || stdout.trim()}`)
    }
    const line = stdout.trim().split(/\r?\n/).pop() ?? ''
    if (line.startsWith('ALREADY:')) return null
    if (line.startsWith('ADDED:')) return line.slice('ADDED:'.length)
    throw new Error(`Unexpected trust output: ${stdout.trim()}`)
  } finally {
    try {
      fs.unlinkSync(tmpPath)
    } catch {
      // best-effort cleanup
    }
  }
}

export async function removeTrustedCa(thumbprint: string): Promise<boolean> {
  if (os.platform() === 'darwin') return removeTrustedCaDarwin(thumbprint)
  if (os.platform() !== 'win32') return false
  // Use X509Store.Remove directly instead of Remove-Item on PSPath —
  // Remove-Item on the Root store can trigger a confirmation dialog that
  // -Force does not suppress for the protected Root store. X509Store.Remove
  // avoids the CryptUI path entirely.
  const script = `
$ErrorActionPreference = 'Stop'
$tp = ${escapePsSingle(thumbprint)}
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store 'Root','CurrentUser'
$store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
$found = $store.Certificates | Where-Object { $_.Thumbprint -eq $tp }
if (-not $found) { $store.Close(); Write-Output 'NOTFOUND'; exit 0 }
$found | ForEach-Object { $store.Remove($_) }
$store.Close()
Write-Output 'REMOVED'
`.trim()
  const { code, stdout, stderr } = await powershell(script)
  if (code !== 0) {
    logger.warn({ msg: 'remove trust failed', code, stderr: stderr.trim() })
    return false
  }
  const line = stdout.trim().split(/\r?\n/).pop() ?? ''
  return line === 'REMOVED'
}

export { powershell as runPowershell }
export { run as runCommand }
export { escapePsSingle as escapePs }

// ─── macOS ───────────────────────────────────────────────────────────────
//
// The `security` tool is the whole API here. Two things about it shape the
// code below.
//
// First, adding a *trusted* root to the login keychain opens a Touch ID /
// password prompt. That is correct and cannot be suppressed — the user is
// authorising a certificate authority for their whole account — but it means
// this can only run from a terminal the user is looking at, never from a
// background service start. `--install` is that terminal; `startServer` never
// calls it.
//
// Second, `security` speaks SHA-1 fingerprints with colons (`AB:CD:…`) while
// the Windows path and our stored thumbprint use unseparated uppercase hex.
// Normalising both sides is what keeps `--diagnose` from reporting a trusted
// CA as untrusted on a Mac.

/** The login keychain, where a per-user trust setting belongs. Trusting into
 *  the System keychain would require sudo and would affect every account on
 *  the machine — far more than a loopback development service should ask. */
function loginKeychain(): string {
  return path.join(os.homedir(), 'Library', 'Keychains', 'login.keychain-db')
}

function normalizeFingerprint(value: string): string {
  return value.replace(/[^0-9a-fA-F]/g, '').toUpperCase()
}

async function isTrustedDarwin(thumbprint: string): Promise<boolean> {
  const want = normalizeFingerprint(thumbprint)
  if (!want) return false
  // Presence in the keychain is not enough: `add-trusted-cert -r trustAsRoot`
  // stores the certificate with a trust record that `dump-trust-settings`
  // surfaces, but WKWebView/Word do not honour it for a loopback TLS chain —
  // the pane loads as an untrusted certificate. The authoritative check is
  // `verify-cert -p ssl`, which runs the same trust evaluation the webview
  // uses. We verify the CA on disk; a leaf that chains to it then verifies too.
  const caPath = path.join(resolveCertDir(), 'ca.crt')
  if (!fs.existsSync(caPath)) return false
  // Refuse to vouch for a CA whose on-disk fingerprint no longer matches the
  // recorded thumbprint — that happens after a regeneration that has not yet
  // re-trusted, and reporting it as trusted would mask the stale state.
  const onDisk = await pemFingerprintDarwin(caPath)
  if (onDisk && normalizeFingerprint(onDisk) !== want) return false
  const { code, stdout } = await run('security', ['verify-cert', '-p', 'ssl', '-c', caPath])
  return code === 0 && /verification successful/i.test(stdout)
}

async function trustLocalCaDarwin(caPem: string): Promise<string | null> {
  const tmpPath = path.join(os.tmpdir(), `openofficellm-ca-${Date.now()}.crt`)
  fs.writeFileSync(tmpPath, caPem, { mode: 0o600 })
  try {
    const fingerprint = await pemFingerprintDarwin(tmpPath)
    if (fingerprint && (await isTrustedDarwin(fingerprint))) return null

    // `-r trustRoot` (not `trustAsRoot`): on macOS 14+/26, `trustAsRoot` writes
    // a trust record that `dump-trust-settings` does not surface and that
    // WKWebView/Word do not honour for a loopback TLS chain — the pane loads
    // as an untrusted certificate. `trustRoot` records an explicit SSL trust
    // policy that does stick and that the system trust evaluator honours.
    // `-p ssl` scopes it to TLS, so the same certificate cannot vouch for code
    // signing or S/MIME.
    const { code, stderr, stdout } = await run('security', [
      'add-trusted-cert',
      '-r',
      'trustRoot',
      '-p',
      'ssl',
      '-k',
      loginKeychain(),
      tmpPath,
    ])
    if (code !== 0) {
      throw new Error(
        `security add-trusted-cert failed (code ${code}): ${stderr.trim() || stdout.trim()}. ` +
          'If you cancelled the authorisation prompt, run --trust-cert again and approve it.',
      )
    }
    return fingerprint
  } finally {
    try {
      fs.unlinkSync(tmpPath)
    } catch {
      // best-effort cleanup
    }
  }
}

/** SHA-1 fingerprint of a PEM on disk, as unseparated uppercase hex. */
async function pemFingerprintDarwin(pemPath: string): Promise<string | null> {
  const { code, stdout } = await run('openssl', [
    'x509',
    '-noout',
    '-fingerprint',
    '-sha1',
    '-in',
    pemPath,
  ])
  if (code !== 0) return null
  const m = /Fingerprint=([0-9A-Fa-f:]+)/.exec(stdout)
  return m ? normalizeFingerprint(m[1]) : null
}

async function removeTrustedCaDarwin(thumbprint: string): Promise<boolean> {
  const want = normalizeFingerprint(thumbprint)
  if (!want) return false
  // `delete-certificate -Z` takes the hash directly, which is the only way to
  // be sure we remove ours and not another certificate sharing the label.
  const { code } = await run('security', ['delete-certificate', '-Z', want, '-t', loginKeychain()])
  if (code !== 0) {
    logger.warn({ msg: 'remove trust failed on darwin', code, thumbprint: want })
    return false
  }
  return true
}
