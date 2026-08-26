#!/usr/bin/env node
// Builds the full installer payload and compiles it with Inno Setup.
//
// Pipeline:
//   1. npm run build          → host dist + add-in dist + shared
//   2. build:sea              → host.exe (Node SEA binary)
//   3. Stage everything into installer/staging/
//   4. ISCC installer.iss     → dist/OpenOfficeLLM-Setup-<version>.exe
//
// The staging directory mirrors what the installer copies to the user's machine:
//
//   installer/staging/
//     host.exe               ← the SEA binary
//     web/                   ← add-in bundle (index.html, commands.html, assets/, icons)
//     version.txt            ← version string for runtime version checks
//
// Prerequisites:
//   - Node 22 (already installed — you're running this)
//   - Inno Setup 6 (ISCC.exe on PATH, or at the default install location)
//
// Secrets: DPAPI is invoked at runtime via Windows PowerShell (.NET
// System.Security.Cryptography.ProtectedData), so no native addon is built or
// staged here. If DPAPI is unavailable on the host, secrets fall back to
// AES-256-GCM.

import { execSync, execFileSync } from 'node:child_process'
import {
  copyFileSync,
  cpSync,
  existsSync,
  mkdirSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const root = resolve(here, '..')

const HOST_DIR = join(root, 'packages', 'host')
const ADDIN_DIST = join(root, 'packages', 'addin', 'dist')
const STAGING = join(root, 'installer', 'staging')
const INSTALLER_ISS = join(root, 'installer', 'installer.iss')
const DIST_OUT = join(root, 'dist')

function run(cmd, opts = {}) {
  console.log(`  > ${cmd}`)
  execSync(cmd, { stdio: 'inherit', cwd: root, ...opts })
}

function step(name, fn) {
  console.log(`\n[build-installer] ${name}`)
  fn()
}

// Read the version from the root package.json — the single source of truth.
const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'))
const VERSION = pkg.version

// ─── 1. Build the JS ───────────────────────────────────────────────────────
step('building host + add-in + shared', () => {
  run('npm run build')
})

// ─── 2. Build the SEA binary ──────────────────────────────────────────────
step('building host.exe (Node SEA)', () => {
  run('npm --workspace @openofficellm/host run build:sea')
  if (!existsSync(join(HOST_DIR, 'host.exe'))) {
    console.error('[build-installer] host.exe was not produced — check build-sea.mjs output')
    process.exit(1)
  }
})

// ─── 3. Stage the payload ─────────────────────────────────────────────────
step('staging installer payload', () => {
  // Clean any previous staging.
  rmSync(STAGING, { recursive: true, force: true })
  mkdirSync(STAGING, { recursive: true })

  // host.exe
  copyFileSync(join(HOST_DIR, 'host.exe'), join(STAGING, 'host.exe'))
  console.log('  staged: host.exe')

  // web/ — the entire add-in dist directory
  if (!existsSync(ADDIN_DIST)) {
    console.error('[build-installer] packages/addin/dist not found — run "npm run build" first')
    process.exit(1)
  }
  cpSync(ADDIN_DIST, join(STAGING, 'web'), { recursive: true })
  console.log('  staged: web/')

  // version.txt — for runtime version checks (/api/health reports this)
  writeFileSync(join(STAGING, 'version.txt'), VERSION, 'utf8')
  console.log(`  staged: version.txt (${VERSION})`)
})

// ─── 4. Compile with Inno Setup ───────────────────────────────────────────
step('compiling installer with Inno Setup', () => {
  // Find ISCC.exe — on PATH, or at the default install locations.
  // ISCC exits non-zero from /? (it prints help to stderr and returns 1),
  // so we accept any output as "found" rather than requiring exit 0.
  let iscc = null
  const candidates = [
    'ISCC.exe',
    join('C:', 'Program Files (x86)', 'Inno Setup 6', 'ISCC.exe'),
    join('C:', 'Program Files', 'Inno Setup 6', 'ISCC.exe'),
    join('C:', 'Program Files', 'Inno Setup 7', 'ISCC.exe'),
  ]
  for (const c of candidates) {
    try {
      execFileSync(c, ['/?'], { stdio: 'pipe' })
      iscc = c
      break
    } catch (e) {
      // ISCC returns exit code 1 from /? but still produces output — that
      // means the binary exists and runs. Any ENOENT means it's not there.
      if (e.code !== 'ENOENT' && e.status !== undefined) {
        iscc = c
        break
      }
    }
  }
  if (!iscc) {
    console.error('')
    console.error('[build-installer] Inno Setup (ISCC.exe) not found.')
    console.error('  Install Inno Setup 6 from https://jrsoftware.org/isdl.php')
    console.error('  The staging payload is ready at: installer/staging/')
    console.error('  Run ISCC manually:  ISCC installer/installer.iss')
    process.exit(1)
  }

  mkdirSync(DIST_OUT, { recursive: true })
  // ISCC reads AppVersion from the .iss file. We patch it via a /D define so
  // the version stays in sync with package.json without editing the .iss.
  execFileSync(iscc, [`/DAPP_VERSION=${VERSION}`, INSTALLER_ISS], { stdio: 'inherit' })

  const expected = join(DIST_OUT, `OpenOfficeLLM-Setup-${VERSION}.exe`)
  if (existsSync(expected)) {
    console.log(`\n[build-installer] done: ${expected}`)
  } else {
    console.warn(`\n[build-installer] ISCC ran but ${expected} was not found — check ISCC output`)
  }
})

// ─── 5. Create the update zip ─────────────────────────────────────────────
step('creating update zip', () => {
  const zipPath = join(DIST_OUT, `OpenOfficeLLM-${VERSION}-win.zip`)
  execFileSync('tar', ['-a', '-cf', zipPath, '-C', STAGING, '.'], { stdio: 'inherit' })
  console.log(`  created: ${zipPath}`)
})
