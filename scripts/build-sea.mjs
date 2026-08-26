#!/usr/bin/env node
// Builds the host service into a single self-contained executable via Node SEA
// (Single Executable Application). Run after `tsup` has produced dist/index.js.
//
// The SEA process (identical on Windows and macOS):
//   1. node --experimental-sea-config sea-config.json  → sea-prep.blob
//   2. Copy the running node binary → staging file
//   3. postject injects the blob at the NODE_SEA_BLOB sentinel
//
// Platform-specific steps:
//   Windows: strip any existing Authenticode signature (v1 ships unsigned; a
//     stale signature from the Node binary would make Windows report a broken
//     sig), then rename to host.exe.
//   macOS: rename to openofficellm-host (Mach-O, no .exe suffix) and ad-hoc
//     codesign it — see the finalize step for why signing is mandatory.
//
// Result: packages/host/host.exe (Windows) or packages/host/openofficellm-host
// (macOS) — a ~45 MB binary that runs without Node installed. The web bundle
// is NOT inside the binary; the installer stages it alongside it (see
// scripts/build-installer.mjs).

import { execFileSync } from 'node:child_process'
import { copyFileSync, chmodSync, existsSync, renameSync, unlinkSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const hostDir = resolve(here, '..', 'packages', 'host')

const IS_DARWIN = process.platform === 'darwin'

const SEA_CONFIG = join(hostDir, 'sea-config.json')
const DIST_DIR = join(hostDir, 'dist')
const DIST_ENTRY = join(DIST_DIR, 'index.cjs')
const BLOB = join(hostDir, 'sea-prep.blob')
const EXE_OUT = join(hostDir, IS_DARWIN ? 'openofficellm-host' : 'host.exe')
const EXE_STAGING = join(hostDir, IS_DARWIN ? 'openofficellm-host.tmp' : 'host.exe.tmp')

function step(name, fn) {
  console.log(`[build-sea] ${name}`)
  fn()
}

// 1. Verify the tsup build ran.
if (!existsSync(DIST_ENTRY)) {
  console.error(`[build-sea] dist/index.js not found — run "npm run build" first.`)
  process.exit(1)
}

// 2. Generate the SEA blob from the built code.
step('generating SEA blob', () => {
  execFileSync('node', ['--experimental-sea-config', SEA_CONFIG], {
    cwd: hostDir,
    stdio: 'inherit',
  })
  if (!existsSync(BLOB)) {
    console.error('[build-sea] sea-prep.blob was not produced — check sea-config.json')
    process.exit(1)
  }
})

// 3. Copy the Node binary. We stage to a temp name so a failed injection does
//    not leave a half-written host binary that looks finished.
step(`copying node binary → ${EXE_STAGING}`, () => {
  // process.execPath is the running node binary — the exact binary the user's
  // Node 22 install provides. Copying it preserves the platform's own node.
  if (IS_DARWIN && existsSync(EXE_STAGING)) {
    // A leftover staging file from a failed run is read-only (r-xr-xr-x),
    // which copyFileSync cannot overwrite.
    unlinkSync(EXE_STAGING)
  }
  copyFileSync(process.execPath, EXE_STAGING)
  if (IS_DARWIN) {
    // The installed node binary is read-only (r-xr-xr-x); postject needs
    // write access to inject the blob.
    chmodSync(EXE_STAGING, 0o755)
  }
})

// 4. Inject the blob. postject writes the resource at the sentinel marker
//    that Node's SEA loader looks for at startup.
step('injecting blob into host binary', () => {
  // Use the local postject JS entry directly — the .bin shims are shell
  // scripts that don't run under `node` on Windows.
  const postjectCli = resolve(here, '..', 'node_modules', 'postject', 'dist', 'cli.js')
  if (!existsSync(postjectCli)) {
    console.error('[build-sea] postject not found — run "npm install" first')
    process.exit(1)
  }
  // The sentinel fuse is the magic string Node's SEA loader searches for to
  // know a blob is embedded. This is the canonical value from the Node docs.
  const FUSE = 'NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2'
  const postjectArgs = [
    postjectCli,
    EXE_STAGING,
    'NODE_SEA_BLOB',
    BLOB,
    '--sentinel-fuse',
    FUSE,
    '--overwrite',
  ]
  if (IS_DARWIN) {
    // Node's macOS SEA loader looks for the blob in the NODE_SEA segment
    // (see node_sea.cc: options.macho_segment_name = "NODE_SEA"), not
    // postject's default __POSTJECT segment. Without this the loader finds
    // nothing and crashes on a null pointer at startup.
    postjectArgs.push('--macho-segment-name', 'NODE_SEA')
  }
  execFileSync('node', postjectArgs, { stdio: 'inherit' })
})

// 5. (Windows only, best effort) strip the Authenticode signature. The Node
//    binary is not signed, but if it ever is, a stale signature on a modified
//    binary makes Windows report "Digital signature corrupted" — worse than
//    unsigned. signtool is not guaranteed to be present; ignore the error if
//    it isn't. macOS has no Authenticode; the codesign step below handles it.
if (!IS_DARWIN) {
  step('stripping any stale signature', () => {
    try {
      execFileSync('signtool', ['remove', '/s', EXE_STAGING], { stdio: 'ignore' })
      console.log('[build-sea] signature stripped')
    } catch {
      console.log('[build-sea] no signtool or no signature — skipping (this is fine for v1)')
    }
  })
}

// 6. Promote the staging binary to the final name.
step(`finalizing ${EXE_OUT}`, () => {
  if (existsSync(EXE_OUT)) unlinkSync(EXE_OUT)
  renameSync(EXE_STAGING, EXE_OUT)
  // Clean up the blob — it's build intermediate, not a shipped artifact.
  if (existsSync(BLOB)) unlinkSync(BLOB)

  if (IS_DARWIN) {
    // Ad-hoc codesign is mandatory on Apple Silicon: the kernel refuses to
    // execute unsigned arm64 binaries (killed with "Bad CPU type" / killed:
    // 9). `--sign -` requests an ad-hoc signature with no identity, which is
    // enough to run locally. A Developer ID signature + notarization is
    // future work for distribution outside this machine.
    try {
      execFileSync('codesign', ['--force', '--sign', '-', EXE_OUT], { stdio: 'inherit' })
      console.log('[build-sea] ad-hoc codesigned')
    } catch {
      console.warn(
        '[build-sea] WARNING: codesign failed or is unavailable — the binary is unsigned. ' +
          'It may still run on Intel Macs, but Apple Silicon will refuse to execute it. ' +
          'Sign it manually with: codesign --force --sign - ' +
          EXE_OUT,
      )
    }
    // The copied node binary is already executable, but be explicit so the
    // artifact is runnable regardless of how it was produced.
    chmodSync(EXE_OUT, 0o755)
  }

  console.log(`[build-sea] done: ${EXE_OUT}`)
})
