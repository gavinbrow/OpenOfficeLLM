#!/usr/bin/env node
// Builds the host service into a single self-contained executable via Node SEA
// (Single Executable Application). Run after `tsup` has produced dist/index.js.
//
// The SEA process:
//   1. node --experimental-sea-config sea-config.json  → sea-prep.blob
//   2. Copy node.exe → host.exe
//   3. postject injects the blob into host.exe at the NODE_SEA_BLOB sentinel
//   4. Strip any existing Authenticode signature (v1 ships unsigned; a stale
//      signature from the Node binary would make Windows report a broken sig)
//
// Result: packages/host/host.exe — a ~45 MB binary that runs without Node
// installed. The web bundle and win-dpapi.node are NOT inside the exe; the
// installer stages them alongside it (see scripts/build-installer.mjs).

import { execFileSync } from 'node:child_process'
import { copyFileSync, existsSync, renameSync, unlinkSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const hostDir = resolve(here, '..', 'packages', 'host')

const SEA_CONFIG = join(hostDir, 'sea-config.json')
const DIST_DIR = join(hostDir, 'dist')
const DIST_ENTRY = join(DIST_DIR, 'index.cjs')
const BLOB = join(hostDir, 'sea-prep.blob')
const EXE_OUT = join(hostDir, 'host.exe')
const EXE_STAGING = join(hostDir, 'host.exe.tmp')

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
//    not leave a half-written host.exe that looks finished.
step('copying node.exe → host.exe', () => {
  // process.execPath is the running node.exe — the exact binary the user's
  // Node 22 install provides. Copying it preserves the platform's own node.
  copyFileSync(process.execPath, EXE_STAGING)
})

// 4. Inject the blob. postject writes the resource at the sentinel marker
//    that Node's SEA loader looks for at startup.
step('injecting blob into host.exe', () => {
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
  execFileSync(
    'node',
    [postjectCli, EXE_STAGING, 'NODE_SEA_BLOB', BLOB, '--sentinel-fuse', FUSE, '--overwrite'],
    { stdio: 'inherit' },
  )
})

// 5. (Best effort) strip the Authenticode signature. The Node binary is not
//    signed, but if it ever is, a stale signature on a modified binary makes
//    Windows report "Digital signature corrupted" — worse than unsigned.
//    signtool is not guaranteed to be present; ignore the error if it isn't.
step('stripping any stale signature', () => {
  try {
    execFileSync('signtool', ['remove', '/s', EXE_STAGING], { stdio: 'ignore' })
    console.log('[build-sea] signature stripped')
  } catch {
    console.log('[build-sea] no signtool or no signature — skipping (this is fine for v1)')
  }
})

// 6. Promote the staging exe to the final name.
step('finalizing host.exe', () => {
  if (existsSync(EXE_OUT)) unlinkSync(EXE_OUT)
  renameSync(EXE_STAGING, EXE_OUT)
  // Clean up the blob — it's build intermediate, not a shipped artifact.
  if (existsSync(BLOB)) unlinkSync(BLOB)
  console.log(`[build-sea] done: ${EXE_OUT}`)
})
