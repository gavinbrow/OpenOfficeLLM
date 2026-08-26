#!/usr/bin/env node
// Builds the macOS app bundle and packages it into a DMG.
//
// Pipeline:
//   1. npm run build          → host dist + add-in dist + shared
//   2. build:sea              → openofficellm-host (Node SEA Mach-O binary)
//   3. Stage everything into installer/staging-macos/app-root/
//   4. Generate AppIcon.icns from the branded sparkle PNGs
//   5. codesign (ad-hoc)      → required for arm64 execution
//   6. hdiutil create         → dist/OpenOfficeLLM-<version>-macOS.dmg
//
// The .app layout mirrors what the host expects at runtime — the SEA binary
// resolves the web bundle at <dir-of-process.execPath>/web (see
// packages/host/src/server.ts findAddinDist), so the bundle sits next to the
// binary inside Contents/MacOS/:
//
//   OpenOfficeLLM.app/
//     Contents/
//       Info.plist
//       PkgInfo
//       MacOS/
//         launcher              ← sh entry point (CFBundleExecutable)
//         openofficellm-host    ← the SEA binary
//         web/                  ← add-in bundle (index.html, commands.html, assets/, icons)
//         version.txt           ← version string for runtime version checks
//       Resources/
//         AppIcon.icns
//
// The DMG pairs the app with an "Applications" symlink so the user gets the
// standard drag-to-Applications install window.
//
// Prerequisites:
//   - macOS (iconutil, codesign, hdiutil are part of the OS)
//   - Node 22 (already installed — you're running this)
//
// The SEA build must run on macOS: build-sea.mjs copies process.execPath, so
// the resulting binary is a Mach-O only when built on a Mac. A host.exe from a
// Windows build cannot be packaged into a DMG.

import { execFileSync, execSync } from 'node:child_process'
import {
  chmodSync,
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
const STAGING = join(root, 'installer', 'staging-macos')
const APP_ROOT = join(STAGING, 'app-root')
const APP_DIR = join(APP_ROOT, 'OpenOfficeLLM.app')
const CONTENTS = join(APP_DIR, 'Contents')
const MACOS_DIR = join(CONTENTS, 'MacOS')
const RESOURCES_DIR = join(CONTENTS, 'Resources')
const DIST_OUT = join(root, 'dist')

function run(cmd, opts = {}) {
  console.log(`  > ${cmd}`)
  execSync(cmd, { stdio: 'inherit', cwd: root, ...opts })
}

function step(name, fn) {
  console.log(`\n[build-dmg] ${name}`)
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
step('building openofficellm-host (Node SEA)', () => {
  // A previous failed run can leave a read-only .tmp behind (the copied node
  // binary is r-xr-xr-x), which makes build-sea's copyFileSync fail with
  // EACCES. Clear both platform variants before rebuilding.
  rmSync(join(HOST_DIR, 'openofficellm-host.tmp'), { force: true })
  rmSync(join(HOST_DIR, 'host.exe.tmp'), { force: true })
  run('npm --workspace @openofficellm/host run build:sea')
  if (!existsSync(join(HOST_DIR, 'openofficellm-host'))) {
    if (existsSync(join(HOST_DIR, 'host.exe'))) {
      console.error(
        '[build-dmg] only host.exe was produced — a Windows binary cannot be packaged into a Mac DMG.',
      )
    }
    console.error('[build-dmg] openofficellm-host was not produced — run the SEA build on macOS.')
    process.exit(1)
  }
})

// ─── 3. Stage the payload ─────────────────────────────────────────────────
step('staging app bundle', () => {
  // Clean any previous staging.
  rmSync(STAGING, { recursive: true, force: true })
  mkdirSync(MACOS_DIR, { recursive: true })
  mkdirSync(RESOURCES_DIR, { recursive: true })

  // openofficellm-host — the SEA binary.
  copyFileSync(join(HOST_DIR, 'openofficellm-host'), join(MACOS_DIR, 'openofficellm-host'))
  console.log('  staged: Contents/MacOS/openofficellm-host')

  // web/ — the entire add-in dist directory, next to the binary so
  // findAddinDist() resolves it at <dir-of-process.execPath>/web.
  if (!existsSync(ADDIN_DIST)) {
    console.error('[build-dmg] packages/addin/dist not found — run "npm run build" first')
    process.exit(1)
  }
  cpSync(ADDIN_DIST, join(MACOS_DIR, 'web'), { recursive: true })
  console.log('  staged: Contents/MacOS/web/')

  // version.txt — for runtime version checks (/api/health reports this).
  writeFileSync(join(MACOS_DIR, 'version.txt'), VERSION, 'utf8')
  console.log(`  staged: Contents/MacOS/version.txt (${VERSION})`)

  // PkgInfo — the classic four-char type/creator signature.
  writeFileSync(join(CONTENTS, 'PkgInfo'), 'APPL????', 'utf8')
  console.log('  staged: Contents/PkgInfo')
})

// ─── 4. Generate AppIcon.icns ─────────────────────────────────────────────
step('generating AppIcon.icns', () => {
  const iconsetTmp = join(STAGING, 'iconset-tmp')
  rmSync(iconsetTmp, { recursive: true, force: true })
  mkdirSync(iconsetTmp, { recursive: true })

  // gen-icons.mjs writes icon-<n>.png for each requested size.
  execFileSync(
    'node',
    [join(here, 'gen-icons.mjs'), '--out', iconsetTmp, '--sizes', '16,32,128,256,512,1024'],
    { stdio: 'inherit' },
  )

  // Assemble the iconset. Each generated PNG feeds one or two slots:
  //   16   → icon_16x16.png
  //   32   → icon_16x16@2x.png  + icon_32x32.png
  //   128  → icon_32x32@2x.png  + icon_128x128.png
  //   256  → icon_128x128@2x.png + icon_256x256.png
  //   512  → icon_256x256@2x.png + icon_512x512.png
  //   1024 → icon_512x512@2x.png
  const iconsetDir = join(STAGING, 'AppIcon.iconset')
  mkdirSync(iconsetDir, { recursive: true })
  const slots = [
    [16, 'icon_16x16.png'],
    [32, 'icon_16x16@2x.png'],
    [32, 'icon_32x32.png'],
    [128, 'icon_32x32@2x.png'],
    [128, 'icon_128x128.png'],
    [256, 'icon_128x128@2x.png'],
    [256, 'icon_256x256.png'],
    [512, 'icon_256x256@2x.png'],
    [512, 'icon_512x512.png'],
    [1024, 'icon_512x512@2x.png'],
  ]
  for (const [size, name] of slots) {
    copyFileSync(join(iconsetTmp, `icon-${size}.png`), join(iconsetDir, name))
  }

  try {
    execFileSync(
      'iconutil',
      ['-c', 'icns', iconsetDir, '-o', join(RESOURCES_DIR, 'AppIcon.icns')],
      {
        stdio: 'inherit',
      },
    )
  } catch (e) {
    if (e.code === 'ENOENT') {
      console.error('[build-dmg] iconutil not found — the DMG build must run on macOS.')
      process.exit(1)
    }
    throw e
  }
  console.log('  staged: Contents/Resources/AppIcon.icns')
})

// ─── 5. Write Info.plist and the launcher ─────────────────────────────────
step('writing Info.plist and launcher', () => {
  const plist = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleIdentifier</key>
  <string>com.openofficellm.app</string>
  <key>CFBundleName</key>
  <string>OpenOfficeLLM</string>
  <key>CFBundleDisplayName</key>
  <string>OpenOfficeLLM</string>
  <key>CFBundleVersion</key>
  <string>${VERSION}</string>
  <key>CFBundleShortVersionString</key>
  <string>${VERSION}</string>
  <key>CFBundleExecutable</key>
  <string>launcher</string>
  <key>CFBundleIconFile</key>
  <string>AppIcon</string>
  <key>CFBundlePackageType</key>
  <string>APPL</string>
  <key>LSMinimumSystemVersion</key>
  <string>12.0</string>
  <key>LSUIElement</key>
  <true/>
</dict>
</plist>
`
  writeFileSync(join(CONTENTS, 'Info.plist'), plist, 'utf8')
  console.log('  staged: Contents/Info.plist')

  // The launcher runs --install on every launch: it is idempotent, but it
  // prompts for keychain trust on first run (Touch ID), which is why it lives
  // here rather than in a postinstall step that would run without a user
  // present. Then it hands off to the host without opening a browser.
  const launcher = `#!/bin/sh
# First-run provisioning: idempotent, safe to run on every launch.
DIR="$(cd "$(dirname "$0")" && pwd)"

# Running from the mounted DMG would register a LaunchAgent pointing at
# /Volumes/OpenOfficeLLM/… which vanishes on eject. Send the user to install
# first instead of leaving a broken autostart behind.
case "$DIR" in
  /Volumes/*)
    osascript -e 'display dialog "Drag OpenOfficeLLM to your Applications folder before opening it." buttons {"OK"} default button "OK" with title "OpenOfficeLLM"'
    exit 0
    ;;
esac

# Clear our own quarantine attribute so the Gatekeeper dialog is genuinely a
# one-time event. Reaching this line means Gatekeeper already let us run --
# either the user removed the attribute by hand or clicked "Open Anyway" -- so
# this cannot suppress the first prompt, and is not trying to. What it prevents
# is the SECOND one: "Open Anyway" pins its approval to the bundle's current
# cdhash, so the next auto-update re-signs the app, invalidates that approval,
# and a still-quarantined bundle gets challenged all over again. Dropping the
# attribute now takes the bundle out of Gatekeeper's scope for good.
xattr -dr com.apple.quarantine "$DIR/../.." 2>/dev/null || true

"$DIR/openofficellm-host" --install
exec "$DIR/openofficellm-host" --no-browser
`
  const launcherPath = join(MACOS_DIR, 'launcher')
  writeFileSync(launcherPath, launcher, 'utf8')
  chmodSync(launcherPath, 0o755)
  console.log('  staged: Contents/MacOS/launcher (0755)')
})

// ─── 6. Create the DMG (copies + codesigns the app) ──────────────────────
step('creating DMG', () => {
  mkdirSync(DIST_OUT, { recursive: true })
  // hdiutil -srcfolder copies the folder's CONTENTS, so a source folder that
  // pairs the app with an "Applications" symlink to /Applications gives the
  // user the standard drag-to-Applications window.
  const dmgSrc = join(STAGING, 'dmg-src')
  rmSync(dmgSrc, { recursive: true, force: true })
  // Pair the app with an "Applications" symlink to /Applications so the user
  // gets the standard drag-to-Applications install window.
  cpSync(APP_DIR, join(dmgSrc, 'OpenOfficeLLM.app'), { recursive: true })
  execFileSync('ln', ['-s', '/Applications', join(dmgSrc, 'Applications')])

  // Codesign AFTER cpSync — copying with cpSync recreates files with new
  // inodes, which invalidates the code directory hashes. Signing here, on the
  // copy that hdiutil reads, leaves APP_DIR unsigned but the DMG contents
  // correctly signed.
  try {
    execFileSync(
      'codesign',
      ['--force', '--deep', '--sign', '-', join(dmgSrc, 'OpenOfficeLLM.app')],
      { stdio: 'inherit' },
    )
  } catch (e) {
    if (e.code === 'ENOENT') {
      console.error('[build-dmg] codesign not found — the DMG build must run on macOS.')
      process.exit(1)
    }
    throw e
  }
  console.log('  codesigned dmg-src/OpenOfficeLLM.app')

  const dmg = join(DIST_OUT, `OpenOfficeLLM-${VERSION}-macOS.dmg`)

  // Do NOT let hdiutil size the image itself. A bare
  // `create -srcfolder ... -format UDZO` derives the volume size from the
  // source and gets it wrong: on the CI runner it created and mounted the
  // volume, then failed partway through copying files in with
  //
  //   hdiutil: create failed - No space left on device
  //   could not access /Volumes/OpenOfficeLLM/OpenOfficeLLM.app/Contents/MacOS/openofficellm-host
  //
  // It ran out of room on the largest file in the bundle, the ~100 MB SEA
  // binary. The estimate leaves nothing for HFS+ metadata or the extended
  // attributes codesign writes, and the runner's own free disk was never the
  // problem — staging and codesigning the same bytes had just succeeded.
  //
  // So: measure the source, add real headroom, create a read-write image at
  // that explicit size, then compress it. Same two-step create-dmg uses. The
  // slack costs nothing in the shipped artifact — UDZO compresses untouched
  // free blocks down to almost nothing.
  const srcMb = Math.ceil(
    Number.parseInt(
      execFileSync('du', ['-sk', dmgSrc], { encoding: 'utf8' }).trim().split(/\s+/)[0],
      10,
    ) / 1024,
  )
  const sizeMb = srcMb + 200
  const rwDmg = join(STAGING, 'OpenOfficeLLM-rw.dmg')
  rmSync(rwDmg, { force: true })

  try {
    execFileSync(
      'hdiutil',
      [
        'create',
        '-volname',
        'OpenOfficeLLM',
        '-srcfolder',
        dmgSrc,
        '-fs',
        'HFS+',
        '-format',
        'UDRW',
        '-size',
        `${sizeMb}m`,
        '-ov',
        rwDmg,
      ],
      { stdio: 'inherit' },
    )
  } catch (e) {
    if (e.code === 'ENOENT') {
      console.error('[build-dmg] hdiutil not found — the DMG build must run on macOS.')
      process.exit(1)
    }
    throw e
  }
  console.log(`  staged read-write image: ${sizeMb} MB for ${srcMb} MB of content`)

  execFileSync(
    'hdiutil',
    ['convert', rwDmg, '-format', 'UDZO', '-imagekey', 'zlib-level=9', '-ov', '-o', dmg],
    { stdio: 'inherit' },
  )
  rmSync(rwDmg, { force: true })

  if (existsSync(dmg)) {
    console.log(`\n[build-dmg] done: ${dmg}`)
  } else {
    console.warn(`\n[build-dmg] hdiutil ran but ${dmg} was not found — check hdiutil output`)
  }
})

// ─── 7. Create the update zip ─────────────────────────────────────────────
// Sign a second copy for the zip — the zip is used by the auto-updater to
// replace the app in place. The ditto archive preserves signatures, so the
// signed app in /Applications stays signed after the helper script swaps it.
step('creating update zip', () => {
  const zipSrc = join(STAGING, 'zip-src')
  rmSync(zipSrc, { recursive: true, force: true })
  cpSync(APP_DIR, join(zipSrc, 'OpenOfficeLLM.app'), { recursive: true })
    execFileSync(
      'codesign',
      ['--force', '--deep', '--sign', '-', join(zipSrc, 'OpenOfficeLLM.app')],
      { stdio: 'inherit' },
    )
  const zipPath = join(DIST_OUT, `OpenOfficeLLM-${VERSION}-macOS.zip`)
  execFileSync(
    'ditto',
    ['-c', '-k', '--sequesterRsrc', '--keepParent', join(zipSrc, 'OpenOfficeLLM.app'), zipPath],
    {
      stdio: 'inherit',
    },
  )
  console.log(`  created: ${zipPath}`)
})
