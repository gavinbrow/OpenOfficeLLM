#!/bin/sh
# Installs OpenOfficeLLM on macOS without tripping Gatekeeper.
#
# Usage:
#   ./scripts/install-macos.sh              # latest release
#   ./scripts/install-macos.sh v0.1.4       # a specific tag
#
# Why this exists, and why it is not just a convenience wrapper around the DMG:
#
# The app is ad-hoc signed (codesign --sign -) and NOT notarized -- see
# build-sea.mjs and build-dmg.mjs. Gatekeeper only evaluates bundles that carry
# the com.apple.quarantine extended attribute, and a browser attaches that
# attribute to anything it downloads. So a DMG opened from Safari or Chrome
# produces a quarantined app, Gatekeeper refuses it, and on macOS 15 (Sequoia)
# the resulting dialog offers only "Done" and "Move to Trash" -- Apple removed
# the old Control-click -> Open bypass, so there is no way through it.
#
# curl does not set that attribute. Downloading the same bytes here means the
# installed bundle is never in Gatekeeper's scope in the first place, and the
# install completes with no prompt at all. This is the supported macOS install
# path until the app is Developer ID-signed and notarized (Docs/TODO.md).
#
# It is also the way out of the dialog LOOP. --install registers a LaunchAgent
# with KeepAlive=true (autostart.ts), so launchd relaunches the host forever.
# When the binary is one Gatekeeper refuses, every relaunch raises the malware
# dialog again -- a blocked install does not fail quietly, it spams. The agent
# is torn down below before anything else happens.

set -e

REPO="gavinbrow/OpenOfficeLLM"
TAG="${1:-}"
APP="/Applications/OpenOfficeLLM.app"
LABEL="com.openofficellm.host"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"

if [ "$(uname -s)" != "Darwin" ]; then
  echo "This installer is macOS-only. On Windows use OpenOfficeLLM-Setup-<version>.exe." >&2
  exit 1
fi

# ── 1. Stop the running host and the agent that keeps resurrecting it ───────
# bootout is the modern spelling; the || true covers "not loaded", which is the
# normal case on a first install and must not abort the script under set -e.
echo "==> stopping any running instance"
launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null || true
rm -f "$PLIST"
pkill -f 'OpenOfficeLLM.app/Contents/MacOS/openofficellm-host' 2>/dev/null || true

# ── 2. Resolve the download URL ─────────────────────────────────────────────
if [ -n "$TAG" ]; then
  API="https://api.github.com/repos/$REPO/releases/tags/$TAG"
else
  API="https://api.github.com/repos/$REPO/releases/latest"
fi

echo "==> resolving release"
URL="$(curl -fsSL "$API" \
  | grep -o 'https://[^"]*macOS\.zip' \
  | head -n 1)"

if [ -z "$URL" ]; then
  echo "Could not find a macOS .zip asset in $API" >&2
  exit 1
fi
echo "    $URL"

# ── 3. Download and unpack ─────────────────────────────────────────────────
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

echo "==> downloading"
curl -fL --progress-bar "$URL" -o "$TMP/OpenOfficeLLM.zip"

# ditto, not unzip: the archive is built with `ditto -c -k --sequesterRsrc`,
# which stashes extended attributes in a parallel __MACOSX tree. Only ditto
# puts them back. unzip drops them, and dropping them corrupts the code
# signature -- which turns a working ad-hoc signature into "the application is
# damaged and can't be opened", a strictly worse failure than the one this
# script exists to avoid.
echo "==> unpacking"
ditto -x -k "$TMP/OpenOfficeLLM.zip" "$TMP/out"

if [ ! -d "$TMP/out/OpenOfficeLLM.app" ]; then
  echo "Archive did not contain OpenOfficeLLM.app" >&2
  exit 1
fi

# ── 4. Install ─────────────────────────────────────────────────────────────
# Replace wholesale rather than syncing into the existing bundle: rsyncing
# contents into a directory leaves that directory's own quarantine attribute
# behind, which is exactly the bug that made in-app updates re-trigger the
# dialog (see MACOS_HELPER in update.ts).
echo "==> installing to $APP"
rm -rf "$APP"
ditto "$TMP/out/OpenOfficeLLM.app" "$APP"

# Belt and braces. curl should not have set it, but a re-run of this script
# over a DMG-installed copy would otherwise inherit it.
xattr -dr com.apple.quarantine "$APP" 2>/dev/null || true

if codesign --verify --deep --strict "$APP" 2>/dev/null; then
  echo "    signature verified"
else
  echo "    warning: signature did not verify; re-signing ad-hoc"
  codesign --force --deep --sign - "$APP"
fi

# ── 5. Launch ──────────────────────────────────────────────────────────────
# The launcher runs --install, which trusts the local CA. macOS asks for Touch
# ID or your password at that point -- that prompt is expected and is not
# Gatekeeper.
echo "==> launching"
open "$APP"

cat <<'EOF'

Installed. macOS will ask for Touch ID or your password once, to trust the
local certificate authority -- that is the CA prompt, not a security warning.

Then, in Word or Excel:
  Insert tab -> Add-ins (or My Add-ins) -> MY ADD-INS -> OpenOfficeLLM

Restart Word/Excel first if either was already open; macOS only re-reads the
sideload folder at launch.
EOF
