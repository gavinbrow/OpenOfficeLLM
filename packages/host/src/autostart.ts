// macOS logon autostart: the per-user LaunchAgent plist and its file I/O.
//
// Windows keeps logon autostart in the HKCU Run key (see register.ts); macOS
// has no registry, and the launchd-idiomatic per-user equivalent is a plist
// under ~/Library/LaunchAgents registered with the user's GUI domain. This
// module owns only the plist file I/O, parameterised on the home directory so
// the tests can run against a throwaway home. The launchctl bootstrap/bootout
// invocations live in register.ts, which keeps them out of the testable seam.

import fs from 'node:fs'
import path from 'node:path'

export const LAUNCH_AGENT_LABEL = 'com.openofficellm.host'
export const LAUNCH_AGENT_FILENAME = `${LAUNCH_AGENT_LABEL}.plist`

/** ~/Library/LaunchAgents/com.openofficellm.host.plist */
export function launchAgentPlistPath(homeDir: string): string {
  return path.join(homeDir, 'Library', 'LaunchAgents', LAUNCH_AGENT_FILENAME)
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function unescapeXml(value: string): string {
  return value
    .replace(/&quot;/g, '"')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
}

function renderPlist(args: string[]): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>${LAUNCH_AGENT_LABEL}</string>
  <key>ProgramArguments</key>
  <array>
${args.map((a) => `    <string>${escapeXml(a)}</string>`).join('\n')}
  </array>
  <key>RunAtLoad</key>
  <true/>
  <key>KeepAlive</key>
  <true/>
  <key>ProcessType</key>
  <string>Background</string>
</dict>
</plist>
`
}

/**
 * Write the LaunchAgent plist and return its path.
 *
 * `args` is the full argv launchd will exec, verbatim: ProgramArguments is an
 * exec-array, not a shell command line, so quoting a path here would make the
 * quotes literal and break launch — unlike the Windows Run key, which is
 * re-parsed as a command line. The caller builds the array; in dev that is
 * `<node> <script> --no-browser` and packaged it is `<binary> --no-browser`.
 */
export function writeLaunchAgentPlist(homeDir: string, args: string[]): string {
  const plistPath = launchAgentPlistPath(homeDir)
  fs.mkdirSync(path.dirname(plistPath), { recursive: true })
  fs.writeFileSync(plistPath, renderPlist(args), 'utf8')
  return plistPath
}

/**
 * The command a written LaunchAgent would run, or null when no plist exists.
 *
 * Reads the ProgramArguments block and rejoins its <string> elements with
 * spaces. Deliberately minimal — the plist format is fixed by us, so a full
 * plist parser buys nothing, and a malformed or foreign plist yields null
 * rather than a guess.
 */
export function readLaunchAgentCommand(homeDir: string): string | null {
  const plistPath = launchAgentPlistPath(homeDir)
  let xml: string
  try {
    xml = fs.readFileSync(plistPath, 'utf8')
  } catch {
    return null
  }
  // Slice from the ProgramArguments key to its </array> — the plist contains
  // other <string> blocks after it (ProcessType etc.), which would leak into
  // a whole-file scan.
  const block = xml.split('<key>ProgramArguments</key>')[1]?.split('</array>')[0]
  const strings = [...(block?.matchAll(/<string>([\s\S]*?)<\/string>/g) ?? [])].map((m) =>
    unescapeXml(m[1]),
  )
  return strings.length > 0 ? strings.join(' ') : null
}

/** Delete the LaunchAgent plist, reporting whether it existed. */
export function removeLaunchAgentPlist(homeDir: string): boolean {
  const plistPath = launchAgentPlistPath(homeDir)
  if (!fs.existsSync(plistPath)) return false
  try {
    fs.unlinkSync(plistPath)
    return true
  } catch {
    return false
  }
}
