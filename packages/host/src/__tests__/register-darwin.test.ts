// Darwin add-in registration and autostart, driven through the pure seams.
//
// These run on any platform: the darwin implementation reads its paths from
// helpers that take an explicit home directory, so nothing here touches the
// real ~/Library/Containers or ~/Library/LaunchAgents, and the launchctl
// invocations live in functions the tests never call.

import { describe, it, expect, beforeEach, afterEach, vi, afterAll } from 'vitest'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import {
  wefDirs,
  installAutostartDarwin,
  removeAutostartDarwin,
  readAutostartDarwin,
  unregisterAddinDarwin,
  registerAddinDarwin,
  readRegistrationDarwin,
  setDarwinRunner,
  resetDarwinRunner,
} from '../register.js'
import { logger } from '../logging.js'
import {
  LAUNCH_AGENT_LABEL,
  LAUNCH_AGENT_FILENAME,
  writeLaunchAgentPlist,
  readLaunchAgentCommand,
  removeLaunchAgentPlist,
  launchAgentPlistPath,
} from '../autostart.js'

// The functions under test call os.homedir() to find the "real" home; the
// tests redirect that to a throwaway directory under the system tmpdir, which
// is what keeps them from touching ~/Library/Containers and
// ~/Library/LaunchAgents on the machine running the tests.
const TMP_DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'ool-register-darwin-'))
const HOME = path.join(TMP_DIR, 'home')
fs.mkdirSync(HOME, { recursive: true })

function wefPathFor(app: string): string {
  return path.join(HOME, 'Library', 'Containers', app, 'Data', 'Documents', 'wef')
}

function writeManifestSource(content = '<OfficeApp/>'): string {
  const p = path.join(HOME, 'manifest.xml')
  fs.mkdirSync(path.dirname(p), { recursive: true })
  fs.writeFileSync(p, content)
  return p
}

function makeWord(): void {
  fs.mkdirSync(wefPathFor('com.microsoft.Word'), { recursive: true })
}

function makeExcel(): void {
  fs.mkdirSync(wefPathFor('com.microsoft.Excel'), { recursive: true })
}

function makePlist(command: string): void {
  fs.mkdirSync(path.dirname(launchAgentPlistPath(HOME)), { recursive: true })
  fs.writeFileSync(
    launchAgentPlistPath(HOME),
    `<?xml version="1.0"?><plist>
<dict>
  <key>ProgramArguments</key>
  <array><string>${command}</string></array>
</dict>
</plist>`,
  )
}

// Real launchctl would load and unload a service in the tester's own GUI
// domain — exactly the side effect these tests must not have. The darwin
// autostart functions run child processes through the replaceable runner, so
// the tests stub that instead.
function stubRunner(
  results: Array<[string, { code: number; stderr: string }]>,
): ReturnType<typeof vi.fn> {
  const stub = vi.fn(async (cmd: string, _args: string[]) => {
    const expected = results.find(([want]) => want === cmd)
    // Unstubbed commands "fail" — like launchctl bootout for a service that
    // was never loaded.
    return expected
      ? { code: expected[1].code, stdout: '', stderr: expected[1].stderr }
      : { code: 1, stdout: '', stderr: 'stubbed' }
  })
  setDarwinRunner(stub)
  return stub
}

beforeEach(() => {
  fs.rmSync(TMP_DIR, { recursive: true, force: true })
  fs.mkdirSync(HOME, { recursive: true })
  vi.spyOn(os, 'homedir').mockReturnValue(HOME)
  // The real logger writes to a file under the (mocked) home directory; keep
  // the tests from opening that stream at all.
  vi.spyOn(logger, 'warn').mockImplementation(() => undefined)
  vi.spyOn(logger, 'info').mockImplementation(() => undefined)
  stubRunner([])
})

afterEach(() => {
  resetDarwinRunner()
  vi.restoreAllMocks()
})

afterAll(() => {
  resetDarwinRunner()
  vi.restoreAllMocks()
  try {
    fs.rmSync(TMP_DIR, { recursive: true, force: true })
  } catch {
    // ignore
  }
})

describe('darwin wef registration', () => {
  it('wefDirs points at both Office app sideload folders', () => {
    const dirs = wefDirs(HOME)
    expect(Object.keys(dirs)).toEqual([
      wefPathFor('com.microsoft.Word'),
      wefPathFor('com.microsoft.Excel'),
    ])
    for (const k of Object.keys(dirs)) {
      expect(dirs[k]).toBe(k)
    }
  })

  it('registers into both apps and reports changed', async () => {
    makeWord()
    makeExcel()
    const manifest = writeManifestSource()
    const result = await registerAddinDarwin(manifest)
    expect(result.changed).toBe(true)
    for (const app of ['com.microsoft.Word', 'com.microsoft.Excel']) {
      const target = path.join(wefPathFor(app), 'openofficellm.xml')
      expect(fs.existsSync(target)).toBe(true)
      expect(fs.readFileSync(target, 'utf8')).toBe('<OfficeApp/>')
    }
  })

  it('second register is a no-op', async () => {
    makeWord()
    makeExcel()
    const manifest = writeManifestSource()
    await registerAddinDarwin(manifest)
    expect((await registerAddinDarwin(manifest)).changed).toBe(false)
  })

  it('re-registers when the copy differs from the source', async () => {
    makeWord()
    makeExcel()
    const manifest = writeManifestSource('<OfficeApp/>')
    await registerAddinDarwin(manifest)
    fs.writeFileSync(
      path.join(wefPathFor('com.microsoft.Word'), 'openofficellm.xml'),
      '<OfficeApp v="2"/>',
    )
    const result = await registerAddinDarwin(manifest)
    expect(result.changed).toBe(true)
    expect(
      fs.readFileSync(path.join(wefPathFor('com.microsoft.Word'), 'openofficellm.xml'), 'utf8'),
    ).toBe('<OfficeApp/>')
  })

  it('throws when neither Word nor Excel is installed', async () => {
    const manifest = writeManifestSource()
    await expect(registerAddinDarwin(manifest)).rejects.toThrow(/Neither Word nor Excel/)
  })

  it('logs a warning and registers into the installed app when only one is present', async () => {
    const warn = vi.spyOn(logger, 'warn').mockImplementation(() => undefined)
    makeWord()
    const manifest = writeManifestSource()
    const result = await registerAddinDarwin(manifest)
    expect(result.changed).toBe(true)
    expect(fs.existsSync(path.join(wefPathFor('com.microsoft.Word'), 'openofficellm.xml'))).toBe(
      true,
    )
    expect(warn).toHaveBeenCalledWith(
      expect.objectContaining({ msg: 'Office app not installed, skipping its wef folder' }),
    )
    warn.mockRestore()
  })

  it('throws when the manifest does not exist', async () => {
    makeWord()
    await expect(registerAddinDarwin(path.join(HOME, 'missing.xml'))).rejects.toThrow(
      /Manifest not found/,
    )
  })

  it('readRegistration reports registered entries only for existing copies', async () => {
    makeWord()
    const manifest = writeManifestSource()
    await registerAddinDarwin(manifest)
    const state = await readRegistrationDarwin(manifest)
    expect(state.registered).toBe(true)
    expect(state.entries).toEqual({
      [wefPathFor('com.microsoft.Word')]: wefPathFor('com.microsoft.Word'),
    })
    expect(state.stale).toEqual([])
  })

  it('readRegistration flags stale when the source manifest is gone', async () => {
    makeWord()
    makeExcel()
    const manifest = writeManifestSource()
    await registerAddinDarwin(manifest)
    fs.unlinkSync(manifest)
    const state = await readRegistrationDarwin(manifest)
    expect(state.registered).toBe(false)
    expect(state.stale).toEqual([
      wefPathFor('com.microsoft.Word'),
      wefPathFor('com.microsoft.Excel'),
    ])
  })

  it('readRegistration is empty when nothing is sideloaded', async () => {
    const state = await readRegistrationDarwin()
    expect(state.registered).toBe(false)
    expect(state.entries).toEqual({})
    expect(state.stale).toEqual([])
  })

  it('unregister removes the copies and reports whether anything was removed', async () => {
    makeWord()
    makeExcel()
    const manifest = writeManifestSource()
    await registerAddinDarwin(manifest)
    expect(await unregisterAddinDarwin(manifest)).toBe(true)
    expect(fs.existsSync(path.join(wefPathFor('com.microsoft.Word'), 'openofficellm.xml'))).toBe(
      false,
    )
    expect(await unregisterAddinDarwin(manifest)).toBe(false)
  })

  it('wefDirs without an argument defaults to the real home (os.homedir is mocked in tests)', () => {
    expect(Object.keys(wefDirs())).toEqual([
      path.join(HOME, 'Library', 'Containers', 'com.microsoft.Word', 'Data', 'Documents', 'wef'),
      path.join(HOME, 'Library', 'Containers', 'com.microsoft.Excel', 'Data', 'Documents', 'wef'),
    ])
  })
})

describe('darwin autostart', () => {
  // installAutostartDarwin builds ProgramArguments from execPath/argv — under
  // vitest the entry is the node binary plus the test script, so compute the
  // expected string the same way the implementation does.
  const expectedCommand = [process.execPath, path.resolve(process.argv[1]), '--no-browser'].join(
    ' ',
  )

  it('writeLaunchAgentPlist writes a loadable-looking plist', () => {
    const plistPath = writeLaunchAgentPlist(HOME, ['/usr/bin/fakehost', '--no-browser'])
    expect(plistPath).toBe(launchAgentPlistPath(HOME))
    const xml = fs.readFileSync(plistPath, 'utf8')
    expect(xml).toContain('<key>Label</key>')
    expect(xml).toContain(`<string>${LAUNCH_AGENT_LABEL}</string>`)
    expect(xml).toContain('<key>ProgramArguments</key>')
    expect(xml).toContain('<string>/usr/bin/fakehost</string>')
    expect(xml).toContain('<string>--no-browser</string>')
    expect(xml).toContain('<key>RunAtLoad</key>')
    expect(xml).toContain('<key>KeepAlive</key>')
    expect(xml).toContain('<key>ProcessType</key>')
  })

  it('installAutostartDarwin writes the plist and returns the command', async () => {
    const command = await installAutostartDarwin()
    expect(command).toBe(expectedCommand)
    expect(readLaunchAgentCommand(HOME)).toBe(command)
    const xml = fs.readFileSync(launchAgentPlistPath(HOME), 'utf8')
    expect(xml).toContain(`<string>${LAUNCH_AGENT_LABEL}</string>`)
    expect(xml).toContain('<string>--no-browser</string>')
  })

  it('readLaunchAgentCommand returns the command or null', () => {
    expect(readLaunchAgentCommand(HOME)).toBeNull()
    makePlist('/usr/bin/fakehost')
    expect(readLaunchAgentCommand(HOME)).toBe('/usr/bin/fakehost')
  })

  it('readLaunchAgentCommand unescapes xml entities in the command', () => {
    makePlist('/bin/&quot;quoted&quot; &amp; &lt;angled&gt;')
    expect(readLaunchAgentCommand(HOME)).toBe('/bin/"quoted" & <angled>')
  })

  it('readAutostartDarwin is null before install and the command after', async () => {
    expect(await readAutostartDarwin()).toBeNull()
    await installAutostartDarwin()
    expect(await readAutostartDarwin()).toBe(expectedCommand)
  })

  it('installAutostartDarwin treats an already-loaded service as success', async () => {
    const stub = stubRunner([['launchctl', { code: 5, stderr: 'service already loaded' }]])
    const command = await installAutostartDarwin()
    expect(command).toBe(expectedCommand)
    expect(fs.existsSync(launchAgentPlistPath(HOME))).toBe(true)
    expect(stub).toHaveBeenCalledWith('launchctl', expect.arrayContaining(['bootstrap']))
  })

  it('installAutostartDarwin still returns the command when bootstrap fails', async () => {
    stubRunner([['launchctl', { code: 1, stderr: 'boom' }]])
    const command = await installAutostartDarwin()
    expect(command).toBe(expectedCommand)
    expect(fs.existsSync(launchAgentPlistPath(HOME))).toBe(true)
  })

  it('removeAutostartDarwin removes the plist and tolerates the service not existing', async () => {
    await installAutostartDarwin()
    expect(await removeAutostartDarwin()).toBe(true)
    expect(fs.existsSync(launchAgentPlistPath(HOME))).toBe(false)
    expect(await removeAutostartDarwin()).toBe(false)
  })

  it('removeLaunchAgentPlist reports whether the plist existed', () => {
    expect(removeLaunchAgentPlist(HOME)).toBe(false)
    makePlist('/usr/bin/fakehost')
    expect(removeLaunchAgentPlist(HOME)).toBe(true)
    expect(removeLaunchAgentPlist(HOME)).toBe(false)
  })

  it('LAUNCH_AGENT_FILENAME matches the label', () => {
    expect(LAUNCH_AGENT_FILENAME).toBe(`${LAUNCH_AGENT_LABEL}.plist`)
  })
})
