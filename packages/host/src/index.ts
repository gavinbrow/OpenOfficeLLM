// OpenOfficeLLM local host service entrypoint.
// Serves the React task pane over loopback HTTPS and brokers every model call.
import { parseArgs } from 'node:util'
import process from 'node:process'
import { startServer } from './server.js'
import { logger, initLogger, type LogLevel } from './logging.js'
import { acquireLock, releaseLock } from './lock.js'
import { trustLocalCa, removeTrustedCa } from './trust.js'
import { loadCertMaterial, getCaThumbprint } from './tls.js'
import { writeManifest, manifestPath } from './manifest.js'
import {
  registerAddin,
  unregisterAddin,
  installAutostart,
  removeAutostart,
  readRegistration,
  readAutostart,
} from './register.js'
import { loadConfig, updateSettings } from './config.js'
import { addPairing, removePairing, extensionOrigin } from './pairing.js'
import { DEFAULT_PORT } from './paths.js'

const IS_DARWIN = process.platform === 'darwin'

export interface CliOptions {
  port?: number
  config?: string
  noBrowser: boolean
  verbose: boolean
  repair: boolean
  diagnose: boolean
  trustCert: boolean
  removeCert: boolean
  install: boolean
  uninstall: boolean
  noAutostart: boolean
  /** Extension id to allow (`--pair <id>`). */
  pair?: string
  /** Extension id to revoke (`--unpair <id>`). */
  unpair?: string
  /** List paired extensions and exit. */
  listPairs: boolean
}

export function parseCliArgs(argv: string[]): CliOptions {
  const { values } = parseArgs({
    args: argv,
    options: {
      port: { type: 'string', short: 'p' },
      config: { type: 'string', short: 'c' },
      'no-browser': { type: 'boolean', default: false },
      verbose: { type: 'boolean', default: false },
      repair: { type: 'boolean', default: false },
      diagnose: { type: 'boolean', default: false },
      'trust-cert': { type: 'boolean', default: false },
      'remove-cert': { type: 'boolean', default: false },
      install: { type: 'boolean', default: false },
      uninstall: { type: 'boolean', default: false },
      'no-autostart': { type: 'boolean', default: false },
      pair: { type: 'string' },
      unpair: { type: 'string' },
      'list-pairs': { type: 'boolean', default: false },
    },
    allowPositionals: false,
  })
  let port: number | undefined
  if (values.port !== undefined) {
    const n = Number(values.port)
    if (!Number.isInteger(n) || n < 1 || n > 65535) {
      throw new Error(`--port must be an integer in 1..65535, got "${values.port}"`)
    }
    port = n
  }
  return {
    port,
    config: values.config,
    noBrowser: values['no-browser'] === true,
    verbose: values.verbose === true,
    repair: values.repair === true,
    diagnose: values.diagnose === true,
    trustCert: values['trust-cert'] === true,
    // Must mirror trustCert exactly. The previous form was
    // `values['remove-cert'] === false ? false : true`, which is only correct
    // while the parseArgs default stays `false` — drop that default and an
    // undefined value flips this to true, so a bare `host.exe` would delete the
    // user's CA and exit instead of starting the server.
    removeCert: values['remove-cert'] === true,
    install: values.install === true,
    uninstall: values.uninstall === true,
    noAutostart: values['no-autostart'] === true,
    pair: values.pair,
    unpair: values.unpair,
    listPairs: values['list-pairs'] === true,
  }
}

/**
 * Manage the browser-extension allowlist.
 *
 * Pairing is a trust decision — a paired origin may read the auth token — so
 * it is a deliberate command the user runs, never something inferred from a
 * request. Chrome shows an extension's id on chrome://extensions with
 * Developer mode on; that string is the whole input.
 */
function runPairing(opts: CliOptions): number {
  const settings = loadConfig().settings

  if (opts.pair !== undefined) {
    const { origins, added, error } = addPairing(settings.pairedExtensions, opts.pair)
    if (error) {
      console.error(error)
      return 1
    }
    if (!added) {
      console.log(`Already paired: ${extensionOrigin(opts.pair)}`)
      return 0
    }
    updateSettings({ ...settings, pairedExtensions: origins })
    console.log(`Paired ${extensionOrigin(opts.pair)}.`)
    console.log('Restart is not required — the allowlist is read per request.')
    return 0
  }

  if (opts.unpair !== undefined) {
    const { origins, removed } = removePairing(settings.pairedExtensions, opts.unpair)
    if (!removed) {
      console.error(`Not paired: ${opts.unpair}`)
      return 1
    }
    updateSettings({ ...settings, pairedExtensions: origins })
    console.log(`Unpaired ${opts.unpair}. It can no longer read the auth token.`)
    return 0
  }

  if (settings.pairedExtensions.length === 0) {
    console.log('No extensions paired. Pair one with:  --pair <extension-id>')
  } else {
    console.log('Paired extensions:')
    for (const o of settings.pairedExtensions) console.log(`  ${o}`)
  }
  return 0
}

async function runCertAction(opts: CliOptions): Promise<number> {
  if (opts.removeCert) {
    const thumbprint = getCaThumbprint()
    if (!thumbprint) {
      console.error('No CA thumbprint found — nothing to remove.')
      return 1
    }
    await removeTrustedCa(thumbprint)
    console.log(
      IS_DARWIN
        ? `Removed CA with thumbprint ${thumbprint} from the login keychain (if present).`
        : `Removed CA with thumbprint ${thumbprint} from CurrentUser\\Root (if present).`,
    )
    return 0
  }
  if (opts.trustCert) {
    const { caPem } = loadCertMaterial()
    const thumbprint = await trustLocalCa(caPem)
    if (thumbprint === null) {
      console.log('CA already trusted — nothing to do.')
    } else {
      console.log(`Trusted CA with thumbprint ${thumbprint}.`)
      console.log(
        IS_DARWIN
          ? 'macOS may have shown a Touch ID / password prompt — this is expected.'
          : 'Windows may have shown a Security Warning dialog — this is expected.',
      )
    }
    return 0
  }
  return -1
}

/**
 * One-shot provisioning: trust the CA, write the manifest, register it with
 * Office, and start at logon. Everything here is idempotent, so running it
 * again after an upgrade or a port change is the supported repair path.
 */
async function runInstall(opts: CliOptions): Promise<number> {
  const port = opts.port ?? loadConfig().port ?? DEFAULT_PORT

  console.log('Installing OpenOfficeLLM for the current user (no admin required).\n')

  // Certificate first. This is the only step that can show UI, and it is the
  // only one the user has to react to — get it out of the way before anything
  // else has a chance to fail.
  console.log('[1/4] Trusting the local certificate authority...')
  if (IS_DARWIN) {
    console.log('      macOS will show a Touch ID / password prompt asking you to trust the')
    console.log('      local certificate authority. If prompted, approve it — the CA is')
    console.log('      stored in the login keychain.')
  } else {
    console.log('      Windows will show a "Security Warning" dialog asking you to confirm.')
    console.log('      This is expected: Windows always asks before adding a root certificate,')
    console.log('      even for a per-user, self-generated one. Choose Yes to continue.')
  }
  const { caPem } = loadCertMaterial()
  const thumbprint = await trustLocalCa(caPem)
  console.log(
    thumbprint === null
      ? '      Already trusted — no dialog needed.'
      : `      Trusted (thumbprint ${thumbprint}).`,
  )

  console.log(`[2/4] Writing the add-in manifest for port ${port}...`)
  const written = writeManifest({ port })
  console.log(`      ${written.path}`)

  console.log('[3/4] Registering the add-in with Office...')
  if (IS_DARWIN) {
    console.log('      Copying the add-in manifest to the Word and Excel sideload folders...')
  }
  await registerAddin(written.path)
  console.log(
    IS_DARWIN
      ? '      Copied to Word and Excel wef folders.'
      : '      Registered under HKCU\\...\\Office\\16.0\\WEF\\Developer.',
  )

  if (opts.noAutostart) {
    console.log('[4/4] Skipping autostart (--no-autostart).')
  } else {
    console.log('[4/4] Enabling start at logon...')
    const command = await installAutostart()
    console.log(`      ${command}`)
  }

  console.log('\nDone. Start the host, then in Word or Excel:')
  printInsertSteps()
  if (IS_DARWIN) {
    console.log('\nIf Word or Excel was open during install, restart it — macOS re-reads the')
    console.log('wef folder when the app launches.')
  }
  console.log('\nThe host must be running whenever you use the add-in — Office loads the')
  console.log('pane from it. Autostart handles that after the next logon.')
  return 0
}

/**
 * How to actually open the add-in.
 *
 * Worth spelling out every time. A manifest registered under WEF\Developer does
 * not put anything on the ribbon until it has been inserted once in that
 * session, so telling the user to "look for the OpenOfficeLLM group on the Home
 * tab" sends them hunting for something that is not there yet — in Excel just
 * as much as in Word.
 */
function printInsertSteps(): void {
  if (IS_DARWIN) {
    console.log('  1. Restart Word or Excel (if it is open)')
    console.log('  2. Insert tab -> Add-ins (or My Add-ins)')
    console.log('  3. Under "MY ADD-INS", click OpenOfficeLLM')
    console.log('\n  The add-in is available in both Word and Excel.')
    return
  }
  console.log('  1. Home tab -> Add-ins')
  console.log('  2. Under "Developer Add-ins", click OpenOfficeLLM')
  console.log('  3. The OpenOfficeLLM group and its "AI Assistant" button appear on Home')
  console.log('\n  Repeat once per Word/Excel session: sideloaded add-ins are session-scoped')
  console.log('  and Office does not restore them on restart. See Docs/SPIKE-LNA.md.')
}

async function runUninstall(): Promise<number> {
  console.log('Removing OpenOfficeLLM registration for the current user.\n')
  const target = manifestPath()

  const unregistered = await unregisterAddin(target)
  console.log(unregistered ? 'Add-in unregistered from Office.' : 'Add-in was not registered.')

  const autostartRemoved = await removeAutostart()
  console.log(autostartRemoved ? 'Autostart removed.' : 'Autostart was not set.')

  const thumbprint = getCaThumbprint()
  if (thumbprint) {
    // Match on thumbprint, never subject name — a subject-name match would
    // happily delete a different install's CA, or anything else that picked
    // the same common name.
    const removed = await removeTrustedCa(thumbprint)
    console.log(
      removed
        ? IS_DARWIN
          ? `CA ${thumbprint} removed from the login keychain.`
          : `CA ${thumbprint} removed from CurrentUser\\Root.`
        : 'CA was not trusted.',
    )
  }

  console.log(
    IS_DARWIN
      ? '\nConfiguration, secrets, and chat history under ~/Library/Application Support/OpenOfficeLLM'
      : '\nConfiguration, secrets, and chat history under %APPDATA%\\OpenOfficeLLM',
  )
  console.log('were left in place. Delete that folder to remove them.')
  return 0
}

async function runDiagnose(opts: CliOptions): Promise<number> {
  const cfg = loadConfig()
  const port = opts.port ?? cfg.port ?? DEFAULT_PORT
  const target = manifestPath()

  console.log('OpenOfficeLLM diagnostics\n')
  console.log(`Configured port:  ${port}`)
  console.log(`Manifest:         ${target}`)

  const thumbprint = getCaThumbprint()
  console.log(`CA thumbprint:    ${thumbprint ?? '(none generated yet)'}`)
  if (thumbprint) {
    const { isTrusted } = await import('./trust.js')
    console.log(`CA trusted:       ${(await isTrusted(thumbprint)) ? 'yes' : 'NO — run --install'}`)
  }

  const reg = await readRegistration(target)
  console.log(`Add-in registered: ${reg.registered ? 'yes' : 'NO — run --install'}`)
  for (const [k, v] of Object.entries(reg.entries)) {
    const mine = k.toLowerCase() === target.toLowerCase()
    console.log(`  ${mine ? '*' : ' '} ${k}${v !== k ? ` -> ${v}` : ''}`)
  }
  if (reg.stale.length > 0) {
    console.log(
      `  ${reg.stale.length} stale entr${reg.stale.length === 1 ? 'y' : 'ies'} (missing file) — run --install to clean up`,
    )
  }

  const autostart = await readAutostart()
  console.log(`Autostart:        ${autostart ?? 'not set'}`)

  const health = await probeHealth(port)
  console.log(`Host service:     ${health}`)

  if (!health.startsWith('responding')) {
    console.log('\nThe pane cannot load while the host is down — Office serves it from this')
    console.log('service, so Word/Excel will show "Sorry, we can\'t load the add-in".')
    console.log(
      IS_DARWIN
        ? 'Start it with: npm start (or launch the OpenOfficeLLM app)'
        : 'Start it with: npm start',
    )
  } else if (reg.registered) {
    console.log('\nTo open the add-in:')
    printInsertSteps()
  }
  return 0
}

/**
 * Probe /api/health over TLS, trusting our own CA.
 *
 * Deliberately not global `fetch`: undici validates against Node's bundled
 * root list and ignores the Windows store, so fetch reports a perfectly
 * healthy service as unreachable. A diagnostic that cries wolf is worse than
 * no diagnostic — this is the tool someone reaches for when the add-in is
 * already misbehaving.
 */
async function probeHealth(port: number): Promise<string> {
  const https = await import('node:https')
  let ca: string
  try {
    ca = loadCertMaterial().caPem
  } catch {
    return 'unknown (no certificate material yet — run --install)'
  }
  return new Promise<string>((resolve) => {
    const req = https.get(
      { host: '127.0.0.1', port, path: '/api/health', ca, rejectUnauthorized: true, timeout: 3000 },
      (res) => {
        const chunks: Buffer[] = []
        res.on('data', (c: Buffer) => chunks.push(c))
        res.on('end', () => {
          const body = Buffer.concat(chunks).toString('utf8')
          resolve(`responding (${res.statusCode}) ${body}`)
        })
      },
    )
    req.on('timeout', () => {
      req.destroy()
      resolve('timed out after 3s')
    })
    req.on('error', (e) => resolve(`not responding (${e.message})`))
  })
}

async function main(): Promise<void> {
  let opts: CliOptions
  try {
    opts = parseCliArgs(process.argv.slice(2))
  } catch (e) {
    console.error(String((e as Error).message ?? e))
    process.exit(2)
  }

  const level: LogLevel = opts.verbose ? 'debug' : 'info'
  initLogger({ level, verboseBodies: opts.verbose })

  // --repair is --install: every provisioning step is idempotent, so
  // re-running it is exactly how you recover from a moved port, a cleared
  // certificate store, or a lost registry key.
  if (opts.install || opts.repair) {
    try {
      process.exit(await runInstall(opts))
    } catch (e) {
      console.error(`Install failed: ${String((e as Error).message ?? e)}`)
      process.exit(1)
    }
  }
  if (opts.uninstall) {
    try {
      process.exit(await runUninstall())
    } catch (e) {
      console.error(`Uninstall failed: ${String((e as Error).message ?? e)}`)
      process.exit(1)
    }
  }
  if (opts.pair !== undefined || opts.unpair !== undefined || opts.listPairs) {
    process.exit(runPairing(opts))
  }

  if (opts.diagnose) {
    process.exit(await runDiagnose(opts))
  }

  if (opts.trustCert || opts.removeCert) {
    const code = await runCertAction(opts)
    process.exit(code)
  }

  if (!acquireLock()) {
    logger.error({
      msg: 'Another host instance is already running. Delete host.lock if stale.',
    })
    process.exit(1)
  }

  let shutdownStarted = false
  const shutdown = async (signal: string) => {
    if (shutdownStarted) return
    shutdownStarted = true
    logger.info({ msg: 'shutting down', signal })
    try {
      await stopServer()
    } catch (e) {
      logger.error({ msg: 'error during shutdown', error: String(e) })
    }
    releaseLock()
    process.exit(0)
  }

  const stopServer = await startServer(opts, {
    onShutdown: () => releaseLock(),
  })

  process.on('SIGINT', () => void shutdown('SIGINT'))
  process.on('SIGTERM', () => void shutdown('SIGTERM'))
  process.on('SIGHUP', () => void shutdown('SIGHUP'))
}

main().catch((e) => {
  console.error('Fatal:', e)
  process.exit(1)
})
