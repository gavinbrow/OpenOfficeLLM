// Finding the local host service and obtaining its bearer token.
//
// The task pane gets both for free: it is served by the host, so it is
// same-origin and the token arrives in a <meta> tag. The extension has neither
// advantage, so it has to (a) find which port the host bound and (b) ask for a
// token at an endpoint that only answers to paired extensions.
//
// Pairing is one command the user runs once:
//
//     openofficellm --pair <extension-id>
//
// Until then every request is refused, which is the correct default: the token
// is a key to every configured model provider, and handing it to whichever
// extension asked first would make the allowlist decorative.

import { configureApi } from '@openofficellm/ui'

/** The default port, and the small range the host walks when it is taken.
 *  Kept short deliberately — this is a scan of the user's own loopback
 *  interface, and a wide one is both slow and rude. */
const DEFAULT_PORT = 7317
const PORT_SPAN = 10

/** Chrome extension origins have no port, so one host permission entry covers
 *  every port the service might land on. */
export const EXTENSION_ORIGIN = `chrome-extension://${chrome.runtime.id}`

export interface PairingSuccess {
  ok: true
  baseUrl: string
  version: string
}

export interface PairingFailure {
  ok: false
  /** Distinguishes "the host is not running" from "the host is running and
   *  refused us", because the two need completely different fixes and a single
   *  "could not connect" would send the user to the wrong one. */
  reason: 'unreachable' | 'not_paired' | 'error'
  message: string
}

export type PairingResult = PairingSuccess | PairingFailure

const PORT_KEY = 'host.port'

/** Last known good port, so the common case is one request rather than ten. */
async function rememberedPort(): Promise<number | undefined> {
  try {
    const v = await chrome.storage.local.get(PORT_KEY)
    const p = v[PORT_KEY]
    return typeof p === 'number' ? p : undefined
  } catch {
    return undefined
  }
}

async function rememberPort(port: number): Promise<void> {
  try {
    await chrome.storage.local.set({ [PORT_KEY]: port })
  } catch {
    // Storage being unavailable costs a rescan next time, nothing more.
  }
}

function candidatePorts(remembered: number | undefined): number[] {
  const span = Array.from({ length: PORT_SPAN }, (_, i) => DEFAULT_PORT + i)
  if (remembered === undefined) return span
  return [remembered, ...span.filter((p) => p !== remembered)]
}

function baseUrlFor(port: number): string {
  return `https://127.0.0.1:${port}`
}

/** Headers that identify this extension to the host. See the note on
 *  `classifyOrigin` in the host: Chrome does not reliably attach an Origin to
 *  a privileged extension fetch, so the extension names itself explicitly. */
function idHeaders(): Record<string, string> {
  return { 'X-OpenOfficeLLM-Extension': EXTENSION_ORIGIN }
}

async function probe(port: number, signal: AbortSignal): Promise<boolean> {
  try {
    const res = await fetch(`${baseUrlFor(port)}/api/health`, {
      method: 'GET',
      headers: { ...idHeaders(), 'Cache-Control': 'no-cache' },
      signal,
    })
    // 403 still means something is listening and speaking our protocol — it is
    // the host telling us we are not paired. That is a found host, not a miss.
    return res.ok || res.status === 403
  } catch {
    return false
  }
}

/**
 * Locate the host service and fetch a token.
 *
 * On success the shared API client is configured, and every subsequent call
 * from the UI package goes to the right place with the right credentials
 * without knowing any of this happened.
 */
export async function pairWithHost(): Promise<PairingResult> {
  const controller = new AbortController()
  // A loopback service either answers immediately or is not there. Waiting
  // longer than this per port only makes a failed scan feel like a hang.
  const timer = setTimeout(() => controller.abort(), 4000)

  try {
    const remembered = await rememberedPort()
    let found: number | undefined
    for (const port of candidatePorts(remembered)) {
      if (await probe(port, controller.signal)) {
        found = port
        break
      }
    }

    if (found === undefined) {
      return {
        ok: false,
        reason: 'unreachable',
        message: 'The OpenOfficeLLM host service is not running. Start it, then reopen this panel.',
      }
    }

    await rememberPort(found)
    const baseUrl = baseUrlFor(found)

    const res = await fetch(`${baseUrl}/pair`, {
      method: 'GET',
      headers: { ...idHeaders(), 'Cache-Control': 'no-cache' },
      signal: controller.signal,
    })

    if (res.status === 403) {
      return {
        ok: false,
        reason: 'not_paired',
        message: `This extension is not paired with the host. Run:\n\nopenofficellm --pair ${chrome.runtime.id}`,
      }
    }
    if (!res.ok) {
      return { ok: false, reason: 'error', message: `Host returned HTTP ${res.status}.` }
    }

    const body = (await res.json()) as { token?: string; version?: string }
    if (!body.token) {
      return { ok: false, reason: 'error', message: 'Host issued no token.' }
    }

    configureApi({ baseUrl, token: body.token, headers: idHeaders() })
    return { ok: true, baseUrl, version: body.version ?? 'unknown' }
  } catch (e) {
    const err = e as Error
    if (err.name === 'AbortError') {
      return {
        ok: false,
        reason: 'unreachable',
        message: 'Timed out looking for the host service.',
      }
    }
    // A TLS failure lands here, and it is the single most likely first-run
    // problem on a machine where the local CA was never trusted.
    return {
      ok: false,
      reason: 'error',
      message: `${err.message ?? String(e)} — if this mentions a certificate, run the host with --trust-cert.`,
    }
  } finally {
    clearTimeout(timer)
  }
}
