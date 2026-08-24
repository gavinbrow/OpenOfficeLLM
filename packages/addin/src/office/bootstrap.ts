// Office.js readiness and theme detection.
//
// Types come from @types/office-js (a devDependency — no runtime cost). The
// runtime itself is the office.js script tag, which the host service rewrites
// to its local cache when the CDN is unreachable.

import type { DetectedHost } from '@openofficellm/shared'

export interface OfficeTheme {
  isDark: boolean
}

export interface OfficeBootstrapResult {
  host: DetectedHost
  theme: OfficeTheme
}

/** Narrow accessor for the global. `Office` is undefined when the pane is
 *  opened in a plain browser during development. */
function office(): typeof Office | undefined {
  return (globalThis as { Office?: typeof Office }).Office
}

let _host: DetectedHost = 'none'
let _theme: OfficeTheme = detectPrefersDark()
let _documentKey: string = freshDocumentKey()

/** Identity for a document we cannot name.
 *
 *  Office gives every document window its own add-in instance, so a key minted
 *  once per instance is already per-document — it just cannot be recognised
 *  again after a reload. That is the right failure direction: an unrecognised
 *  document gets a fresh chat rather than inheriting someone else's. */
function freshDocumentKey(): string {
  return `anon:${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`
}

/** Stable per-document identity, used to keep each document's chat separate.
 *
 *  The task pane's localStorage is scoped to the add-in's origin, not to the
 *  document, so without this every document in Word reopens the last chat from
 *  whichever document was used previously. */
function deriveDocumentKey(): string {
  try {
    const url = office()?.context?.document?.url
    if (typeof url === 'string' && url.trim().length > 0) return `doc:${url.trim()}`
  } catch {
    // An unsaved document can throw here rather than return an empty string.
  }
  return freshDocumentKey()
}

export function getDocumentKey(): string {
  return _documentKey
}

function detectPrefersDark(): OfficeTheme {
  if (typeof window === 'undefined' || !window.matchMedia) return { isDark: false }
  return { isDark: window.matchMedia('(prefers-color-scheme: dark)').matches }
}

function applyThemeClass(isDark: boolean): void {
  const el = document.documentElement
  el.classList.toggle('dark', isDark)
}

function deriveTheme(): OfficeTheme {
  try {
    const officeTheme = office()?.context?.officeTheme
    if (officeTheme?.bodyBackgroundColor) {
      // Office reports colours as #RRGGBB, but some builds return #AARRGGBB.
      // Taking the last six hex digits is correct for both.
      const raw = officeTheme.bodyBackgroundColor.replace('#', '')
      const hex = raw.length > 6 ? raw.slice(-6) : raw
      const r = parseInt(hex.slice(0, 2), 16)
      const g = parseInt(hex.slice(2, 4), 16)
      const b = parseInt(hex.slice(4, 6), 16)
      const luma = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255
      return { isDark: luma < 0.5 }
    }
  } catch {
    // fall through to media query
  }
  return detectPrefersDark()
}

export function getHost(): DetectedHost {
  return _host
}

export function getTheme(): OfficeTheme {
  return _theme
}

/** Initialise Office.js readiness detection and theme handling. Idempotent.
 *  Falls through to a 'none' host if Office.onReady doesn't fire within 8s —
 *  this prevents the pane from hanging forever if the office.js script tag
 *  loads but the host never invokes onReady (e.g. CDN blocked, WebView2
 *  offline, or a broken office.js build). */
export async function bootstrap(): Promise<OfficeBootstrapResult> {
  applyThemeClass(_theme.isDark)

  if (typeof window === 'undefined' || !office()?.onReady) {
    _host = 'none'
    _theme = detectPrefersDark()
    applyThemeClass(_theme.isDark)
    listenToPrefersColorScheme()
    return { host: _host, theme: _theme }
  }

  return new Promise<OfficeBootstrapResult>((resolve) => {
    let settled = false
    const onTimeout = () => {
      if (settled) return
      settled = true
      _host = 'none'
      _theme = detectPrefersDark()
      applyThemeClass(_theme.isDark)
      listenToPrefersColorScheme()
      resolve({ host: _host, theme: _theme })
    }
    // 8s is generous — Office.onReady typically fires within a few hundred ms.
    const timer = setTimeout(onTimeout, 8000)

    void office()!.onReady((info) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      const hostKey = String(info.host ?? '').toLowerCase()
      if (hostKey === 'word' || hostKey === 'document') _host = 'word'
      else if (hostKey === 'excel' || hostKey === 'workbook') _host = 'excel'
      else _host = 'none'

      _documentKey = deriveDocumentKey()
      _theme = deriveTheme()
      applyThemeClass(_theme.isDark)
      listenToPrefersColorScheme()
      resolve({ host: _host, theme: _theme })
    })
  })
}

function listenToPrefersColorScheme(): void {
  if (typeof window === 'undefined' || !window.matchMedia) return
  const mq = window.matchMedia('(prefers-color-scheme: dark)')
  const onChange = (e: MediaQueryListEvent) => {
    if (office()?.context?.officeTheme) return
    _theme = { isDark: e.matches }
    applyThemeClass(_theme.isDark)
  }
  if (typeof mq.addEventListener === 'function') mq.addEventListener('change', onChange)
  else mq.addListener(onChange)
}
