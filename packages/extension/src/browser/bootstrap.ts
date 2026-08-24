// Working out what the side panel is attached to.
//
// The Office equivalent of this file waits for `Office.onReady` and asks which
// application it is in. Here the question is messier: a side panel belongs to a
// *window*, and the window's active tab changes under it. Some tabs cannot be
// read at all — `chrome://` pages, the Web Store, PDFs, and any site the user
// has not granted host access to — and the honest answer for those is the same
// 'none' host the pane reports outside Office.

import type { DetectedHost } from '@openofficellm/shared'
import { PageAdapter } from './pageAdapter'

export interface AttachedTab {
  id: number
  url: string
  title: string
}

/** Schemes no extension may script, whatever permissions it holds. Listed so
 *  the panel can explain itself instead of surfacing a bare injection error. */
const BLOCKED_SCHEMES = [
  'chrome:',
  'chrome-extension:',
  'edge:',
  'about:',
  'devtools:',
  'view-source:',
]
const BLOCKED_HOSTS = ['chromewebstore.google.com', 'chrome.google.com/webstore']

export function isScriptableUrl(url: string): boolean {
  if (!url) return false
  let parsed: URL
  try {
    parsed = new URL(url)
  } catch {
    return false
  }
  if (BLOCKED_SCHEMES.includes(parsed.protocol)) return false
  if (BLOCKED_HOSTS.some((h) => url.startsWith(`https://${h}`))) return false
  // file:// works only when the user ticks "Allow access to file URLs", which
  // is not something this can detect ahead of time — so it is allowed through
  // and the injection failure, if any, is reported for what it is.
  return parsed.protocol === 'http:' || parsed.protocol === 'https:' || parsed.protocol === 'file:'
}

let _host: DetectedHost = 'none'
let _adapter: PageAdapter | null = null
let _tab: AttachedTab | null = null
let _dark = prefersDark()

function prefersDark(): boolean {
  if (typeof window === 'undefined' || !window.matchMedia) return false
  return window.matchMedia('(prefers-color-scheme: dark)').matches
}

export function getHost(): DetectedHost {
  return _host
}

export function getAdapter(): PageAdapter | null {
  return _adapter
}

export function getTab(): AttachedTab | null {
  return _tab
}

export function isDark(): boolean {
  return _dark
}

/**
 * Identity for the attached tab, used to keep each tab's chat separate.
 *
 * Keyed on the tab, not the URL. In Word each *document* gets a chat; the
 * browser analogue is the tab, because that is the thing the user thinks of as
 * "where I was". Keying on URL instead would start a fresh chat on every
 * navigation, including a link the user followed precisely because of what the
 * assistant just said.
 *
 * Chrome tab ids are unique for the tab's lifetime and are reissued after a
 * browser restart, so a restart lands on a fresh chat — the same failure
 * direction the task pane takes for an unsaved document.
 */
export function getDocumentKey(): string {
  return _tab ? `tab:${_tab.id}` : 'detached'
}

async function activeTab(): Promise<chrome.tabs.Tab | undefined> {
  // `lastFocusedWindow` rather than `currentWindow`: a side panel is not itself
  // a window, and `currentWindow` can resolve to the panel's own context and
  // return nothing.
  const [tab] = await chrome.tabs.query({ active: true, lastFocusedWindow: true })
  return tab
}

/** Whether this extension may script the given URL right now. */
export async function hasAccessTo(url: string): Promise<boolean> {
  if (!isScriptableUrl(url)) return false
  try {
    const origin = new URL(url).origin
    return await chrome.permissions.contains({ origins: [`${origin}/*`] })
  } catch {
    return false
  }
}

/** Ask for access to a site. Must be called from a user gesture — Chrome
 *  silently denies a permission prompt that is not. */
export async function requestAccessTo(url: string): Promise<boolean> {
  try {
    const origin = new URL(url).origin
    return await chrome.permissions.request({ origins: [`${origin}/*`] })
  } catch {
    return false
  }
}

/**
 * Point the shell at whatever tab is in front, and report whether it changed.
 *
 * Callers use the return value to decide whether to swap conversations, so it
 * must be false on a re-check that found the same tab — otherwise every focus
 * event would reset the chat the user is in the middle of.
 */
export async function attachToActiveTab(): Promise<{ changed: boolean; host: DetectedHost }> {
  const tab = await activeTab()
  const previousKey = getDocumentKey()

  if (!tab || tab.id === undefined || !isScriptableUrl(tab.url ?? '')) {
    _host = 'none'
    _adapter = null
    _tab =
      tab && tab.id !== undefined
        ? { id: tab.id, url: tab.url ?? '', title: tab.title ?? '' }
        : null
    return { changed: getDocumentKey() !== previousKey, host: _host }
  }

  const url = tab.url ?? ''
  _tab = { id: tab.id, url, title: tab.title ?? '' }

  if (!(await hasAccessTo(url))) {
    // The tab is readable in principle but this extension has not been granted
    // the site. That is a 'none' host with a specific, fixable cause, which
    // the access gate in the panel explains and offers to fix.
    _host = 'none'
    _adapter = null
    return { changed: getDocumentKey() !== previousKey, host: _host }
  }

  _adapter = new PageAdapter(tab.id)
  _host = 'browser'
  return { changed: getDocumentKey() !== previousKey, host: _host }
}

/** Watch for the user switching tab or navigating, and re-attach. */
export function watchTabs(onChange: (changed: boolean) => void): () => void {
  const recheck = () => {
    void attachToActiveTab().then(({ changed }) => onChange(changed))
  }

  const onActivated = () => recheck()
  const onUpdated = (_id: number, info: chrome.tabs.TabChangeInfo) => {
    // Only when navigation completes. Firing on every 'loading' tick would
    // re-attach several times per page load for no gain.
    if (info.status === 'complete' || info.url !== undefined) recheck()
  }
  const onFocus = () => recheck()

  chrome.tabs.onActivated.addListener(onActivated)
  chrome.tabs.onUpdated.addListener(onUpdated)
  chrome.windows.onFocusChanged.addListener(onFocus)

  return () => {
    chrome.tabs.onActivated.removeListener(onActivated)
    chrome.tabs.onUpdated.removeListener(onUpdated)
    chrome.windows.onFocusChanged.removeListener(onFocus)
  }
}

/** Follow the browser's colour scheme, applying the class the UI's tokens key
 *  off. The task pane takes this from Office's reported theme; a side panel has
 *  no such signal, so the media query is the whole answer. */
export function watchColorScheme(onChange: (dark: boolean) => void): void {
  if (typeof window === 'undefined' || !window.matchMedia) return
  const mq = window.matchMedia('(prefers-color-scheme: dark)')
  const apply = (dark: boolean) => {
    _dark = dark
    document.documentElement.classList.toggle('dark', dark)
    onChange(dark)
  }
  apply(mq.matches)
  mq.addEventListener('change', (e) => apply(e.matches))
}
