// The content script: the only code in this project that runs inside somebody
// else's page.
//
// It reads and answers questions. It does not write. That is not a limitation
// waiting to be lifted — a web page is served by its owner and rendered for the
// user, and an assistant that silently rewrites it produces something that
// looks authoritative, screenshots convincingly, and corresponds to nothing.
// The Word adapter edits because the document belongs to the user. This one
// does not, and the adapter it backs deliberately omits `applyEdits`.
//
// Bundled as a standalone IIFE (see vite.content.config.ts) because MV3 content
// scripts are not ES modules. Everything it needs is either inlined or a
// type-only import, which erases.

import {
  MAX_PAGE_CHARS,
  type PageContextResult,
  type PageHit,
  type PageLink,
  type PageMetadata,
  type PageRequest,
  type PageResponse,
} from './protocol'

/** Elements whose text is never page content: chrome, navigation, and the
 *  things that survive an ad blocker. Excluded before extraction rather than
 *  filtered afterwards, because once it is all one string the cookie banner is
 *  indistinguishable from the first paragraph. */
const NOISE_SELECTOR = [
  'script',
  'style',
  'noscript',
  'template',
  'svg',
  'canvas',
  'iframe',
  'nav',
  'header',
  'footer',
  'aside',
  '[aria-hidden="true"]',
  '[role="navigation"]',
  '[role="banner"]',
  '[role="contentinfo"]',
  '[role="complementary"]',
  '[hidden]',
].join(',')

/** Containers that usually hold the actual article, most specific first. */
const CONTENT_SELECTOR = ['article', 'main', '[role="main"]', '#content', '.content'].join(',')

function isVisible(el: Element): boolean {
  const style = window.getComputedStyle(el)
  if (style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0') {
    return false
  }
  // A zero-height box with no overflow shows nothing, but `details` collapse
  // that way legitimately, so only reject when there is genuinely no box.
  const rect = (el as HTMLElement).getBoundingClientRect?.()
  return !rect || rect.width > 0 || rect.height > 0
}

/**
 * The readable root of the page.
 *
 * Prefers a declared content container, but only when it actually holds the
 * bulk of the text: plenty of sites wrap their nav in `<main>` and the article
 * in a div, and a `<main>` holding forty characters is a worse answer than the
 * body. The 40% threshold is a heuristic, and a wrong guess here costs a
 * slightly noisier read, not a failure.
 */
function readableRoot(): Element {
  const body = document.body
  const bodyLen = (body?.innerText ?? '').length
  const candidates = Array.from(document.querySelectorAll(CONTENT_SELECTOR))
  let best: Element | null = null
  let bestLen = 0
  for (const el of candidates) {
    if (!isVisible(el)) continue
    const len = ((el as HTMLElement).innerText ?? '').length
    if (len > bestLen) {
      best = el
      bestLen = len
    }
  }
  if (best && bodyLen > 0 && bestLen / bodyLen >= 0.4) return best
  return body ?? document.documentElement
}

/** A copy of `root` with the noise removed, so extraction does not mutate the
 *  live page. Cloning is what keeps this script strictly read-only. */
function cleanedClone(root: Element): HTMLElement {
  const clone = root.cloneNode(true) as HTMLElement
  for (const el of Array.from(clone.querySelectorAll(NOISE_SELECTOR))) {
    el.remove()
  }
  return clone
}

/** Collapse the runs of blank lines that survive element removal. */
function tidy(text: string): string {
  return text
    .replace(/\r\n/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
}

function cap(text: string): { text: string; truncated: boolean } {
  if (text.length <= MAX_PAGE_CHARS) return { text, truncated: false }
  return { text: `${text.slice(0, MAX_PAGE_CHARS)}\n\n[…page truncated]`, truncated: true }
}

function headingOutline(root: Element): string {
  const lines: string[] = []
  for (const h of Array.from(root.querySelectorAll('h1,h2,h3,h4,h5,h6'))) {
    if (!isVisible(h)) continue
    const text = (h.textContent ?? '').trim()
    if (!text) continue
    const level = Number(h.tagName.slice(1))
    lines.push(`${'#'.repeat(level)} ${text}`)
    if (lines.length >= 200) break
  }
  return lines.join('\n')
}

function selectionText(): string {
  const sel = window.getSelection()
  return sel ? tidy(sel.toString()) : ''
}

function readContext(scope: 'selection' | 'page'): PageContextResult {
  const root = readableRoot()
  if (scope === 'selection') {
    const text = selectionText()
    if (text) {
      const { text: capped, truncated } = cap(text)
      return {
        scope: 'selection',
        text: capped,
        outline: '',
        title: document.title,
        url: location.href,
        truncated,
      }
    }
    // Nothing selected. Fall through to the page and report the scope that was
    // actually used — a caller told "selection" when it got the whole article
    // will summarize the wrong thing and never know why.
  }
  const { text, truncated } = cap(tidy(cleanedClone(root).innerText ?? ''))
  return {
    scope: 'page',
    text,
    outline: headingOutline(root),
    title: document.title,
    url: location.href,
    truncated,
  }
}

/** Nearest preceding heading, used to say *where* a search hit is. */
function locationOf(node: Node): string {
  let el: Element | null =
    node.nodeType === Node.ELEMENT_NODE ? (node as Element) : node.parentElement
  while (el) {
    let sib: Element | null = el.previousElementSibling
    while (sib) {
      if (/^H[1-6]$/.test(sib.tagName)) {
        const t = (sib.textContent ?? '').trim()
        if (t) return t.slice(0, 80)
      }
      sib = sib.previousElementSibling
    }
    el = el.parentElement
  }
  return 'page'
}

function searchPage(query: string, limit: number): PageHit[] {
  const needle = query.toLowerCase()
  if (!needle) return []
  const root = readableRoot()
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
  const hits: PageHit[] = []
  let node: Node | null
  while ((node = walker.nextNode())) {
    const text = node.nodeValue ?? ''
    const idx = text.toLowerCase().indexOf(needle)
    if (idx === -1) continue
    if (node.parentElement && !isVisible(node.parentElement)) continue
    // Enough either side to read the hit in context without pasting the
    // paragraph twice for two hits in the same sentence.
    const start = Math.max(0, idx - 60)
    const end = Math.min(text.length, idx + needle.length + 60)
    hits.push({
      location: locationOf(node),
      text: `${start > 0 ? '…' : ''}${text.slice(start, end).trim()}${end < text.length ? '…' : ''}`,
    })
    if (hits.length >= limit) break
  }
  return hits
}

function metaContent(...names: string[]): string {
  for (const name of names) {
    const el =
      document.querySelector(`meta[property="${name}"]`) ??
      document.querySelector(`meta[name="${name}"]`)
    const v = el?.getAttribute('content')?.trim()
    if (v) return v
  }
  return ''
}

function readMetadata(): PageMetadata {
  return {
    title: document.title,
    url: location.href,
    description: metaContent('og:description', 'description', 'twitter:description'),
    siteName: metaContent('og:site_name', 'application-name'),
    lang: document.documentElement.lang || '',
    published: metaContent('article:published_time', 'datePublished', 'date'),
    byline: metaContent('article:author', 'author'),
  }
}

function readLinks(limit: number): PageLink[] {
  const out: PageLink[] = []
  const seen = new Set<string>()
  for (const a of Array.from(readableRoot().querySelectorAll('a[href]'))) {
    const href = (a as HTMLAnchorElement).href
    const text = (a.textContent ?? '').trim()
    // Anchors and javascript: URLs are navigation, not references.
    if (!href || !text || href.startsWith('javascript:') || href === `${location.href}#`) continue
    if (seen.has(href)) continue
    seen.add(href)
    out.push({ text: text.slice(0, 120), href })
    if (out.length >= limit) break
  }
  return out
}

function handle(req: PageRequest): PageResponse {
  try {
    switch (req.kind) {
      case 'ping':
        return { ok: true, kind: 'ping' }
      case 'context':
        return { ok: true, kind: 'context', data: readContext(req.scope) }
      case 'search':
        return { ok: true, kind: 'search', data: searchPage(req.query, req.limit) }
      case 'metadata':
        return { ok: true, kind: 'metadata', data: readMetadata() }
      case 'links':
        return { ok: true, kind: 'links', data: readLinks(req.limit) }
      default:
        return { ok: false, error: 'Unknown page request.' }
    }
  } catch (e) {
    return { ok: false, error: (e as Error).message ?? String(e) }
  }
}

// Injected repeatedly across navigations, so guard against double-registering
// a listener — two replies to one message is an error in Chrome, not a
// harmless duplicate.
const FLAG = '__openofficellm_page_reader__'
declare global {
  interface Window {
    [FLAG]?: boolean
  }
}

if (!window[FLAG]) {
  window[FLAG] = true
  chrome.runtime.onMessage.addListener((message: unknown, _sender, sendResponse) => {
    sendResponse(handle(message as PageRequest))
    // Synchronous reply: returning true would leave the channel open for a
    // response that has already been sent.
    return false
  })
}
