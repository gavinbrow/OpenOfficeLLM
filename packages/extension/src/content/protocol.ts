// The wire format between the side panel and the content script.
//
// Kept in its own file with no imports so the content script — which is bundled
// standalone as an IIFE, because MV3 content scripts are not ES modules — can
// share these types with the panel without dragging anything else in.

export type PageRequest =
  | { kind: 'context'; scope: 'selection' | 'page' }
  | { kind: 'search'; query: string; limit: number }
  | { kind: 'metadata' }
  | { kind: 'links'; limit: number }
  | { kind: 'ping' }

export interface PageContextResult {
  /** What was actually read. A request for the selection falls back to the
   *  page when nothing is selected, and says so here rather than pretending. */
  scope: 'selection' | 'page'
  text: string
  /** Heading outline, as `#`-prefixed lines. */
  outline: string
  title: string
  url: string
  /** True when the text was cut short at the size cap. */
  truncated: boolean
}

export interface PageHit {
  /** A human-readable position: the nearest heading, or a block index. */
  location: string
  text: string
}

export interface PageMetadata {
  title: string
  url: string
  description: string
  siteName: string
  lang: string
  /** Published/modified dates when the page declares them. */
  published: string
  byline: string
}

export interface PageLink {
  text: string
  href: string
}

export type PageResponse =
  | { ok: true; kind: 'context'; data: PageContextResult }
  | { ok: true; kind: 'search'; data: PageHit[] }
  | { ok: true; kind: 'metadata'; data: PageMetadata }
  | { ok: true; kind: 'links'; data: PageLink[] }
  | { ok: true; kind: 'ping' }
  | { ok: false; error: string }

/** Cap on extracted page text, in characters.
 *
 *  A long article runs well past any sensible context window, and the failure
 *  mode of sending it all is not a truncated answer — it is a provider
 *  rejecting the request outright after the user has waited. Cutting here and
 *  flagging `truncated` lets the model ask for a narrower read instead. */
export const MAX_PAGE_CHARS = 120_000
