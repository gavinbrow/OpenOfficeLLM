// HostAdapter over the active browser tab.
//
// The Word adapter talks to Office.js in its own process. This one talks to a
// content script in someone else's page, across `chrome.tabs.sendMessage`, so
// every call can fail in ways a document API cannot: the tab navigated, the
// script was never injected, the user revoked the host permission mid-turn.
// Those all surface as thrown errors with a sentence the model can act on,
// rather than as empty strings that read like an empty page.

import type { ContextScope, DocumentContext } from '@openofficellm/shared'
import { estimateTokens, type HostAdapter, type SearchHit } from '@openofficellm/ui'
import type {
  PageContextResult,
  PageLink,
  PageMetadata,
  PageRequest,
  PageResponse,
} from '../content/protocol'

/** Where the built content script lands. Must match vite.content.config.ts. */
const CONTENT_SCRIPT = 'content.js'

/** The `kind` discriminator of a successful page reply. */
type PageResponseKind = Extract<PageResponse, { ok: true }>['kind']

export class TabUnavailableError extends Error {}

/**
 * The tab the panel is attached to.
 *
 * Held rather than looked up per call because the side panel follows the
 * window, not the tab: reading "the active tab" during a multi-step turn would
 * let a tab switch halfway through silently redirect the second half of the
 * work to a different page.
 */
export class PageAdapter implements HostAdapter {
  readonly host = 'browser' as const

  constructor(private readonly tabId: number) {}

  private async send(req: PageRequest): Promise<PageResponse> {
    try {
      return (await chrome.tabs.sendMessage(this.tabId, req)) as PageResponse
    } catch {
      // No receiver: the script is not in the page yet, or the page reloaded
      // out from under it. Injecting is idempotent — the script guards against
      // double-registering — so the retry is safe.
      await this.inject()
      try {
        return (await chrome.tabs.sendMessage(this.tabId, req)) as PageResponse
      } catch (e) {
        throw new TabUnavailableError(
          `Cannot read this tab (${(e as Error).message ?? 'no response'}). ` +
            'It may be a browser page, a PDF, or a site this extension has not been granted access to.',
        )
      }
    }
  }

  private async inject(): Promise<void> {
    await chrome.scripting.executeScript({
      target: { tabId: this.tabId, allFrames: false },
      files: [CONTENT_SCRIPT],
    })
  }

  private async expect<K extends PageResponseKind>(
    req: PageRequest,
    kind: K,
  ): Promise<Extract<PageResponse, { ok: true; kind: K }>> {
    const res = await this.send(req)
    if (!res.ok) throw new TabUnavailableError(res.error)
    if (res.kind !== kind) throw new TabUnavailableError(`Unexpected page reply "${res.kind}".`)
    return res as Extract<PageResponse, { ok: true; kind: K }>
  }

  async getContext(scope: ContextScope): Promise<DocumentContext> {
    if (scope === 'none') {
      return { host: 'browser', scope: 'none', text: '' }
    }
    // The browser offers exactly two scopes; anything else came from a config
    // written for a different host and means "read the page".
    const wanted: 'selection' | 'page' = scope === 'selection' ? 'selection' : 'page'
    const { data } = await this.expect({ kind: 'context', scope: wanted }, 'context')
    return this.toDocumentContext(data)
  }

  private toDocumentContext(data: PageContextResult): DocumentContext {
    // The URL and title lead the text rather than sitting in a separate field:
    // "summarize this" is unanswerable without knowing what "this" is, and the
    // model reads the context block long before it reads any metadata tool.
    const header = `# ${data.title}\n${data.url}\n\n`
    const text = data.text ? header + data.text : ''
    return {
      host: 'browser',
      scope: data.scope === 'selection' ? 'selection' : 'page',
      text,
      outline: data.outline || undefined,
      tokenEstimate: estimateTokens(text),
    }
  }

  async search(query: string, limit = 20): Promise<SearchHit[]> {
    const { data } = await this.expect({ kind: 'search', query, limit }, 'search')
    return data
  }

  async metadata(): Promise<PageMetadata> {
    const { data } = await this.expect({ kind: 'metadata' }, 'metadata')
    return data
  }

  async links(limit = 100): Promise<PageLink[]> {
    const { data } = await this.expect({ kind: 'links', limit }, 'links')
    return data
  }

  /** True if the content script answers. Used to decide whether the panel has
   *  a page to work with before it claims it does. */
  async reachable(): Promise<boolean> {
    try {
      await this.expect({ kind: 'ping' }, 'ping')
      return true
    } catch {
      return false
    }
  }

  // `applyEdits`, `snapshot` and `restore` are deliberately absent — see the
  // header of content/page.ts. The adapter contract makes them optional
  // precisely so a read-only host can say so in the type system rather than by
  // throwing at the point of use.
}
