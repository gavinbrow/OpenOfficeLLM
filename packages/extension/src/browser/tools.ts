// The tool catalog for a browser tab.
//
// Read-only, and short. The Word catalog is large because Word has a large
// vocabulary of things a user might want changed; a web page has one thing the
// user wants — to know what it says — and a handful of ways to ask. Adding
// speculative write tools here would not just be unused, it would be wrong:
// see the header of content/page.ts.

import type { DetectedHost, ToolDefinition } from '@openofficellm/shared'
import type { ToolExecContext, ToolOutcome } from '@openofficellm/ui'
import { PageAdapter } from './pageAdapter'

const READ_TOOLS: ToolDefinition[] = [
  {
    name: 'read_page',
    description:
      'Read the main readable content of the page the user is looking at: navigation, headers, footers and sidebars are stripped. Use this whenever the question is about "this page", "this article", or what the user is currently reading.',
    parameters: { type: 'object', properties: {}, required: [] },
  },
  {
    name: 'read_selection',
    description:
      'Read the text the user currently has selected on the page. Falls back to the whole page when nothing is selected, and says which it returned.',
    parameters: { type: 'object', properties: {}, required: [] },
  },
  {
    name: 'search_page',
    description:
      'Find every occurrence of a phrase on the page, with surrounding context and the nearest heading. Much cheaper than read_page when the user asks about one specific thing on a long page.',
    parameters: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Text to look for. Case-insensitive.' },
        limit: { type: 'number', description: 'Maximum hits to return. Default 20.' },
      },
      required: ['query'],
    },
  },
  {
    name: 'read_metadata',
    description:
      'Title, URL, site name, description, language, author and publication date, as the page declares them. Use to establish what a page is and who published it before trusting its claims.',
    parameters: { type: 'object', properties: {}, required: [] },
  },
  {
    name: 'read_links',
    description:
      'Links in the main content, with their text. Use when the user asks what a page links to, or to find the URL behind something they named.',
    parameters: {
      type: 'object',
      properties: {
        limit: { type: 'number', description: 'Maximum links to return. Default 100.' },
      },
      required: [],
    },
  },
]

/**
 * Tools available on this host.
 *
 * `allowWrites` is accepted and ignored: there are no write tools to gate. The
 * parameter stays because the Shell contract is shared, and a signature that
 * quietly differed per shell would be a trap for the next person to add one.
 */
export function toolCatalog(host: DetectedHost, _allowWrites: boolean): ToolDefinition[] {
  return host === 'browser' ? [...READ_TOOLS] : []
}

export function isWriteTool(_name: string): boolean {
  return false
}

function parseArgs(json: string): Record<string, unknown> {
  try {
    const v = JSON.parse(json || '{}')
    return v && typeof v === 'object' ? (v as Record<string, unknown>) : {}
  } catch {
    return {}
  }
}

function numberArg(args: Record<string, unknown>, key: string, fallback: number): number {
  const v = args[key]
  return typeof v === 'number' && Number.isFinite(v) && v > 0 ? Math.floor(v) : fallback
}

export async function executeDocumentTool(
  name: string,
  argsJson: string,
  ctx: ToolExecContext,
): Promise<ToolOutcome> {
  const args = parseArgs(argsJson)
  const adapter = ctx.adapter
  if (!(adapter instanceof PageAdapter)) {
    return { content: 'No page is attached, so page tools are unavailable.', isError: true }
  }

  try {
    switch (name) {
      case 'read_page':
      case 'read_selection': {
        const scope = name === 'read_selection' ? 'selection' : 'page'
        const c = await adapter.getContext(scope)
        if (!c.text) {
          return { content: 'The page appears to have no readable text.', isError: false }
        }
        // Say what was actually read when it differs from what was asked. A
        // model that assumes it got a selection will attribute the whole
        // article to the user's cursor.
        const note =
          scope === 'selection' && c.scope !== 'selection'
            ? 'Nothing was selected, so this is the whole page.\n\n'
            : ''
        const outline = c.outline ? `\n\n## Outline\n${c.outline}` : ''
        return { content: note + c.text + outline, isError: false }
      }

      case 'search_page': {
        const query = typeof args.query === 'string' ? args.query : ''
        if (!query.trim()) {
          return { content: 'search_page needs a "query" to look for.', isError: true }
        }
        const hits = await adapter.search(query, numberArg(args, 'limit', 20))
        if (hits.length === 0) {
          return { content: `No occurrences of "${query}" on this page.`, isError: false }
        }
        const body = hits.map((h) => `[${h.location}] ${h.text}`).join('\n')
        return {
          content: `${hits.length} occurrence${hits.length === 1 ? '' : 's'} of "${query}":\n${body}`,
          isError: false,
        }
      }

      case 'read_metadata': {
        const m = await adapter.metadata()
        const rows = Object.entries(m)
          .filter(([, v]) => typeof v === 'string' && v.length > 0)
          .map(([k, v]) => `${k}: ${String(v)}`)
        return { content: rows.join('\n') || 'The page declares no metadata.', isError: false }
      }

      case 'read_links': {
        const links = await adapter.links(numberArg(args, 'limit', 100))
        if (links.length === 0) {
          return { content: 'No links in the main content of this page.', isError: false }
        }
        return { content: links.map((l) => `${l.text} → ${l.href}`).join('\n'), isError: false }
      }

      default:
        return { content: `Unknown tool: ${name}`, isError: true }
    }
  } catch (e) {
    return { content: `Tool "${name}" failed: ${(e as Error).message ?? String(e)}`, isError: true }
  }
}
