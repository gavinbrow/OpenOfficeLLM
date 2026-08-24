// The browser tool catalog and dispatcher.
//
// The load-bearing assertion in here is that there are no write tools. The
// adapter cannot apply edits and the content script never mutates the page, so
// a write tool appearing in this catalog would be a tool that always fails —
// and worse, one the model would keep trying.

import { describe, it, expect, vi } from 'vitest'
import type { ToolExecContext } from '@openofficellm/ui'
import { toolCatalog, isWriteTool, executeDocumentTool } from '../tools'
import { PageAdapter } from '../pageAdapter'

/** A PageAdapter whose tab replies with canned data. Constructed through the
 *  real class because the dispatcher checks `instanceof`. */
function fakeAdapter(overrides: Partial<Record<string, unknown>> = {}): PageAdapter {
  const adapter = new PageAdapter(1)
  Object.assign(adapter, overrides)
  return adapter
}

function ctx(adapter: PageAdapter): ToolExecContext {
  return { adapter, mode: 'propose', propose: () => {} }
}

describe('toolCatalog', () => {
  it('offers page reads on the browser host', () => {
    const names = toolCatalog('browser', true).map((t) => t.name)
    expect(names).toEqual([
      'read_page',
      'read_selection',
      'search_page',
      'read_metadata',
      'read_links',
    ])
  })

  it('offers nothing when no page is attached', () => {
    expect(toolCatalog('none', true)).toEqual([])
  })

  it('offers nothing for a host this shell does not serve', () => {
    expect(toolCatalog('word', true)).toEqual([])
    expect(toolCatalog('excel', true)).toEqual([])
  })

  it('is identical with writes allowed and disallowed, because there are none', () => {
    expect(toolCatalog('browser', false)).toEqual(toolCatalog('browser', true))
    for (const tool of toolCatalog('browser', true)) {
      expect(isWriteTool(tool.name)).toBe(false)
    }
  })

  it('describes every tool and gives it a parameters object', () => {
    for (const tool of toolCatalog('browser', true)) {
      expect(tool.description.length).toBeGreaterThan(20)
      expect(tool.parameters.type).toBe('object')
    }
  })
})

describe('executeDocumentTool', () => {
  it('reads the page', async () => {
    const adapter = fakeAdapter({
      getContext: vi.fn().mockResolvedValue({
        host: 'browser',
        scope: 'page',
        text: '# Title\nhttps://x.test\n\nBody text',
        outline: '# Title',
      }),
    })
    const out = await executeDocumentTool('read_page', '{}', ctx(adapter))
    expect(out.isError).toBe(false)
    expect(out.content).toContain('Body text')
    expect(out.content).toContain('## Outline')
  })

  it('says so when a selection read fell back to the whole page', async () => {
    const adapter = fakeAdapter({
      getContext: vi.fn().mockResolvedValue({ host: 'browser', scope: 'page', text: 'all of it' }),
    })
    const out = await executeDocumentTool('read_selection', '{}', ctx(adapter))
    // Without this the model attributes the whole article to the user's cursor.
    expect(out.content).toMatch(/Nothing was selected/)
  })

  it('does not add the fallback note when a selection really was read', async () => {
    const adapter = fakeAdapter({
      getContext: vi
        .fn()
        .mockResolvedValue({ host: 'browser', scope: 'selection', text: 'just this' }),
    })
    const out = await executeDocumentTool('read_selection', '{}', ctx(adapter))
    expect(out.content).not.toMatch(/Nothing was selected/)
    expect(out.content).toBe('just this')
  })

  it('reports an empty page rather than returning nothing', async () => {
    const adapter = fakeAdapter({
      getContext: vi.fn().mockResolvedValue({ host: 'browser', scope: 'page', text: '' }),
    })
    const out = await executeDocumentTool('read_page', '{}', ctx(adapter))
    expect(out.isError).toBe(false)
    expect(out.content).toMatch(/no readable text/)
  })

  it('refuses a search with no query instead of scanning for the empty string', async () => {
    const search = vi.fn()
    const out = await executeDocumentTool(
      'search_page',
      '{"query":"  "}',
      ctx(fakeAdapter({ search })),
    )
    expect(out.isError).toBe(true)
    expect(search).not.toHaveBeenCalled()
  })

  it('reports zero hits as a normal answer, not an error', async () => {
    const adapter = fakeAdapter({ search: vi.fn().mockResolvedValue([]) })
    const out = await executeDocumentTool('search_page', '{"query":"nope"}', ctx(adapter))
    expect(out.isError).toBe(false)
    expect(out.content).toMatch(/No occurrences/)
  })

  it('renders hits with the heading they were found under', async () => {
    const adapter = fakeAdapter({
      search: vi.fn().mockResolvedValue([{ location: 'Pricing', text: '…costs £9…' }]),
    })
    const out = await executeDocumentTool('search_page', '{"query":"costs"}', ctx(adapter))
    expect(out.content).toContain('[Pricing] …costs £9…')
  })

  it('passes a sane limit through and ignores a nonsense one', async () => {
    const search = vi.fn().mockResolvedValue([])
    const adapter = fakeAdapter({ search })
    await executeDocumentTool('search_page', '{"query":"a","limit":5}', ctx(adapter))
    expect(search).toHaveBeenLastCalledWith('a', 5)
    await executeDocumentTool('search_page', '{"query":"a","limit":-3}', ctx(adapter))
    expect(search).toHaveBeenLastCalledWith('a', 20)
  })

  it('drops metadata fields the page did not declare', async () => {
    const adapter = fakeAdapter({
      metadata: vi.fn().mockResolvedValue({
        title: 'T',
        url: 'https://x.test',
        description: '',
        siteName: '',
        lang: 'en',
        published: '',
        byline: '',
      }),
    })
    const out = await executeDocumentTool('read_metadata', '{}', ctx(adapter))
    expect(out.content).toContain('title: T')
    expect(out.content).toContain('lang: en')
    expect(out.content).not.toContain('description:')
  })

  it('surfaces an adapter failure as a tool error the model can read', async () => {
    const adapter = fakeAdapter({
      getContext: vi.fn().mockRejectedValue(new Error('tab navigated away')),
    })
    const out = await executeDocumentTool('read_page', '{}', ctx(adapter))
    expect(out.isError).toBe(true)
    expect(out.content).toContain('tab navigated away')
  })

  it('rejects an unknown tool name', async () => {
    const out = await executeDocumentTool('rewrite_page', '{}', ctx(fakeAdapter()))
    expect(out.isError).toBe(true)
    expect(out.content).toMatch(/Unknown tool/)
  })

  it('survives malformed arguments rather than throwing', async () => {
    const adapter = fakeAdapter({
      getContext: vi.fn().mockResolvedValue({ host: 'browser', scope: 'page', text: 'ok' }),
    })
    const out = await executeDocumentTool('read_page', 'not json', ctx(adapter))
    expect(out.isError).toBe(false)
  })
})
