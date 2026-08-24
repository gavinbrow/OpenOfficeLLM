import { describe, it, expect, vi } from 'vitest'
import type { Edit } from '@openofficellm/shared'
import { executeDocumentTool, toolCatalog, isWriteTool } from '../tools'
import type { ApplyResult, HostAdapter } from '@openofficellm/ui'

/** A recording stand-in for a host adapter. */
function fakeAdapter(overrides: Partial<HostAdapter> = {}) {
  const applied: Edit[][] = []
  const adapter: HostAdapter = {
    host: 'word',
    getContext: vi.fn(async (scope) => ({
      host: 'word' as const,
      scope,
      text: `context for ${scope}`,
      outline: 'OUTLINE',
      tokenEstimate: 5,
    })),
    applyEdits: vi.fn(async (edits: Edit[]): Promise<ApplyResult> => {
      applied.push(edits)
      return { ok: true, summary: 'applied' }
    }),
    snapshot: vi.fn(),
    restore: vi.fn(),
    search: vi.fn(async () => [{ location: '3', text: 'a hit' }]),
    ...overrides,
  } as HostAdapter
  return { adapter, applied }
}

function ctx(mode: 'propose' | 'direct' | 'agentic', adapter: HostAdapter) {
  const proposed: Array<{ edit: Edit; description: string }> = []
  return {
    toolCtx: {
      adapter,
      mode,
      propose: (edit: Edit, description: string) => proposed.push({ edit, description }),
    },
    proposed,
  }
}

describe('toolCatalog', () => {
  it('offers Word tools in Word and Excel tools in Excel', () => {
    const word = toolCatalog('word', true).map((t) => t.name)
    const excel = toolCatalog('excel', true).map((t) => t.name)
    expect(word).toContain('replace_selection')
    expect(word).not.toContain('write_formula')
    expect(excel).toContain('write_formula')
    expect(excel).not.toContain('replace_selection')
  })

  // A model that can rewrite what it cannot read is strictly worse than one
  // that can do neither.
  it('withholds write tools when writes are not allowed', () => {
    const names = toolCatalog('word', false).map((t) => t.name)
    expect(names).toContain('read_document')
    expect(names.some(isWriteTool)).toBe(false)
  })

  it('offers nothing outside Office', () => {
    expect(toolCatalog('none', true)).toEqual([])
  })

  it('gives every tool a real parameter schema', () => {
    for (const tool of [...toolCatalog('word', true), ...toolCatalog('excel', true)]) {
      expect(tool.parameters.type).toBe('object')
      expect(tool.description.length).toBeGreaterThan(10)
    }
  })
})

describe('read tools', () => {
  it('read_document returns the outline and body', async () => {
    const { adapter } = fakeAdapter()
    const out = await executeDocumentTool('read_document', '{}', ctx('propose', adapter).toolCtx)
    expect(out.isError).toBe(false)
    expect(out.content).toContain('OUTLINE')
    expect(out.content).toContain('context for document')
  })

  it('search_document requires a query', async () => {
    const { adapter } = fakeAdapter()
    const out = await executeDocumentTool('search_document', '{}', ctx('propose', adapter).toolCtx)
    expect(out.isError).toBe(true)
    expect(out.content).toContain('query')
  })

  it('search_document formats hits with their locations', async () => {
    const { adapter } = fakeAdapter()
    const out = await executeDocumentTool(
      'search_document',
      '{"query":"freight"}',
      ctx('propose', adapter).toolCtx,
    )
    expect(out.content).toContain('[3] a hit')
  })
})

describe('edit modes', () => {
  it('propose stages the edit and does NOT touch the document', async () => {
    const { adapter, applied } = fakeAdapter()
    const c = ctx('propose', adapter)
    const out = await executeDocumentTool('replace_selection', '{"text":"tightened"}', c.toolCtx)
    expect(applied).toHaveLength(0)
    expect(c.proposed).toHaveLength(1)
    expect(c.proposed[0].edit).toEqual({ kind: 'replaceSelection', text: 'tightened' })
    expect(out.isError).toBe(false)
  })

  // The model must not tell the user it changed the document when it has not.
  it('propose tells the model nothing has been applied', async () => {
    const { adapter } = fakeAdapter()
    const out = await executeDocumentTool(
      'replace_selection',
      '{"text":"x"}',
      ctx('propose', adapter).toolCtx,
    )
    expect(out.content).toMatch(/NOT been applied/i)
  })

  it('direct applies immediately', async () => {
    const { adapter, applied } = fakeAdapter()
    const c = ctx('direct', adapter)
    await executeDocumentTool('replace_selection', '{"text":"now"}', c.toolCtx)
    expect(c.proposed).toHaveLength(0)
    expect(applied).toEqual([[{ kind: 'replaceSelection', text: 'now' }]])
  })

  it('agentic applies immediately', async () => {
    const { adapter, applied } = fakeAdapter()
    await executeDocumentTool('replace_selection', '{"text":"go"}', ctx('agentic', adapter).toolCtx)
    expect(applied).toHaveLength(1)
  })

  it('surfaces a rejected edit as a tool error the model can read', async () => {
    const { adapter } = fakeAdapter({
      applyEdits: vi.fn(async () => ({ ok: false, summary: 'Paragraph 99 does not exist.' })),
    })
    const out = await executeDocumentTool(
      'replace_paragraph',
      '{"paragraph":99,"text":"x"}',
      ctx('direct', adapter).toolCtx,
    )
    expect(out.isError).toBe(true)
    expect(out.content).toContain('does not exist')
  })
})

describe('argument handling', () => {
  it('defaults insert_text to after the selection', async () => {
    const { adapter } = fakeAdapter()
    const c = ctx('propose', adapter)
    await executeDocumentTool('insert_text', '{"text":"tail"}', c.toolCtx)
    expect(c.proposed[0].edit).toEqual({ kind: 'insertAfter', text: 'tail' })
  })

  it('honours an explicit before position', async () => {
    const { adapter } = fakeAdapter()
    const c = ctx('propose', adapter)
    await executeDocumentTool('insert_text', '{"text":"head","position":"before"}', c.toolCtx)
    expect(c.proposed[0].edit).toEqual({ kind: 'insertBefore', text: 'head' })
  })

  // A formula without the leading = writes a text cell that looks right and
  // computes nothing.
  it('adds a missing leading = to formulas', async () => {
    const { adapter } = fakeAdapter()
    const c = ctx('propose', adapter)
    await executeDocumentTool(
      'write_formula',
      '{"cells":[{"cell":"F2","formula":"SUM(A1:A9)"}]}',
      c.toolCtx,
    )
    expect(c.proposed[0].edit).toEqual({
      kind: 'setCellFormulas',
      sheet: '',
      cells: [{ cell: 'F2', formula: '=SUM(A1:A9)' }],
    })
  })

  it('rejects malformed JSON arguments without throwing', async () => {
    const { adapter } = fakeAdapter()
    const out = await executeDocumentTool(
      'replace_selection',
      'not json',
      ctx('direct', adapter).toolCtx,
    )
    expect(out.isError).toBe(true)
  })

  it('rejects a write_range with no cells', async () => {
    const { adapter } = fakeAdapter()
    const out = await executeDocumentTool(
      'write_range',
      '{"cells":[]}',
      ctx('direct', adapter).toolCtx,
    )
    expect(out.isError).toBe(true)
  })

  it('reports an unknown tool rather than failing silently', async () => {
    const { adapter } = fakeAdapter()
    const out = await executeDocumentTool('do_magic', '{}', ctx('direct', adapter).toolCtx)
    expect(out.isError).toBe(true)
    expect(out.content).toContain('Unknown tool')
  })

  // An adapter throwing must not abort the whole turn.
  it('converts an adapter exception into a tool error', async () => {
    const { adapter } = fakeAdapter({
      getContext: vi.fn(async () => {
        throw new Error('Word is busy')
      }),
    })
    const out = await executeDocumentTool('read_document', '{}', ctx('direct', adapter).toolCtx)
    expect(out.isError).toBe(true)
    expect(out.content).toContain('Word is busy')
  })
})

/** Run a Word tool in propose mode and hand back the edit it staged. */
async function stage(tool: string, args: unknown) {
  const { adapter } = fakeAdapter()
  const c = ctx('propose', adapter)
  const out = await executeDocumentTool(tool, JSON.stringify(args), c.toolCtx)
  return { out, edit: c.proposed[0]?.edit, description: c.proposed[0]?.description }
}

describe('format_text targeting', () => {
  it('defaults to the selection', async () => {
    const { edit } = await stage('format_text', { bold: true })
    expect(edit).toEqual({
      kind: 'formatText',
      target: { kind: 'selection' },
      formatting: { bold: true },
    })
  })

  it('takes a single paragraph number', async () => {
    const { edit } = await stage('format_text', { paragraph: 3, italic: true })
    expect(edit).toMatchObject({ target: { kind: 'paragraphs', paragraphs: [3] } })
  })

  it('expands an inclusive from/to range', async () => {
    const { edit } = await stage('format_text', { from: 2, to: 5, bold: true })
    expect(edit).toMatchObject({ target: { paragraphs: [2, 3, 4, 5] } })
  })

  it('takes an explicit document scope, under either name', async () => {
    for (const args of [
      { scope: 'document', font: 'Calibri' },
      { target: 'document', font: 'Calibri' },
    ]) {
      const { edit } = await stage('format_text', args)
      expect(edit).toMatchObject({ target: { kind: 'document' } })
    }
  })

  // A phrase is the only way to reach a run of text inside a paragraph, so it
  // wins over a paragraph number sent alongside it.
  it('prefers a phrase over a paragraph number when both are given', async () => {
    const { edit } = await stage('format_text', { find: 'budget', paragraph: 3, bold: true })
    expect(edit).toMatchObject({ target: { kind: 'search', search: 'budget' } })
  })

  it('carries the search modifiers through', async () => {
    const { edit } = await stage('format_text', {
      find: 'IT',
      wholeWord: true,
      matchCase: true,
      bold: true,
    })
    expect(edit).toMatchObject({
      target: { kind: 'search', wholeWord: true, matchCase: true, firstOnly: false },
    })
  })

  // Models fill unused parameters with explicit nulls, and `Number(null)` is 0
  // — so a null paragraph would quietly retarget the edit onto paragraph 0.
  it('treats an explicit null parameter as absent, not as zero', async () => {
    const { edit } = await stage('format_text', {
      paragraph: null,
      paragraphs: null,
      from: 3,
      to: 4,
      bold: true,
    })
    expect(edit).toMatchObject({ target: { paragraphs: [3, 4] } })
  })

  it('does not mistake a targeting argument for a formatting property', async () => {
    const { out, edit } = await stage('format_text', { paragraph: 1, wholeWord: true, bold: true })
    expect(edit).toMatchObject({ formatting: { bold: true } })
    expect(out.content).not.toContain('Not applied')
  })

  // A model that sends replacement text to a formatting tool is confused, and
  // dropping the text without a word is how it stays confused.
  it('tells the model that format_text cannot change the text itself', async () => {
    const { out } = await stage('format_text', { paragraph: 1, bold: true, text: 'new wording' })
    expect(out.content).toContain('Not applied')
    expect(out.content).toContain('text')
  })

  it('describes the target in the proposal card', async () => {
    const { description } = await stage('format_text', { from: 1, to: 2, bold: true })
    expect(description).toBe('Format paragraphs 1, 2')
  })
})

describe('format_text reporting', () => {
  // Reporting success for an edit that changed nothing is what makes
  // formatting look broken to the user.
  it('refuses a call with no property to change', async () => {
    const { out } = await stage('format_text', { paragraph: 2 })
    expect(out.isError).toBe(true)
    expect(out.content).toContain('at least one property')
  })

  it('names a property it cannot support while applying the rest', async () => {
    const { out, edit } = await stage('format_text', { paragraph: 2, bold: true, border: 'thin' })
    expect(edit).toMatchObject({ formatting: { bold: true } })
    expect(out.content).toContain('Not applied')
    expect(out.content).toContain('border')
  })

  it('fails the call outright when nothing survives validation', async () => {
    const { out } = await stage('format_text', { paragraph: 2, border: 'thin' })
    expect(out.isError).toBe(true)
    expect(out.content).toContain('Nothing was changed')
  })

  it('coerces the shapes models actually emit', async () => {
    const { edit } = await stage('format_text', { paragraph: 0, fontSize: '18', bold: 'true' })
    expect(edit).toMatchObject({ formatting: { size: 18, bold: true } })
  })
})

describe('Word structure tools', () => {
  it('defaults an insertion to the end of the document', async () => {
    const { edit } = await stage('insert_paragraph', { text: 'Appendix' })
    expect(edit).toEqual({
      kind: 'insertParagraph',
      at: 'end',
      text: 'Appendix',
      style: undefined,
    })
  })

  it('accepts a paragraph number, a keyword, or a numeric string as the anchor', async () => {
    expect((await stage('insert_paragraph', { text: 'x', after: 4 })).edit).toMatchObject({ at: 4 })
    expect((await stage('insert_paragraph', { text: 'x', after: '4' })).edit).toMatchObject({
      at: 4,
    })
    expect((await stage('insert_paragraph', { text: 'x', after: 'start' })).edit).toMatchObject({
      at: 'start',
    })
  })

  it('rejects an anchor it cannot make sense of', async () => {
    const { out } = await stage('insert_paragraph', { text: 'x', after: 'somewhere nice' })
    expect(out.isError).toBe(true)
    expect(out.content).toContain('Unknown position')
  })

  it('allows a blank paragraph but not a missing one', async () => {
    expect((await stage('insert_paragraph', { text: '' })).out.isError).toBe(false)
    expect((await stage('insert_paragraph', {})).out.isError).toBe(true)
  })

  it('deletes a single paragraph or a range', async () => {
    expect((await stage('delete_paragraph', { paragraph: 7 })).edit).toEqual({
      kind: 'deleteParagraphs',
      paragraphs: [7],
    })
    expect((await stage('delete_paragraph', { from: 2, to: 4 })).edit).toMatchObject({
      paragraphs: [2, 3, 4],
    })
  })

  // "bullet points 2 and 5" is a normal request that a from/to pair cannot
  // express, and a model given only a range simply declines to call the tool.
  it('lists paragraphs that are not next to each other', async () => {
    const { edit, out } = await stage('set_list', { paragraphs: [2, 4], type: 'bullet' })
    expect(out.isError).toBe(false)
    expect(edit).toMatchObject({ kind: 'setList', paragraphs: [2, 4], listType: 'bullet' })
  })

  it('refuses a list with no paragraphs to act on', async () => {
    const { out } = await stage('set_list', { type: 'bullet' })
    expect(out.isError).toBe(true)
    expect(out.content).toContain('paragraph numbers')
  })

  it('accepts the list-type synonyms models reach for', async () => {
    expect((await stage('set_list', { from: 1, to: 3, type: 'bulleted' })).edit).toMatchObject({
      listType: 'bullet',
      paragraphs: [1, 2, 3],
    })
    expect((await stage('set_list', { from: 1, to: 3, type: 'ordered' })).edit).toMatchObject({
      listType: 'number',
    })
    expect((await stage('set_list', { from: 1, to: 3, type: 'wombat' })).out.isError).toBe(true)
  })

  it('takes a table as rows of cells', async () => {
    const { edit } = await stage('insert_table', {
      rows: [
        ['Region', 'Revenue'],
        ['North', 1200],
      ],
    })
    expect(edit).toMatchObject({
      kind: 'insertTable',
      headerRow: true,
      rows: [
        ['Region', 'Revenue'],
        ['North', '1200'],
      ],
    })
  })

  // Asked for a table of data, models emit an array of objects about as often
  // as an array of arrays.
  it('turns an array of objects into a header row plus values', async () => {
    const { edit } = await stage('insert_table', {
      rows: [
        { Region: 'North', Revenue: 1200 },
        { Region: 'South', Revenue: 900 },
      ],
    })
    expect(edit).toMatchObject({
      rows: [
        ['Region', 'Revenue'],
        ['North', '1200'],
        ['South', '900'],
      ],
    })
  })

  it('rejects a table with no rows', async () => {
    expect((await stage('insert_table', { rows: [] })).out.isError).toBe(true)
  })

  it('defaults a break to a page break', async () => {
    expect((await stage('insert_break', { type: 'nonsense' })).edit).toMatchObject({
      breakType: 'page',
    })
  })

  it('links the selection by default and a phrase on request', async () => {
    expect((await stage('insert_hyperlink', { url: 'https://x.test' })).edit).toMatchObject({
      target: { kind: 'selection' },
      url: 'https://x.test',
    })
    // `text` is the link label, not a search target.
    const labelled = await stage('insert_hyperlink', {
      url: 'https://x.test',
      text: 'the policy',
      find: 'policy',
    })
    expect(labelled.edit).toMatchObject({
      target: { kind: 'search', search: 'policy' },
      text: 'the policy',
    })
  })

  it('allows an empty replacement but not a missing one', async () => {
    expect((await stage('replace_all', { find: 'x', replace: '' })).out.isError).toBe(false)
    expect((await stage('replace_all', { find: 'x' })).out.isError).toBe(true)
  })

  it('applies one margin value to all four edges', async () => {
    const { edit } = await stage('set_page_setup', { margins: 1 })
    expect(edit).toMatchObject({
      setup: { margins: { top: 1, bottom: 1, left: 1, right: 1 } },
    })
  })

  it('lets a named edge override the shared margin', async () => {
    const { edit } = await stage('set_page_setup', { margins: 1, marginTop: 2 })
    expect(edit).toMatchObject({ setup: { margins: { top: 2, left: 1 } } })
  })

  it('refuses a page setup that specifies nothing', async () => {
    expect((await stage('set_page_setup', {})).out.isError).toBe(true)
    // Null margins are "not specified", not "zero-inch margins".
    const nulls = await stage('set_page_setup', { margins: null, marginTop: null })
    expect(nulls.out.isError).toBe(true)
  })

  it('does not turn a null margin into a zero margin', async () => {
    const { edit } = await stage('set_page_setup', { orientation: 'landscape', marginTop: null })
    expect(edit).toEqual({
      kind: 'setPageSetup',
      setup: { orientation: 'landscape', pageSize: undefined },
    })
  })

  it('rejects a paper size Word does not have', async () => {
    const { out } = await stage('set_page_setup', { pageSize: 'a9' })
    expect(out.isError).toBe(true)
    expect(out.content).toContain('a4')
  })
})

describe('formatting tools in the catalog', () => {
  it('offers the full Word formatting surface', () => {
    const names = toolCatalog('word', true).map((t) => t.name)
    for (const tool of [
      'format_text',
      'insert_paragraph',
      'delete_paragraph',
      'set_list',
      'insert_table',
      'insert_break',
      'insert_hyperlink',
      'replace_all',
      'set_header_footer',
      'set_page_setup',
      'read_formatting',
    ]) {
      expect(names).toContain(tool)
    }
  })

  it('keeps the formatting tools out of Excel', () => {
    const names = toolCatalog('excel', true).map((t) => t.name)
    expect(names).not.toContain('format_text')
    expect(names).toContain('format_range')
  })

  it('withholds every new write tool when writes are not allowed', () => {
    const names = toolCatalog('word', false).map((t) => t.name)
    expect(names).not.toContain('format_text')
    expect(names).not.toContain('set_page_setup')
    // Reading formatting is not a write.
    expect(names).toContain('read_formatting')
  })

  it('still counts the retired format_paragraph as a write tool', () => {
    expect(isWriteTool('format_paragraph')).toBe(true)
    expect(isWriteTool('format_text')).toBe(true)
    expect(isWriteTool('read_formatting')).toBe(false)
  })

  // Transcripts persisted by an earlier build replay through the old name.
  it('still dispatches the retired format_paragraph', async () => {
    const { out, edit } = await stage('format_paragraph', { paragraph: 2, bold: true })
    expect(out.isError).toBe(false)
    expect(edit).toMatchObject({
      kind: 'formatText',
      target: { kind: 'paragraphs', paragraphs: [2] },
    })
  })
})

describe('read_formatting', () => {
  it('asks the adapter and returns its report', async () => {
    const { adapter } = fakeAdapter({
      readFormatting: vi.fn(async () => 'Bold: true'),
    })
    const out = await executeDocumentTool(
      'read_formatting',
      '{"paragraph":1}',
      ctx('propose', adapter).toolCtx,
    )
    expect(out.isError).toBe(false)
    expect(out.content).toBe('Bold: true')
    expect(adapter.readFormatting).toHaveBeenCalledWith({ kind: 'paragraphs', paragraphs: [1] })
  })

  // Excel's adapter has no readFormatting, so the tool has to answer rather
  // than throw on `undefined is not a function`.
  it('says so in a host that cannot read formatting', async () => {
    const { adapter } = fakeAdapter()
    expect(adapter.readFormatting).toBeUndefined()
    const out = await executeDocumentTool('read_formatting', '{}', ctx('propose', adapter).toolCtx)
    expect(out.isError).toBe(true)
  })
})
