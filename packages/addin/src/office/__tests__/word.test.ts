import { describe, it, expect, afterEach } from 'vitest'
import { WordAdapter } from '../word'
import { installWordMock, makeParagraph, uninstallOfficeMocks } from './mockOffice'

afterEach(() => uninstallOfficeMocks())

function sampleDoc() {
  return [
    makeParagraph('Quarterly Report', 'Title'),
    makeParagraph('Executive Summary', 'Heading 1'),
    makeParagraph('Revenue rose eleven percent.', 'Normal'),
    makeParagraph('Risks', 'Heading 1'),
    makeParagraph('Freight contracts renew in January.', 'Normal'),
  ]
}

describe('WordAdapter.getContext', () => {
  it('returns the selection when there is one', async () => {
    installWordMock({ paragraphs: sampleDoc(), selectionText: 'Revenue rose eleven percent.' })
    const ctx = await new WordAdapter().getContext('selection')
    expect(ctx.scope).toBe('selection')
    expect(ctx.text).toBe('Revenue rose eleven percent.')
  })

  // The single most common case: the user clicks into the pane and types
  // without selecting anything. Answering about an empty string here is what
  // makes an assistant look like it cannot see the document.
  it('falls back to the whole document when nothing is selected', async () => {
    installWordMock({ paragraphs: sampleDoc(), selectionText: '   ' })
    const ctx = await new WordAdapter().getContext('selection')
    expect(ctx.scope).toBe('document')
    expect(ctx.text).toContain('Freight contracts')
  })

  it('builds an outline from heading paragraphs only', async () => {
    installWordMock({ paragraphs: sampleDoc() })
    const ctx = await new WordAdapter().getContext('document')
    expect(ctx.outline).toContain('Quarterly Report')
    expect(ctx.outline).toContain('Executive Summary')
    // Body text must not appear in the outline.
    expect(ctx.outline).not.toContain('Revenue rose')
  })

  it('numbers outline entries so the model can reference them in edits', async () => {
    installWordMock({ paragraphs: sampleDoc() })
    const ctx = await new WordAdapter().getContext('document')
    expect(ctx.outline).toMatch(/\[¶1\] Executive Summary/)
  })

  it('reports an empty document without throwing', async () => {
    installWordMock({ paragraphs: [] })
    const ctx = await new WordAdapter().getContext('document')
    expect(ctx.text).toBe('')
  })

  it('returns nothing for scope none', async () => {
    installWordMock({ paragraphs: sampleDoc() })
    const ctx = await new WordAdapter().getContext('none')
    expect(ctx.text).toBe('')
    expect(ctx.tokenEstimate).toBe(0)
  })
})

describe('WordAdapter.applyEdits', () => {
  it('replaces the selection in place rather than rewriting the body', async () => {
    const state = installWordMock({ paragraphs: sampleDoc(), selectionText: 'old' })
    const result = await new WordAdapter().applyEdits([
      { kind: 'replaceSelection', text: 'new text' },
    ])
    expect(result.ok).toBe(true)
    expect(state.selectionEdits).toEqual([{ text: 'new text', location: 'Replace' }])
  })

  it('inserts before and after without replacing', async () => {
    const state = installWordMock({ paragraphs: sampleDoc(), selectionText: 'x' })
    await new WordAdapter().applyEdits([
      { kind: 'insertBefore', text: 'B' },
      { kind: 'insertAfter', text: 'A' },
    ])
    expect(state.selectionEdits.map((e) => e.location)).toEqual(['Before', 'After'])
  })

  it('replaces a paragraph by index', async () => {
    const state = installWordMock({ paragraphs: sampleDoc() })
    const result = await new WordAdapter().applyEdits([
      { kind: 'replaceRange', rangeId: '2', text: 'Revenue rose twelve percent.' },
    ])
    expect(result.ok).toBe(true)
    expect(state.paragraphs[2].text).toBe('Revenue rose twelve percent.')
    // Neighbours untouched.
    expect(state.paragraphs[1].text).toBe('Executive Summary')
  })

  // A hallucinated index must come back as a readable failure the model can
  // act on, not a thrown error that kills the turn.
  it('rejects an out-of-range paragraph with a useful message', async () => {
    installWordMock({ paragraphs: sampleDoc() })
    const result = await new WordAdapter().applyEdits([
      { kind: 'replaceRange', rangeId: '99', text: 'nope' },
    ])
    expect(result.ok).toBe(false)
    expect(result.summary).toContain('does not exist')
  })

  it('rejects a non-numeric paragraph reference', async () => {
    installWordMock({ paragraphs: sampleDoc() })
    const result = await new WordAdapter().applyEdits([
      { kind: 'replaceRange', rangeId: 'the intro', text: 'nope' },
    ])
    expect(result.ok).toBe(false)
    expect(result.summary).toContain('Invalid paragraph reference')
  })

  it('adds a comment without changing the text', async () => {
    const state = installWordMock({ paragraphs: sampleDoc() })
    const result = await new WordAdapter().applyEdits([
      { kind: 'addComment', rangeId: '4', text: 'Check this figure.' },
    ])
    expect(result.ok).toBe(true)
    expect(state.paragraphs[4].comments).toEqual(['Check this figure.'])
    expect(state.paragraphs[4].text).toBe('Freight contracts renew in January.')
  })

  it('rejects an Excel-only edit kind', async () => {
    installWordMock({ paragraphs: sampleDoc() })
    const result = await new WordAdapter().applyEdits([
      { kind: 'setCellValues', sheet: 'Sales', cells: [{ cell: 'A1', value: 1 }] },
    ])
    expect(result.ok).toBe(false)
    expect(result.summary).toContain('not supported in Word')
  })

  it('is a no-op for an empty edit list', async () => {
    const state = installWordMock({ paragraphs: sampleDoc() })
    const result = await new WordAdapter().applyEdits([])
    expect(result.ok).toBe(true)
    expect(state.selectionEdits).toHaveLength(0)
  })
})

describe('WordAdapter formatting', () => {
  it('applies character formatting to numbered paragraphs', async () => {
    const state = installWordMock({ paragraphs: sampleDoc() })
    const result = await new WordAdapter().applyEdits([
      {
        kind: 'formatText',
        target: { kind: 'paragraphs', paragraphs: [2, 4] },
        formatting: { bold: true, size: 14, color: '#C00000' },
      },
    ])
    expect(result.ok).toBe(true)
    expect(state.paragraphs[2].font).toMatchObject({ bold: true, size: 14, color: '#C00000' })
    expect(state.paragraphs[4].font).toMatchObject({ bold: true })
    // Untargeted paragraphs are untouched.
    expect(state.paragraphs[3].font).toEqual({})
  })

  // The bug this whole path exists to prevent: `paragraph.style` takes the
  // localized display name and throws on anything else, so a style has to go
  // through `styleBuiltIn` to survive a non-English Word.
  it('sets a built-in style through the locale-neutral setter', async () => {
    const state = installWordMock({ paragraphs: sampleDoc() })
    await new WordAdapter().applyEdits([
      {
        kind: 'formatText',
        target: { kind: 'paragraphs', paragraphs: [2] },
        formatting: { style: 'Heading 2' },
      },
    ])
    expect(state.paragraphs[2].styleBuiltIn).toBe('Heading2')
  })

  it('falls back to the plain style setter for a document-specific style', async () => {
    const state = installWordMock({ paragraphs: sampleDoc() })
    await new WordAdapter().applyEdits([
      {
        kind: 'formatText',
        target: { kind: 'paragraphs', paragraphs: [2] },
        formatting: { style: 'Acme Body' },
      },
    ])
    expect(state.paragraphs[2].style).toBe('Acme Body')
    expect(state.paragraphs[2].styleBuiltIn).toBeUndefined()
  })

  it('applies paragraph properties as well as font properties', async () => {
    const state = installWordMock({ paragraphs: sampleDoc() })
    await new WordAdapter().applyEdits([
      {
        kind: 'formatText',
        target: { kind: 'paragraphs', paragraphs: [0] },
        formatting: { alignment: 'center', spaceAfter: 12, leftIndent: 36, italic: true },
      },
    ])
    expect(state.paragraphs[0].props).toEqual({
      alignment: 'Centered',
      spaceAfter: 12,
      leftIndent: 36,
    })
    expect(state.paragraphs[0].font).toMatchObject({ italic: true })
  })

  it('formats every occurrence of a phrase, not the whole paragraph', async () => {
    const state = installWordMock({
      paragraphs: [makeParagraph('The budget is fixed. The budget holds.')],
    })
    const result = await new WordAdapter().applyEdits([
      {
        kind: 'formatText',
        target: { kind: 'search', search: 'budget' },
        formatting: { bold: true },
      },
    ])
    expect(result.summary).toContain('2 occurrences')
    const hits = state.ranges.filter((r) => r.text === 'budget')
    expect(hits).toHaveLength(2)
    expect(hits.every((h) => h.font.bold === true)).toBe(true)
    // The paragraph itself was not blanket-bolded.
    expect(state.paragraphs[0].font.bold).toBeUndefined()
  })

  it('reports a phrase that is not in the document instead of claiming success', async () => {
    installWordMock({ paragraphs: sampleDoc() })
    const result = await new WordAdapter().applyEdits([
      {
        kind: 'formatText',
        target: { kind: 'search', search: 'ebitda' },
        formatting: { bold: true },
      },
    ])
    expect(result.ok).toBe(false)
    expect(result.summary).toContain('No text matching')
  })

  // Nothing selected is an insertion point, and Word applies paragraph
  // formatting to the paragraph holding the cursor. Doing nothing and
  // reporting success is the failure mode to avoid.
  it('falls back to the paragraph at the cursor when nothing is selected', async () => {
    const state = installWordMock({ paragraphs: sampleDoc(), selectionText: '' })
    const result = await new WordAdapter().applyEdits([
      { kind: 'formatText', target: { kind: 'selection' }, formatting: { alignment: 'right' } },
    ])
    expect(result.summary).toContain('paragraph at the cursor')
    expect(state.paragraphs[0].props.alignment).toBe('Right')
  })

  it('rejects a paragraph number the document does not have', async () => {
    installWordMock({ paragraphs: sampleDoc() })
    const result = await new WordAdapter().applyEdits([
      {
        kind: 'formatText',
        target: { kind: 'paragraphs', paragraphs: [99] },
        formatting: { bold: true },
      },
    ])
    expect(result.ok).toBe(false)
    expect(result.summary).toContain('does not exist')
  })

  // Proposals persisted by an earlier build carry the old edit shape.
  it('still applies the pre-format_text edit shape', async () => {
    const state = installWordMock({ paragraphs: sampleDoc() })
    const result = await new WordAdapter().applyEdits([
      { kind: 'applyFormatting', rangeId: '1', formatting: { bold: true, style: 'Heading 3' } },
    ])
    expect(result.ok).toBe(true)
    expect(state.paragraphs[1].font.bold).toBe(true)
    expect(state.paragraphs[1].styleBuiltIn).toBe('Heading3')
  })

  it('reads back the formatting in force', async () => {
    const paragraphs = sampleDoc()
    paragraphs[1].font.bold = true
    installWordMock({ paragraphs })
    const report = await new WordAdapter().readFormatting({ kind: 'paragraphs', paragraphs: [1] })
    expect(report).toContain('Bold: true')
    expect(report).toContain('Executive Summary')
  })
})

describe('WordAdapter structure edits', () => {
  it('inserts a paragraph after a numbered one and styles it', async () => {
    const state = installWordMock({ paragraphs: sampleDoc() })
    const result = await new WordAdapter().applyEdits([
      { kind: 'insertParagraph', at: 1, text: 'Outlook', style: 'Heading 1' },
    ])
    expect(result.ok).toBe(true)
    expect(state.paragraphs[2].text).toBe('Outlook')
    expect(state.paragraphs[2].styleBuiltIn).toBe('Heading1')
  })

  it('appends at the end and prepends at the start', async () => {
    const state = installWordMock({ paragraphs: sampleDoc() })
    await new WordAdapter().applyEdits([
      { kind: 'insertParagraph', at: 'end', text: 'Last' },
      { kind: 'insertParagraph', at: 'start', text: 'First' },
    ])
    expect(state.paragraphs[0].text).toBe('First')
    expect(state.paragraphs[state.paragraphs.length - 1].text).toBe('Last')
  })

  // Deleting front-to-back would renumber the rest mid-loop and delete the
  // wrong paragraphs.
  it('deletes a range without the indices sliding underneath it', async () => {
    const state = installWordMock({ paragraphs: sampleDoc() })
    const result = await new WordAdapter().applyEdits([
      { kind: 'deleteParagraphs', paragraphs: [1, 2, 3] },
    ])
    expect(result.ok).toBe(true)
    expect(state.paragraphs.map((p) => p.text)).toEqual([
      'Quarterly Report',
      'Freight contracts renew in January.',
    ])
  })

  it('refuses a delete range that runs off the end', async () => {
    installWordMock({ paragraphs: sampleDoc() })
    const result = await new WordAdapter().applyEdits([
      { kind: 'deleteParagraphs', paragraphs: [3, 40] },
    ])
    expect(result.ok).toBe(false)
    expect(result.summary).toContain('has 5')
  })

  it('makes a range of paragraphs a bulleted list', async () => {
    const state = installWordMock({ paragraphs: sampleDoc() })
    const result = await new WordAdapter().applyEdits([
      { kind: 'setList', paragraphs: [2, 3, 4], listType: 'bullet' },
    ])
    expect(result.ok).toBe(true)
    expect(state.lists).toEqual([{ id: 1, level: 0, kind: 'bullet', bullet: 'Solid' }])
    // Every paragraph in the range joins the same list.
    expect(state.paragraphs.slice(2, 5).every((p) => p.list?.id === 1)).toBe(true)
  })

  it('numbers a list at the level asked for', async () => {
    const state = installWordMock({ paragraphs: sampleDoc() })
    await new WordAdapter().applyEdits([
      { kind: 'setList', paragraphs: [2, 3], listType: 'number', level: 2 },
    ])
    expect(state.lists[0]).toMatchObject({ kind: 'number', level: 1 })
  })

  it('strips list formatting', async () => {
    const state = installWordMock({ paragraphs: sampleDoc() })
    await new WordAdapter().applyEdits([{ kind: 'setList', paragraphs: [0, 1], listType: 'none' }])
    expect(state.paragraphs[0].detachedFromList).toBe(true)
    expect(state.paragraphs[1].detachedFromList).toBe(true)
  })

  it('inserts a table with a header row and a style', async () => {
    const state = installWordMock({ paragraphs: sampleDoc() })
    const result = await new WordAdapter().applyEdits([
      {
        kind: 'insertTable',
        at: 'end',
        rows: [
          ['Region', 'Revenue'],
          ['North', '1200'],
        ],
        headerRow: true,
      },
    ])
    expect(result.ok).toBe(true)
    expect(state.tables[0]).toMatchObject({
      rowCount: 2,
      columnCount: 2,
      headerRowCount: 1,
      styleBuiltIn: 'GridTable4_Accent1',
    })
  })

  // Word rejects a values grid that is not exactly rowCount × columnCount.
  it('pads ragged rows to a rectangle', async () => {
    const state = installWordMock({ paragraphs: sampleDoc() })
    await new WordAdapter().applyEdits([
      { kind: 'insertTable', at: 'end', rows: [['a', 'b', 'c'], ['d']] },
    ])
    expect(state.tables[0].values).toEqual([
      ['a', 'b', 'c'],
      ['d', '', ''],
    ])
  })

  it('inserts breaks', async () => {
    const state = installWordMock({ paragraphs: sampleDoc() })
    await new WordAdapter().applyEdits([
      { kind: 'insertBreak', at: 2, breakType: 'page' },
      { kind: 'insertBreak', at: 'end', breakType: 'section' },
    ])
    expect(state.breaks).toEqual([
      { type: 'Page', location: 'After' },
      { type: 'SectionNext', location: 'End' },
    ])
  })

  it('links text and gives it the hyperlink style', async () => {
    const state = installWordMock({
      paragraphs: [makeParagraph('See the handbook for details.')],
    })
    const result = await new WordAdapter().applyEdits([
      {
        kind: 'insertHyperlink',
        target: { kind: 'search', search: 'handbook' },
        url: 'https://example.com/handbook',
      },
    ])
    expect(result.ok).toBe(true)
    expect(state.hyperlinks).toEqual([{ text: 'handbook', url: 'https://example.com/handbook' }])
    expect(state.ranges.find((r) => r.text === 'handbook')?.styleBuiltIn).toBe('Hyperlink')
  })

  it('refuses something that is not a link', async () => {
    installWordMock({ paragraphs: sampleDoc() })
    const result = await new WordAdapter().applyEdits([
      { kind: 'insertHyperlink', target: { kind: 'selection' }, url: 'javascript:alert(1)' },
    ])
    expect(result.ok).toBe(false)
    expect(result.summary).toContain('is not a link')
  })

  it('replaces every occurrence and counts them', async () => {
    const state = installWordMock({
      paragraphs: [makeParagraph('FY24 was strong.'), makeParagraph('FY24 closed early.')],
    })
    const result = await new WordAdapter().applyEdits([
      { kind: 'replaceAll', find: 'FY24', replace: 'FY2024' },
    ])
    expect(result.summary).toContain('2 occurrences')
    expect(state.paragraphs.map((p) => p.text)).toEqual([
      'FY2024 was strong.',
      'FY2024 closed early.',
    ])
  })

  it('says so when there is nothing to replace', async () => {
    installWordMock({ paragraphs: sampleDoc() })
    const result = await new WordAdapter().applyEdits([
      { kind: 'replaceAll', find: 'FY24', replace: 'FY2024' },
    ])
    expect(result.ok).toBe(false)
    expect(result.summary).toContain('No occurrences')
  })

  it('sets a footer with an automatic page number', async () => {
    const state = installWordMock({ paragraphs: sampleDoc() })
    const result = await new WordAdapter().applyEdits([
      {
        kind: 'setHeaderFooter',
        part: 'footer',
        text: 'Quarterly Report',
        alignment: 'center',
        pageNumber: true,
      },
    ])
    expect(result.ok).toBe(true)
    expect(state.headerFooters[0]).toMatchObject({
      part: 'footer',
      cleared: true,
      alignment: 'Centered',
      fields: ['Page'],
    })
    expect(state.headerFooters[0].text).toContain('Quarterly Report')
  })

  it('changes page setup through the document OOXML', async () => {
    const state = installWordMock({
      paragraphs: sampleDoc(),
      ooxml:
        '<w:body><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440"/></w:sectPr></w:body>',
    })
    const result = await new WordAdapter().applyEdits([
      { kind: 'setPageSetup', setup: { orientation: 'landscape', margins: { top: 0.5 } } },
    ])
    expect(result.ok).toBe(true)
    expect(state.ooxml).toContain('w:orient="landscape"')
    expect(state.ooxml).toContain('w:top="720"')
  })

  it('reports a page-setup change the document cannot take', async () => {
    installWordMock({ paragraphs: sampleDoc(), ooxml: '<w:body><w:p/></w:body>' })
    const result = await new WordAdapter().applyEdits([
      { kind: 'setPageSetup', setup: { orientation: 'landscape' } },
    ])
    expect(result.ok).toBe(false)
    expect(result.summary).toContain('does not expose page settings')
  })

  // A batch that fails halfway must say which step failed, not report success.
  it('stops the batch at the first failing edit', async () => {
    const state = installWordMock({ paragraphs: sampleDoc() })
    const result = await new WordAdapter().applyEdits([
      {
        kind: 'formatText',
        target: { kind: 'paragraphs', paragraphs: [0] },
        formatting: { bold: true },
      },
      { kind: 'deleteParagraphs', paragraphs: [0, 99] },
    ])
    expect(result.ok).toBe(false)
    expect(state.paragraphs).toHaveLength(5)
  })
})

describe('WordAdapter.search', () => {
  it('returns paragraph indices usable with replace_paragraph', async () => {
    installWordMock({ paragraphs: sampleDoc() })
    const hits = await new WordAdapter().search('freight')
    expect(hits).toHaveLength(1)
    expect(hits[0].location).toBe('4')
  })

  it('matches case-insensitively and returns nothing for a miss', async () => {
    installWordMock({ paragraphs: sampleDoc() })
    expect(await new WordAdapter().search('REVENUE')).toHaveLength(1)
    expect(await new WordAdapter().search('zzz')).toHaveLength(0)
  })
})

describe('WordAdapter snapshot/restore', () => {
  it('round-trips OOXML', async () => {
    const state = installWordMock({
      paragraphs: sampleDoc(),
      ooxml: '<w:document>ORIGINAL</w:document>',
    })
    const adapter = new WordAdapter()
    const snap = await adapter.snapshot()
    state.ooxml = '<w:document>CHANGED</w:document>'
    await adapter.restore(snap)
    expect(state.ooxml).toBe('<w:document>ORIGINAL</w:document>')
  })

  // Silently replacing a formatted document with plain text would be worse
  // than refusing, so an oversized snapshot must fail loudly on restore.
  it('refuses to restore a snapshot that was too large to capture', async () => {
    installWordMock({ paragraphs: [makeParagraph('x'.repeat(500_000))] })
    const adapter = new WordAdapter()
    const snap = await adapter.snapshot()
    await expect(adapter.restore(snap)).rejects.toThrow(/too large/i)
  })
})
