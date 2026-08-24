import { describe, it, expect } from 'vitest'
import {
  describeFormatting,
  needsParagraphProps,
  normalizeFormatting,
  patchPageSetup,
  resolveAlignment,
  resolveBreak,
  resolveColor,
  resolveStyle,
  resolveUnderline,
  wordAlignment,
} from '../wordFormat'

describe('resolveStyle', () => {
  // `paragraph.style` wants the localized display name and throws on anything
  // else; `styleBuiltIn` is the same on every Word UI language. Everything the
  // model is likely to say has to land in the second bucket.
  it('maps every spelling of a heading onto the locale-neutral id', () => {
    for (const name of ['Heading 1', 'heading1', 'HEADING  1', 'h1', 'Heading_1']) {
      expect(resolveStyle(name)).toEqual({ builtIn: 'Heading1' })
    }
  })

  it('maps prose names onto real styles', () => {
    expect(resolveStyle('Title')).toEqual({ builtIn: 'Title' })
    expect(resolveStyle('block quote')).toEqual({ builtIn: 'Quote' })
    expect(resolveStyle('body text')).toEqual({ builtIn: 'Normal' })
    expect(resolveStyle('Intense Quote')).toEqual({ builtIn: 'IntenseQuote' })
    expect(resolveStyle('Grid Table 4 Accent 1')).toEqual({ builtIn: 'GridTable4_Accent1' })
  })

  // Documents define their own styles and the model cannot enumerate them, so
  // an unrecognised name is a custom style, not an error.
  it('passes an unrecognised name through as a custom style', () => {
    expect(resolveStyle('Acme Body Copy')).toEqual({ custom: 'Acme Body Copy' })
  })
})

describe('resolveColor', () => {
  it('accepts hex with or without the hash', () => {
    expect(resolveColor('#c00000')).toBe('#C00000')
    expect(resolveColor('C00000')).toBe('#C00000')
  })

  it('expands three-digit shorthand', () => {
    expect(resolveColor('#f00')).toBe('#FF0000')
  })

  it('accepts colour names Word understands', () => {
    expect(resolveColor('red')).toBe('red')
  })

  it('rejects anything else rather than handing Word a value it will throw on', () => {
    expect(resolveColor('rgb(255,0,0)')).toBeNull()
    expect(resolveColor('#12')).toBeNull()
    expect(resolveColor('')).toBeNull()
  })
})

describe('value vocabularies', () => {
  it('resolves underline styles and booleans', () => {
    expect(resolveUnderline(true)).toBe('Single')
    expect(resolveUnderline(false)).toBe('None')
    expect(resolveUnderline('double')).toBe('Double')
    expect(resolveUnderline('wavy')).toBe('Wave')
    expect(resolveUnderline('sparkly')).toBeNull()
  })

  // normalizeFormatting stores the lowercase name, so the adapter feeds this
  // function its own previous output.
  it('round-trips its own output', () => {
    expect(resolveUnderline('double')).toBe('Double')
    expect(resolveUnderline('Double'.toLowerCase())).toBe('Double')
  })

  it('resolves alignment spellings onto the Word enum', () => {
    expect(resolveAlignment('centre')).toBe('center')
    expect(resolveAlignment('Justified')).toBe('justify')
    expect(resolveAlignment('sideways')).toBeNull()
    expect(wordAlignment('center')).toBe('Centered')
    expect(wordAlignment('justify')).toBe('Justified')
  })

  it('resolves break types', () => {
    expect(resolveBreak('page')).toBe('Page')
    expect(resolveBreak('section')).toBe('SectionNext')
    expect(resolveBreak('continuous')).toBe('SectionContinuous')
    expect(resolveBreak('column')).toBeNull()
  })
})

describe('normalizeFormatting', () => {
  it('accepts the canonical property names', () => {
    const { formatting, unknown, invalid } = normalizeFormatting({
      bold: true,
      size: 14,
      color: '#C00000',
      alignment: 'center',
    })
    expect(formatting).toEqual({ bold: true, size: 14, color: '#C00000', alignment: 'center' })
    expect(unknown).toEqual([])
    expect(invalid).toEqual([])
  })

  it('accepts the names models actually guess', () => {
    const { formatting, unknown } = normalizeFormatting({
      fontSize: '14pt',
      fontFamily: 'Calibri',
      textAlign: 'centre',
      highlight: 'yellow',
      strikethrough: 'true',
    })
    expect(formatting).toEqual({
      size: 14,
      font: 'Calibri',
      alignment: 'center',
      highlightColor: 'yellow',
      strikeThrough: true,
    })
    expect(unknown).toEqual([])
  })

  // The whole point: a property that cannot be applied must be reported, not
  // dropped. Reporting success for an edit that changed nothing is what makes
  // formatting look broken.
  it('reports properties it cannot support instead of dropping them', () => {
    const { formatting, unknown } = normalizeFormatting({ bold: true, border: '1pt solid' })
    expect(formatting).toEqual({ bold: true })
    expect(unknown).toEqual(['border'])
  })

  it('reports a known property with an unusable value', () => {
    const { formatting, invalid } = normalizeFormatting({ color: 'rgb(1,2,3)', size: -4 })
    expect(formatting).toEqual({})
    expect(invalid).toHaveLength(2)
    expect(invalid[0]).toContain('color')
  })

  it('treats null highlight as "remove the highlight"', () => {
    expect(normalizeFormatting({ highlightColor: null }).formatting).toEqual({
      highlightColor: null,
    })
  })

  it('keeps false apart from absent', () => {
    // `bold: false` must un-bold; only an omitted key means "leave alone".
    expect(normalizeFormatting({ bold: false }).formatting).toEqual({ bold: false })
    expect(normalizeFormatting({}).formatting).toEqual({})
  })

  it('knows which properties need the paragraph rather than the run', () => {
    expect(needsParagraphProps({ bold: true })).toBe(false)
    expect(needsParagraphProps({ alignment: 'left' })).toBe(true)
    expect(needsParagraphProps({ style: 'Heading1' })).toBe(true)
  })

  it('renders itself for the proposal card', () => {
    expect(describeFormatting({ bold: true, size: 14, alignment: 'center' })).toBe(
      'bold, 14pt, center-aligned',
    )
  })
})

describe('patchPageSetup', () => {
  const LETTER =
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'

  it('swaps the page dimensions when turning landscape', () => {
    const out = patchPageSetup(LETTER, { orientation: 'landscape' })
    expect('xml' in out).toBe(true)
    if (!('xml' in out)) return
    expect(out.xml).toContain('w:w="15840"')
    expect(out.xml).toContain('w:h="12240"')
    expect(out.xml).toContain('w:orient="landscape"')
    expect(out.changed).toEqual(['landscape'])
  })

  it('drops the landscape marker when turning back to portrait', () => {
    const landscape = patchPageSetup(LETTER, { orientation: 'landscape' })
    if (!('xml' in landscape)) throw new Error('expected a patch')
    const out = patchPageSetup(landscape.xml, { orientation: 'portrait' })
    if (!('xml' in out)) throw new Error('expected a patch')
    expect(out.xml).not.toContain('w:orient')
    expect(out.xml).toContain('w:w="12240"')
  })

  it('keeps the current orientation when only the paper size changes', () => {
    const landscape = patchPageSetup(LETTER, { orientation: 'landscape' })
    if (!('xml' in landscape)) throw new Error('expected a patch')
    const out = patchPageSetup(landscape.xml, { pageSize: 'a4' })
    if (!('xml' in out)) throw new Error('expected a patch')
    // A4 is 11906 × 16838 portrait, so landscape has to report the long edge.
    expect(out.xml).toContain('w:w="16838"')
    expect(out.xml).toContain('w:h="11906"')
    expect(out.xml).toContain('w:orient="landscape"')
  })

  it('converts inches to twips', () => {
    const out = patchPageSetup(LETTER, { margins: { top: 0.5, left: 2 } })
    if (!('xml' in out)) throw new Error('expected a patch')
    expect(out.xml).toContain('w:top="720"')
    expect(out.xml).toContain('w:left="2880"')
    // Untouched edges keep their value.
    expect(out.xml).toContain('w:right="1440"')
  })

  // A document written by another tool may have no <w:pgMar> at all; replacing
  // nothing while reporting new margins would be a silent no-op.
  it('creates the margin element when the document has none', () => {
    const out = patchPageSetup('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>', {
      margins: { top: 1 },
    })
    if (!('xml' in out)) throw new Error('expected a patch')
    expect(out.xml).toContain('<w:pgMar w:top="1440"/>')
  })

  it('expands a self-closing section into one that can hold the settings', () => {
    const out = patchPageSetup('<w:body><w:sectPr w:rsidR="00A"/></w:body>', {
      orientation: 'landscape',
    })
    if (!('xml' in out)) throw new Error('expected a patch')
    expect(out.xml).toContain('<w:sectPr w:rsidR="00A">')
    expect(out.xml).toContain('w:orient="landscape"')
  })

  it('patches every section and reports each change once', () => {
    const two = LETTER + LETTER
    const out = patchPageSetup(two, { orientation: 'landscape' })
    if (!('xml' in out)) throw new Error('expected a patch')
    expect(out.xml.match(/w:orient="landscape"/g)).toHaveLength(2)
    expect(out.changed).toEqual(['landscape'])
  })

  it('refuses a document with no section properties', () => {
    expect(patchPageSetup('<w:body><w:p/></w:body>', { orientation: 'landscape' })).toEqual({
      error: expect.stringContaining('does not expose page settings'),
    })
  })

  it('rejects an unusable margin rather than writing a nonsense page', () => {
    expect(patchPageSetup(LETTER, { margins: { top: -3 } })).toEqual({
      error: expect.stringContaining('between 0 and 22'),
    })
  })

  it('refuses a call that would change nothing', () => {
    expect(patchPageSetup(LETTER, {})).toEqual({
      error: expect.stringContaining('Nothing to change'),
    })
  })
})
