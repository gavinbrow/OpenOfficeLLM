// Word formatting vocabulary — the translation layer between what a model says
// and what Office.js accepts.
//
// Pure: no Office.js calls, so every rule here is unit-testable without Word.
// That matters because this is where the failure modes live. A model asked to
// make something a heading says "Heading 1", "heading1", "H1", or "Title Style";
// Word's `paragraph.style` setter wants the *localized* style name and throws a
// RichApi error on anything else. `styleBuiltIn` takes a fixed, locale-neutral
// vocabulary instead, so normalizing into that list is what makes formatting
// work on a non-English Word — and what stops a mistyped style from killing the
// whole batch.
//
// The other half of the job is refusing to lie. Silently ignoring a property the
// model asked for and then reporting "formatted paragraph 3" teaches it that the
// request worked; the user sees nothing change and concludes the feature is
// broken. Unknown and unusable properties come back as errors instead.

import type { PageSetup, TextAlignment, TextFormatting } from '@openofficellm/shared'

/** Locale-neutral style identifiers Word accepts for `styleBuiltIn`. */
export const BUILT_IN_STYLES = [
  'Normal',
  'Heading1',
  'Heading2',
  'Heading3',
  'Heading4',
  'Heading5',
  'Heading6',
  'Heading7',
  'Heading8',
  'Heading9',
  'Title',
  'Subtitle',
  'Quote',
  'IntenseQuote',
  'ListParagraph',
  'NoSpacing',
  'Caption',
  'Header',
  'Footer',
  'FootnoteText',
  'FootnoteReference',
  'EndnoteText',
  'EndnoteReference',
  'Hyperlink',
  'Strong',
  'Emphasis',
  'SubtleEmphasis',
  'IntenseEmphasis',
  'SubtleReference',
  'IntenseReference',
  'BookTitle',
  'Bibliography',
  'TocHeading',
  'Toc1',
  'Toc2',
  'Toc3',
  'Toc4',
  'Toc5',
  'Toc6',
  'Toc7',
  'Toc8',
  'Toc9',
  'TableGrid',
  'TableGridLight',
  'PlainTable1',
  'PlainTable2',
  'PlainTable3',
  'PlainTable4',
  'PlainTable5',
  'GridTable1Light',
  'GridTable1Light_Accent1',
  'GridTable1Light_Accent2',
  'GridTable1Light_Accent3',
  'GridTable1Light_Accent4',
  'GridTable1Light_Accent5',
  'GridTable1Light_Accent6',
  'GridTable2',
  'GridTable2_Accent1',
  'GridTable2_Accent2',
  'GridTable2_Accent3',
  'GridTable2_Accent4',
  'GridTable2_Accent5',
  'GridTable2_Accent6',
  'GridTable3',
  'GridTable3_Accent1',
  'GridTable3_Accent2',
  'GridTable3_Accent3',
  'GridTable3_Accent4',
  'GridTable3_Accent5',
  'GridTable3_Accent6',
  'GridTable4',
  'GridTable4_Accent1',
  'GridTable4_Accent2',
  'GridTable4_Accent3',
  'GridTable4_Accent4',
  'GridTable4_Accent5',
  'GridTable4_Accent6',
  'GridTable5Dark',
  'GridTable5Dark_Accent1',
  'GridTable5Dark_Accent2',
  'GridTable5Dark_Accent3',
  'GridTable5Dark_Accent4',
  'GridTable5Dark_Accent5',
  'GridTable5Dark_Accent6',
  'GridTable6Colorful',
  'GridTable6Colorful_Accent1',
  'GridTable6Colorful_Accent2',
  'GridTable6Colorful_Accent3',
  'GridTable6Colorful_Accent4',
  'GridTable6Colorful_Accent5',
  'GridTable6Colorful_Accent6',
  'GridTable7Colorful',
  'GridTable7Colorful_Accent1',
  'GridTable7Colorful_Accent2',
  'GridTable7Colorful_Accent3',
  'GridTable7Colorful_Accent4',
  'GridTable7Colorful_Accent5',
  'GridTable7Colorful_Accent6',
  'ListTable1Light',
  'ListTable1Light_Accent1',
  'ListTable1Light_Accent2',
  'ListTable1Light_Accent3',
  'ListTable1Light_Accent4',
  'ListTable1Light_Accent5',
  'ListTable1Light_Accent6',
  'ListTable2',
  'ListTable2_Accent1',
  'ListTable2_Accent2',
  'ListTable2_Accent3',
  'ListTable2_Accent4',
  'ListTable2_Accent5',
  'ListTable2_Accent6',
  'ListTable3',
  'ListTable3_Accent1',
  'ListTable3_Accent2',
  'ListTable3_Accent3',
  'ListTable3_Accent4',
  'ListTable3_Accent5',
  'ListTable3_Accent6',
  'ListTable4',
  'ListTable4_Accent1',
  'ListTable4_Accent2',
  'ListTable4_Accent3',
  'ListTable4_Accent4',
  'ListTable4_Accent5',
  'ListTable4_Accent6',
  'ListTable5Dark',
  'ListTable5Dark_Accent1',
  'ListTable5Dark_Accent2',
  'ListTable5Dark_Accent3',
  'ListTable5Dark_Accent4',
  'ListTable5Dark_Accent5',
  'ListTable5Dark_Accent6',
  'ListTable6Colorful',
  'ListTable6Colorful_Accent1',
  'ListTable6Colorful_Accent2',
  'ListTable6Colorful_Accent3',
  'ListTable6Colorful_Accent4',
  'ListTable6Colorful_Accent5',
  'ListTable6Colorful_Accent6',
  'ListTable7Colorful',
  'ListTable7Colorful_Accent1',
  'ListTable7Colorful_Accent2',
  'ListTable7Colorful_Accent3',
  'ListTable7Colorful_Accent4',
  'ListTable7Colorful_Accent5',
  'ListTable7Colorful_Accent6',
] as const

export type BuiltInStyle = (typeof BUILT_IN_STYLES)[number]

const BUILT_IN_BY_KEY = new Map<string, BuiltInStyle>(
  BUILT_IN_STYLES.map((s) => [s.toLowerCase().replace(/[^a-z0-9]/g, ''), s]),
)

/** Names users and models actually type, mapped onto the real vocabulary. */
const STYLE_ALIASES: Record<string, BuiltInStyle> = {
  h1: 'Heading1',
  h2: 'Heading2',
  h3: 'Heading3',
  h4: 'Heading4',
  h5: 'Heading5',
  h6: 'Heading6',
  h7: 'Heading7',
  h8: 'Heading8',
  h9: 'Heading9',
  body: 'Normal',
  bodytext: 'Normal',
  default: 'Normal',
  plain: 'Normal',
  paragraph: 'Normal',
  regular: 'Normal',
  heading: 'Heading1',
  subheading: 'Heading2',
  subheader: 'Heading2',
  titlestyle: 'Title',
  documenttitle: 'Title',
  blockquote: 'Quote',
  pullquote: 'IntenseQuote',
  bullet: 'ListParagraph',
  list: 'ListParagraph',
  listitem: 'ListParagraph',
  bold: 'Strong',
  italic: 'Emphasis',
  tableofcontents: 'TocHeading',
  table: 'TableGrid',
  grid: 'TableGrid',
  plaintable: 'PlainTable1',
}

export interface StyleResolution {
  /** Set on `styleBuiltIn` — locale-neutral, so it works on any Word UI language. */
  builtIn?: BuiltInStyle
  /** Set on `style` — a custom style defined in this document's template. */
  custom?: string
}

/**
 * Resolve a style name the model supplied.
 *
 * Anything that is not recognised falls through as a custom style rather than
 * being rejected: documents routinely define their own styles, and the model has
 * no way to enumerate them. Word raises a clear error of its own if the name
 * does not exist, which is a better failure than us refusing a valid style.
 */
export function resolveStyle(raw: string): StyleResolution {
  const key = raw.toLowerCase().replace(/[^a-z0-9]/g, '')
  const direct = BUILT_IN_BY_KEY.get(key)
  if (direct) return { builtIn: direct }
  const alias = STYLE_ALIASES[key]
  if (alias) return { builtIn: alias }
  return { custom: raw }
}

export type WordAlignment = 'Left' | 'Centered' | 'Right' | 'Justified'

const ALIGNMENTS: Record<string, TextAlignment> = {
  left: 'left',
  start: 'left',
  center: 'center',
  centre: 'center',
  centered: 'center',
  middle: 'center',
  right: 'right',
  end: 'right',
  justify: 'justify',
  justified: 'justify',
  both: 'justify',
}

export function resolveAlignment(raw: string): TextAlignment | null {
  return (
    ALIGNMENTS[
      raw
        .trim()
        .toLowerCase()
        .replace(/[^a-z]/g, '')
    ] ?? null
  )
}

const WORD_ALIGNMENTS: Record<TextAlignment, WordAlignment> = {
  left: 'Left',
  center: 'Centered',
  right: 'Right',
  justify: 'Justified',
}

export function wordAlignment(a: TextAlignment): WordAlignment {
  return WORD_ALIGNMENTS[a]
}

export type WordUnderline =
  'None' | 'Single' | 'Double' | 'Thick' | 'Dotted' | 'DashLine' | 'Wave' | 'Word'

const UNDERLINES: Record<string, WordUnderline> = {
  none: 'None',
  off: 'None',
  false: 'None',
  single: 'Single',
  true: 'Single',
  normal: 'Single',
  solid: 'Single',
  double: 'Double',
  thick: 'Thick',
  heavy: 'Thick',
  dotted: 'Dotted',
  dot: 'Dotted',
  dash: 'DashLine',
  dashed: 'DashLine',
  dashline: 'DashLine',
  wave: 'Wave',
  wavy: 'Wave',
  squiggly: 'Wave',
  word: 'Word',
  wordsonly: 'Word',
}

export function resolveUnderline(raw: boolean | string): WordUnderline | null {
  if (typeof raw === 'boolean') return raw ? 'Single' : 'None'
  return (
    UNDERLINES[
      raw
        .trim()
        .toLowerCase()
        .replace(/[^a-z]/g, '')
    ] ?? null
  )
}

export type WordBreak = 'Page' | 'Line' | 'SectionNext' | 'SectionContinuous'

const BREAKS: Record<string, WordBreak> = {
  page: 'Page',
  pagebreak: 'Page',
  line: 'Line',
  linebreak: 'Line',
  newline: 'Line',
  section: 'SectionNext',
  sectionnext: 'SectionNext',
  sectionbreak: 'SectionNext',
  sectioncontinuous: 'SectionContinuous',
  continuous: 'SectionContinuous',
}

export function resolveBreak(raw: string): WordBreak | null {
  return (
    BREAKS[
      raw
        .trim()
        .toLowerCase()
        .replace(/[^a-z]/g, '')
    ] ?? null
  )
}

/** Colour names Word accepts directly. Anything else has to be hex. */
const COLOR_NAMES = new Set([
  'black',
  'white',
  'red',
  'green',
  'blue',
  'yellow',
  'orange',
  'purple',
  'pink',
  'brown',
  'gray',
  'grey',
  'cyan',
  'magenta',
  'lime',
  'navy',
  'teal',
  'olive',
  'maroon',
  'silver',
  'gold',
  'violet',
  'indigo',
  'turquoise',
  'darkblue',
  'darkgreen',
  'darkred',
  'lightblue',
  'lightgreen',
  'lightgray',
  'lightgrey',
])

/**
 * Normalize a colour to something Word's setter accepts.
 *
 * Models write `C00000` as often as `#C00000`, and Word rejects the bare form,
 * so the missing `#` is added rather than failing the edit over punctuation.
 */
export function resolveColor(raw: string): string | null {
  const trimmed = raw.trim()
  if (!trimmed) return null
  const hex = trimmed.replace(/^#/, '')
  if (/^[0-9a-f]{6}$/i.test(hex)) return `#${hex.toUpperCase()}`
  // #RGB shorthand — Word wants all six digits.
  if (/^[0-9a-f]{3}$/i.test(hex)) {
    return `#${hex
      .toUpperCase()
      .split('')
      .map((c) => c + c)
      .join('')}`
  }
  const name = trimmed.toLowerCase().replace(/[^a-z]/g, '')
  return COLOR_NAMES.has(name) ? name : null
}

const FONT_KEYS = [
  'bold',
  'italic',
  'underline',
  'strikeThrough',
  'doubleStrikeThrough',
  'superscript',
  'subscript',
  'smallCaps',
  'allCaps',
  'color',
  'highlightColor',
  'size',
  'font',
] as const

const PARAGRAPH_KEYS = [
  'style',
  'alignment',
  'lineSpacing',
  'spaceBefore',
  'spaceAfter',
  'leftIndent',
  'rightIndent',
  'firstLineIndent',
  'outlineLevel',
] as const

export const FORMATTING_KEYS: string[] = [...FONT_KEYS, ...PARAGRAPH_KEYS]

/** Aliases for the property names, since models guess CSS and Word-UI wording. */
const KEY_ALIASES: Record<string, string> = {
  fontsize: 'size',
  fontname: 'font',
  fontfamily: 'font',
  typeface: 'font',
  fontcolor: 'color',
  textcolor: 'color',
  colour: 'color',
  fontcolour: 'color',
  highlight: 'highlightColor',
  highlightcolour: 'highlightColor',
  background: 'highlightColor',
  backgroundcolor: 'highlightColor',
  strike: 'strikeThrough',
  strikethrough: 'strikeThrough',
  linethrough: 'strikeThrough',
  doublestrikethrough: 'doubleStrikeThrough',
  caps: 'allCaps',
  uppercase: 'allCaps',
  smallcaps: 'smallCaps',
  align: 'alignment',
  textalign: 'alignment',
  justification: 'alignment',
  linespacing: 'lineSpacing',
  linespace: 'lineSpacing',
  leading: 'lineSpacing',
  spacebefore: 'spaceBefore',
  spaceafter: 'spaceAfter',
  spacingbefore: 'spaceBefore',
  spacingafter: 'spaceAfter',
  indent: 'leftIndent',
  leftindent: 'leftIndent',
  rightindent: 'rightIndent',
  firstlineindent: 'firstLineIndent',
  firstline: 'firstLineIndent',
  outlinelevel: 'outlineLevel',
  level: 'outlineLevel',
  paragraphstyle: 'style',
  stylename: 'style',
}

function canonicalKey(raw: string): string | null {
  if ((FORMATTING_KEYS as string[]).includes(raw)) return raw
  const key = raw.toLowerCase().replace(/[^a-z]/g, '')
  const alias = KEY_ALIASES[key]
  if (alias) return alias
  const exact = FORMATTING_KEYS.find((k) => k.toLowerCase() === key)
  return exact ?? null
}

/** Models emit `"true"` and `"12"` as often as `true` and `12`. */
function asBoolean(v: unknown): boolean | null {
  if (typeof v === 'boolean') return v
  if (typeof v === 'string') {
    const s = v.trim().toLowerCase()
    if (s === 'true' || s === 'yes' || s === 'on') return true
    if (s === 'false' || s === 'no' || s === 'off') return false
  }
  return null
}

function asNumber(v: unknown): number | null {
  if (typeof v === 'number') return Number.isFinite(v) ? v : null
  if (typeof v === 'string') {
    // "12pt" / "1.5in" — strip the unit, which is always points here.
    const n = Number.parseFloat(v.trim())
    return Number.isFinite(n) ? n : null
  }
  return null
}

export interface NormalizedFormatting {
  formatting: TextFormatting
  /** Properties this tool does not support at all. */
  unknown: string[]
  /** Known properties whose value could not be used, with the reason. */
  invalid: string[]
}

/**
 * Turn a raw tool-argument bag into formatting Word can apply.
 *
 * Returns what it rejected as well as what it accepted. The caller reports both
 * to the model: a formatter that quietly drops half its instructions and then
 * says "done" is the reason formatting appears to work and then doesn't.
 */
export function normalizeFormatting(raw: Record<string, unknown>): NormalizedFormatting {
  const formatting: TextFormatting = {}
  const unknown: string[] = []
  const invalid: string[] = []

  for (const [rawKey, value] of Object.entries(raw)) {
    if (value === undefined) continue
    const key = canonicalKey(rawKey)
    if (!key) {
      unknown.push(rawKey)
      continue
    }

    switch (key) {
      case 'bold':
      case 'italic':
      case 'strikeThrough':
      case 'doubleStrikeThrough':
      case 'superscript':
      case 'subscript':
      case 'smallCaps':
      case 'allCaps': {
        const b = asBoolean(value)
        if (b === null) invalid.push(`${rawKey} must be true or false`)
        else formatting[key] = b
        break
      }
      case 'underline': {
        const b = asBoolean(value)
        if (b !== null) {
          formatting.underline = b
          break
        }
        const resolved = typeof value === 'string' ? resolveUnderline(value) : null
        if (!resolved) invalid.push(`${rawKey}: "${String(value)}" is not an underline style`)
        // Stored lowercase so it reads as prose in the proposal card;
        // resolveUnderline round-trips its own output.
        else formatting.underline = resolved.toLowerCase()
        break
      }
      case 'color': {
        const c = typeof value === 'string' ? resolveColor(value) : null
        if (!c) invalid.push(`${rawKey}: "${String(value)}" is not a colour (use #RRGGBB)`)
        else formatting.color = c
        break
      }
      case 'highlightColor': {
        // null is meaningful here: it is how a highlight is removed.
        if (value === null || value === 'none' || value === '') {
          formatting.highlightColor = null
          break
        }
        const c = typeof value === 'string' ? resolveColor(value) : null
        if (!c) invalid.push(`${rawKey}: "${String(value)}" is not a colour (use #RRGGBB)`)
        else formatting.highlightColor = c
        break
      }
      case 'size': {
        const n = asNumber(value)
        if (n === null || n <= 0 || n > 1638) invalid.push(`${rawKey} must be a point size`)
        else formatting.size = n
        break
      }
      case 'font': {
        if (typeof value !== 'string' || !value.trim())
          invalid.push(`${rawKey} must be a font name`)
        else formatting.font = value.trim()
        break
      }
      case 'style': {
        if (typeof value !== 'string' || !value.trim())
          invalid.push(`${rawKey} must be a style name`)
        else formatting.style = value.trim()
        break
      }
      case 'alignment': {
        const a = typeof value === 'string' ? resolveAlignment(value) : null
        if (!a) invalid.push(`${rawKey} must be left, center, right, or justify`)
        else formatting.alignment = a
        break
      }
      case 'lineSpacing':
      case 'spaceBefore':
      case 'spaceAfter':
      case 'leftIndent':
      case 'rightIndent':
      case 'firstLineIndent': {
        const n = asNumber(value)
        if (n === null) invalid.push(`${rawKey} must be a number of points`)
        else formatting[key] = n
        break
      }
      case 'outlineLevel': {
        const n = asNumber(value)
        if (n === null || n < 1 || n > 10)
          invalid.push(`${rawKey} must be 1–9, or 10 for body text`)
        else formatting.outlineLevel = Math.round(n)
        break
      }
      default:
        unknown.push(rawKey)
    }
  }

  return { formatting, unknown, invalid }
}

/** True if anything here has to be set on the paragraph rather than the run.
 *  Resolving paragraphs costs an extra round trip, so it is skipped when the
 *  request is character formatting only. */
export function needsParagraphProps(f: TextFormatting): boolean {
  return (PARAGRAPH_KEYS as readonly string[]).some(
    (k) => f[k as keyof TextFormatting] !== undefined,
  )
}

export function hasAnyFormatting(f: TextFormatting): boolean {
  return Object.keys(f).length > 0
}

/** Human-readable rendering, for the proposal card and the tool result. */
export function describeFormatting(f: TextFormatting): string {
  const parts: string[] = []
  if (f.style) parts.push(`style ${f.style}`)
  if (f.bold !== undefined) parts.push(f.bold ? 'bold' : 'not bold')
  if (f.italic !== undefined) parts.push(f.italic ? 'italic' : 'not italic')
  if (f.underline !== undefined) {
    parts.push(
      f.underline === false
        ? 'no underline'
        : `underline ${f.underline === true ? '' : f.underline}`.trim(),
    )
  }
  if (f.strikeThrough !== undefined)
    parts.push(f.strikeThrough ? 'strikethrough' : 'no strikethrough')
  if (f.doubleStrikeThrough) parts.push('double strikethrough')
  if (f.superscript) parts.push('superscript')
  if (f.subscript) parts.push('subscript')
  if (f.smallCaps !== undefined) parts.push(f.smallCaps ? 'small caps' : 'no small caps')
  if (f.allCaps !== undefined) parts.push(f.allCaps ? 'all caps' : 'no all caps')
  if (f.font) parts.push(`font ${f.font}`)
  if (f.size !== undefined) parts.push(`${f.size}pt`)
  if (f.color) parts.push(`colour ${f.color}`)
  if (f.highlightColor !== undefined) {
    parts.push(f.highlightColor === null ? 'no highlight' : `highlight ${f.highlightColor}`)
  }
  if (f.alignment) parts.push(`${f.alignment}-aligned`)
  if (f.lineSpacing !== undefined) parts.push(`line spacing ${f.lineSpacing}pt`)
  if (f.spaceBefore !== undefined) parts.push(`${f.spaceBefore}pt before`)
  if (f.spaceAfter !== undefined) parts.push(`${f.spaceAfter}pt after`)
  if (f.leftIndent !== undefined) parts.push(`left indent ${f.leftIndent}pt`)
  if (f.rightIndent !== undefined) parts.push(`right indent ${f.rightIndent}pt`)
  if (f.firstLineIndent !== undefined) parts.push(`first-line indent ${f.firstLineIndent}pt`)
  if (f.outlineLevel !== undefined) parts.push(`outline level ${f.outlineLevel}`)
  return parts.join(', ')
}

// ─── Page setup ──────────────────────────────────────────────────────────

const TWIPS_PER_INCH = 1440

/** Page sizes in twips, portrait orientation. */
const PAGE_SIZES: Record<string, [number, number]> = {
  letter: [12240, 15840],
  legal: [12240, 20160],
  tabloid: [15840, 24480],
  a3: [16838, 23814],
  a4: [11906, 16838],
  a5: [8391, 11906],
}

function attr(tag: string, name: string): number | null {
  const m = new RegExp(`${name}="(-?\\d+)"`).exec(tag)
  return m ? Number(m[1]) : null
}

function setAttr(tag: string, name: string, value: number): string {
  const existing = new RegExp(`${name}="[^"]*"`)
  if (existing.test(tag)) return tag.replace(existing, `${name}="${value}"`)
  // Not present: splice it in ahead of the closing bracket, self-closing or not.
  return tag.replace(/\/?>$/, (close) => ` ${name}="${value}"${close}`)
}

export type PageSetupResult = { xml: string; changed: string[] } | { error: string }

/**
 * Rewrite page size, orientation, and margins in a document's OOXML.
 *
 * Office.js has no page-setup API at all — margins and orientation live only in
 * the `<w:sectPr>` element — so the only route is to read the body as OOXML,
 * patch it, and write it back. That is the same round trip the revert path
 * already relies on, which is why it is acceptable here.
 */
export function patchPageSetup(ooxml: string, setup: PageSetup): PageSetupResult {
  if (!/<w:sectPr[\s>]/.test(ooxml)) {
    return { error: 'This document does not expose page settings (no section properties found).' }
  }

  const changed: string[] = []
  const size = setup.pageSize ? PAGE_SIZES[setup.pageSize] : undefined
  if (setup.pageSize && !size) {
    return {
      error: `Unknown page size "${setup.pageSize}". Use letter, legal, tabloid, a3, a4, or a5.`,
    }
  }

  const margins = setup.margins ?? {}
  for (const [edge, inches] of Object.entries(margins)) {
    if (typeof inches !== 'number' || !Number.isFinite(inches) || inches < 0 || inches > 22) {
      return { error: `Margin "${edge}" must be a number of inches between 0 and 22.` }
    }
  }

  /**
   * Rewrite one child of `<w:sectPr>`, creating it if the document has none.
   *
   * A document saved by another tool can omit `<w:pgMar>` entirely; a plain
   * regex replace would then change nothing while we cheerfully reported new
   * margins, which is exactly the silent no-op this whole change exists to end.
   */
  function upsert(section: string, tagName: string, rewrite: (tag: string) => string): string {
    const pattern = new RegExp(`<${tagName}[^>]*/?>`)
    const existing = pattern.exec(section)
    if (existing) return section.replace(pattern, rewrite(existing[0]))
    const created = rewrite(`<${tagName}/>`)
    if (section.includes('</w:sectPr>')) {
      return section.replace('</w:sectPr>', `${created}</w:sectPr>`)
    }
    // A self-closing <w:sectPr/> has to become a container to hold the child.
    return section.replace(/<w:sectPr([^>]*)\/>/, `<w:sectPr$1>${created}</w:sectPr>`)
  }

  const xml = ooxml.replace(/<w:sectPr[\s\S]*?<\/w:sectPr>|<w:sectPr[^>]*\/>/g, (section) => {
    let patched = section

    if (size || setup.orientation) {
      patched = upsert(patched, 'w:pgSz', (tag) => {
        const currentW = attr(tag, 'w:w') ?? PAGE_SIZES.letter[0]
        const currentH = attr(tag, 'w:h') ?? PAGE_SIZES.letter[1]
        // Work in portrait terms first, so a size change and an orientation
        // change compose instead of fighting over which dimension is which.
        let [w, h] = size ?? [Math.min(currentW, currentH), Math.max(currentW, currentH)]
        const landscape = setup.orientation
          ? setup.orientation === 'landscape'
          : /w:orient="landscape"/.test(tag)
        if (landscape) [w, h] = [h, w]
        let out = setAttr(setAttr(tag, 'w:w', w), 'w:h', h)
        out = out.replace(/\s*w:orient="[^"]*"/, '')
        if (landscape) out = out.replace(/\s*(\/?)>$/, ' w:orient="landscape"$1>')
        return out
      })
      if (size) changed.push(`page size ${setup.pageSize}`)
      if (setup.orientation) changed.push(setup.orientation)
    }

    const edges: Array<[keyof typeof margins, string]> = [
      ['top', 'w:top'],
      ['bottom', 'w:bottom'],
      ['left', 'w:left'],
      ['right', 'w:right'],
    ]
    const named = edges.filter(([key]) => typeof margins[key] === 'number')
    if (named.length > 0) {
      patched = upsert(patched, 'w:pgMar', (tag) => {
        let out = tag
        for (const [key, name] of named) {
          out = setAttr(out, name, Math.round((margins[key] as number) * TWIPS_PER_INCH))
        }
        return out
      })
      changed.push(`margins (${named.map(([key]) => `${key} ${margins[key]}"`).join(', ')})`)
    }

    return patched
  })

  if (changed.length === 0) {
    return { error: 'Nothing to change — specify orientation, margins, or a page size.' }
  }
  // The same section properties repeat once per section; report each change once.
  return { xml, changed: [...new Set(changed)] }
}
