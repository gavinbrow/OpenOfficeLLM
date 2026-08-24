// Document tool catalog and dispatcher (P4.22).
//
// Tools execute in the PANE, not the host service. Office.js only exists inside
// the task pane's webview, so the host has no way to touch the document — the
// agent loop therefore runs here: the host streams a tool_call back, the pane
// executes it, and the pane re-sends the conversation with the result appended.
//
// Write tools do not apply anything themselves. They hand an Edit to the mode
// dispatcher, which decides whether to apply it now (direct/agentic) or stage it
// for review (propose). Keeping that decision in one place is what makes the
// mode toggle mean something.

import type {
  CellFormula,
  CellValue,
  DetectedHost,
  DocumentAnchor,
  Edit,
  TextAlignment,
  TextTarget,
  ToolDefinition,
} from '@openofficellm/shared'
import type { ToolExecContext, ToolOutcome } from '@openofficellm/ui'
import {
  FORMATTING_KEYS,
  normalizeFormatting,
  resolveAlignment,
  hasAnyFormatting,
} from './wordFormat'

/** The dispatcher's context, defined by the UI package so the same shape
 *  serves every shell. Aliased here because the tool bodies below refer to it
 *  by the shorter name throughout. */
export type ToolContext = ToolExecContext

/**
 * Targeting parameters shared by every tool that acts on existing text.
 *
 * One vocabulary across the tools, rather than each one inventing its own, is
 * what makes this learnable in a single pass: a model that works out how to
 * point `format_text` at paragraph 4 can point `insert_hyperlink` there too.
 * Everything is optional and the fallback is the user's selection, which is
 * both the safest default and the one the user is looking at.
 */
const TARGET_PARAMS: Record<string, unknown> = {
  paragraph: { type: 'number', description: 'Zero-based paragraph number to act on.' },
  paragraphs: {
    type: 'array',
    items: { type: 'number' },
    description: 'Several paragraph numbers.',
  },
  from: { type: 'number', description: 'First paragraph of an inclusive range.' },
  to: { type: 'number', description: 'Last paragraph of an inclusive range.' },
  find: {
    type: 'string',
    description:
      'Act on every occurrence of this text instead of a whole paragraph. The only way to reach a phrase inside a paragraph.',
  },
  matchCase: { type: 'boolean', description: 'With find: match case. Default false.' },
  wholeWord: { type: 'boolean', description: 'With find: whole words only. Default false.' },
  firstOnly: { type: 'boolean', description: 'With find: only the first occurrence.' },
  scope: {
    type: 'string',
    enum: ['selection', 'document'],
    description: 'Act on the selection (default) or the entire document.',
  },
}

/** Character and paragraph properties, flat so they can be set in one call. */
const FORMAT_PARAMS: Record<string, unknown> = {
  bold: { type: 'boolean' },
  italic: { type: 'boolean' },
  underline: {
    type: ['boolean', 'string'],
    description: 'true, or a style: single, double, thick, dotted, dash, wave.',
  },
  strikeThrough: { type: 'boolean' },
  doubleStrikeThrough: { type: 'boolean' },
  superscript: { type: 'boolean' },
  subscript: { type: 'boolean' },
  smallCaps: { type: 'boolean' },
  allCaps: { type: 'boolean' },
  color: { type: 'string', description: 'Text colour as #RRGGBB, e.g. #C00000.' },
  highlightColor: {
    type: ['string', 'null'],
    description: 'Highlight as #RRGGBB, or null to remove the highlight.',
  },
  size: { type: 'number', description: 'Font size in points.' },
  font: { type: 'string', description: 'Font family, e.g. "Calibri".' },
  style: {
    type: 'string',
    description:
      'Word style: Heading 1–9, Title, Subtitle, Quote, Intense Quote, List Paragraph, No Spacing, Caption, Strong, Emphasis, Normal, or a custom style name.',
  },
  alignment: { type: 'string', enum: ['left', 'center', 'right', 'justify'] },
  lineSpacing: { type: 'number', description: 'Points between lines (24 ≈ double for 12pt).' },
  spaceBefore: { type: 'number', description: 'Points of space above the paragraph.' },
  spaceAfter: { type: 'number', description: 'Points of space below the paragraph.' },
  leftIndent: { type: 'number', description: 'Left indent in points (36 = half an inch).' },
  rightIndent: { type: 'number', description: 'Right indent in points.' },
  firstLineIndent: { type: 'number', description: 'First-line indent in points; negative hangs.' },
  outlineLevel: { type: 'number', description: '1–9 for outline levels, 10 for body text.' },
}

const READ_TOOLS_WORD: ToolDefinition[] = [
  {
    name: 'read_selection',
    description:
      'Read the text the user currently has selected in the document. Returns the full document text if nothing is selected.',
    parameters: { type: 'object', properties: {}, required: [] },
  },
  {
    name: 'read_document',
    description:
      'Read the full text of the document, plus its heading outline. Paragraphs are numbered so you can reference them in edits.',
    parameters: { type: 'object', properties: {}, required: [] },
  },
  {
    name: 'search_document',
    description:
      'Find paragraphs containing a phrase. Returns paragraph numbers usable with replace_paragraph.',
    parameters: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Text to search for (case-insensitive).' },
      },
      required: ['query'],
    },
  },
  {
    name: 'read_formatting',
    description:
      "Report the font, colour, style, alignment, spacing, and indents currently applied. Use it to match the document's existing look before you change anything, or to check an edit landed.",
    parameters: { type: 'object', properties: { ...TARGET_PARAMS }, required: [] },
  },
]

const WRITE_TOOLS_WORD: ToolDefinition[] = [
  {
    name: 'replace_selection',
    description:
      "Replace the user's current selection with new text. Use this for rewrites of a selected passage.",
    parameters: {
      type: 'object',
      properties: { text: { type: 'string', description: 'The replacement text.' } },
      required: ['text'],
    },
  },
  {
    name: 'insert_text',
    description: 'Insert text before or after the current selection without replacing it.',
    parameters: {
      type: 'object',
      properties: {
        text: { type: 'string', description: 'The text to insert.' },
        position: {
          type: 'string',
          enum: ['before', 'after'],
          description: 'Where to insert relative to the selection. Defaults to after.',
        },
      },
      required: ['text'],
    },
  },
  {
    name: 'replace_paragraph',
    description:
      'Replace one numbered paragraph. Get the number from read_document or search_document first — never guess it.',
    parameters: {
      type: 'object',
      properties: {
        paragraph: { type: 'number', description: 'Zero-based paragraph index.' },
        text: { type: 'string', description: 'The replacement text.' },
      },
      required: ['paragraph', 'text'],
    },
  },
  {
    name: 'add_comment',
    description: 'Attach a review comment to a paragraph without changing its text.',
    parameters: {
      type: 'object',
      properties: {
        paragraph: { type: 'number', description: 'Zero-based paragraph index.' },
        text: { type: 'string', description: 'The comment body.' },
      },
      required: ['paragraph', 'text'],
    },
  },
  {
    name: 'format_text',
    description:
      'Change how text looks: font, size, colour, highlight, bold/italic/underline, style, alignment, spacing, indents. Target the selection (default), paragraph numbers, a paragraph range, the whole document, or every occurrence of a phrase. Only the properties you name change; the rest are left alone.',
    parameters: {
      type: 'object',
      properties: { ...TARGET_PARAMS, ...FORMAT_PARAMS },
      required: [],
    },
  },
  {
    name: 'insert_paragraph',
    description:
      'Add a new paragraph. Use it for headings, new body text, or blank spacing. Paragraph numbers after the insertion point all shift by one, so re-read before further numbered edits.',
    parameters: {
      type: 'object',
      properties: {
        text: { type: 'string', description: 'The paragraph text. Use "" for a blank line.' },
        after: {
          type: ['number', 'string'],
          description:
            'Paragraph number to insert after, or "start", "end", or "selection". Defaults to the end.',
        },
        style: { type: 'string', description: 'Style for the new paragraph, e.g. "Heading 2".' },
      },
      required: ['text'],
    },
  },
  {
    name: 'delete_paragraph',
    description:
      'Delete one paragraph, several, or an inclusive range. Numbers after a deletion shift, so re-read before further numbered edits.',
    parameters: {
      type: 'object',
      properties: {
        paragraph: { type: 'number', description: 'A single paragraph number.' },
        paragraphs: { type: 'array', items: { type: 'number' }, description: 'Several numbers.' },
        from: { type: 'number', description: 'First paragraph of a range.' },
        to: { type: 'number', description: 'Last paragraph of a range.' },
      },
      required: [],
    },
  },
  {
    name: 'set_list',
    description:
      'Turn paragraphs into a bulleted or numbered list, or strip list formatting off them. The paragraphs need not be adjacent.',
    parameters: {
      type: 'object',
      properties: {
        paragraph: { type: 'number', description: 'A single paragraph number.' },
        paragraphs: {
          type: 'array',
          items: { type: 'number' },
          description: 'Paragraph numbers, e.g. [2, 4].',
        },
        from: { type: 'number', description: 'First paragraph of a contiguous range.' },
        to: { type: 'number', description: 'Last paragraph of a contiguous range.' },
        type: {
          type: 'string',
          enum: ['bullet', 'number', 'none'],
          description: '"none" removes list formatting.',
        },
        level: { type: 'number', description: 'Nesting level, 1 = top. Defaults to 1.' },
      },
      required: ['type'],
    },
  },
  {
    name: 'insert_table',
    description:
      'Insert a table. Pass rows as an array of arrays of cell text; the first row is treated as the header unless headerRow is false.',
    parameters: {
      type: 'object',
      properties: {
        rows: {
          type: 'array',
          description: 'Rows of cell text, e.g. [["Region","Revenue"],["North","1200"]].',
          items: { type: 'array', items: { type: 'string' } },
        },
        headerRow: {
          type: 'boolean',
          description: 'Treat the first row as a header. Default true.',
        },
        style: {
          type: 'string',
          description:
            'Table style, e.g. "Table Grid", "Grid Table 4 Accent 1", "List Table 3". Defaults to Grid Table 4 Accent 1.',
        },
        after: {
          type: ['number', 'string'],
          description: 'Paragraph number to insert after, or "start", "end", "selection".',
        },
      },
      required: ['rows'],
    },
  },
  {
    name: 'insert_break',
    description: 'Insert a page break, line break, or section break.',
    parameters: {
      type: 'object',
      properties: {
        type: { type: 'string', enum: ['page', 'line', 'section', 'sectionContinuous'] },
        after: {
          type: ['number', 'string'],
          description: 'Paragraph number to insert after, or "start", "end", "selection".',
        },
      },
      required: ['type'],
    },
  },
  {
    name: 'insert_hyperlink',
    description:
      'Turn existing text into a link. Target the selection, a paragraph, or every occurrence of a phrase.',
    parameters: {
      type: 'object',
      properties: {
        url: { type: 'string', description: 'http(s):// or mailto: address.' },
        text: { type: 'string', description: 'Optional replacement text for the link.' },
        ...TARGET_PARAMS,
      },
      required: ['url'],
    },
  },
  {
    name: 'replace_all',
    description:
      'Find and replace text throughout the document. Reports how many occurrences changed.',
    parameters: {
      type: 'object',
      properties: {
        find: { type: 'string', description: 'Text to find.' },
        replace: { type: 'string', description: 'Replacement text. Use "" to delete.' },
        matchCase: { type: 'boolean' },
        wholeWord: { type: 'boolean' },
      },
      required: ['find', 'replace'],
    },
  },
  {
    name: 'set_header_footer',
    description:
      'Set the page header or footer text, optionally with an automatic page number. Replaces whatever is there now.',
    parameters: {
      type: 'object',
      properties: {
        part: { type: 'string', enum: ['header', 'footer'] },
        text: { type: 'string', description: 'The text. Use "" for a page number alone.' },
        alignment: { type: 'string', enum: ['left', 'center', 'right', 'justify'] },
        pageNumber: { type: 'boolean', description: 'Append an automatic page number field.' },
      },
      required: ['part', 'text'],
    },
  },
  {
    name: 'set_page_setup',
    description:
      'Set page orientation, paper size, and margins for the document. Margins are in inches.',
    parameters: {
      type: 'object',
      properties: {
        orientation: { type: 'string', enum: ['portrait', 'landscape'] },
        pageSize: { type: 'string', enum: ['letter', 'legal', 'tabloid', 'a3', 'a4', 'a5'] },
        margins: { type: 'number', description: 'Same margin on all four edges, in inches.' },
        marginTop: { type: 'number' },
        marginBottom: { type: 'number' },
        marginLeft: { type: 'number' },
        marginRight: { type: 'number' },
      },
      required: [],
    },
  },
]

const READ_TOOLS_EXCEL: ToolDefinition[] = [
  {
    name: 'read_range',
    description:
      "Read the selected range, or the sheet's used range if the selection is a single cell. Large ranges are sampled — the true row count is reported.",
    parameters: { type: 'object', properties: {}, required: [] },
  },
  {
    name: 'read_sheet',
    description: 'Read the active sheet: structure, column types, and sampled rows.',
    parameters: { type: 'object', properties: {}, required: [] },
  },
  {
    name: 'search_document',
    description: 'Find cells on the active sheet containing a value. Returns cell addresses.',
    parameters: {
      type: 'object',
      properties: { query: { type: 'string', description: 'Text to search for.' } },
      required: ['query'],
    },
  },
]

const WRITE_TOOLS_EXCEL: ToolDefinition[] = [
  {
    name: 'write_range',
    description: 'Write literal values into specific cells.',
    parameters: {
      type: 'object',
      properties: {
        sheet: { type: 'string', description: 'Sheet name. Omit for the active sheet.' },
        cells: {
          type: 'array',
          description: 'Cells to write.',
          items: {
            type: 'object',
            properties: {
              cell: { type: 'string', description: 'Address such as B7.' },
              value: { type: ['string', 'number', 'boolean', 'null'] },
            },
            required: ['cell', 'value'],
          },
        },
      },
      required: ['cells'],
    },
  },
  {
    name: 'write_formula',
    description:
      'Write formulas into cells. Refused if the target already holds a formula, so read first when overwriting is intended.',
    parameters: {
      type: 'object',
      properties: {
        sheet: { type: 'string', description: 'Sheet name. Omit for the active sheet.' },
        cells: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              cell: { type: 'string', description: 'Address such as D2.' },
              formula: { type: 'string', description: 'Formula including the leading =.' },
            },
            required: ['cell', 'formula'],
          },
        },
      },
      required: ['cells'],
    },
  },
  {
    name: 'format_range',
    description: 'Apply font, fill, or number formatting to a cell range.',
    parameters: {
      type: 'object',
      properties: {
        range: { type: 'string', description: 'Address such as A1:D10.' },
        bold: { type: 'boolean' },
        italic: { type: 'boolean' },
        color: { type: 'string', description: 'Font colour as hex.' },
        fill: { type: 'string', description: 'Background colour as hex.' },
        numberFormat: { type: 'string', description: 'Number format, e.g. "0.00%".' },
      },
      required: ['range'],
    },
  },
]

/**
 * The tool catalog for a host.
 *
 * Write tools are withheld entirely when the context scope is 'none' — the user
 * has said the model may not see the document, and a model that can rewrite what
 * it cannot read is strictly worse than one that can do neither.
 */
export function toolCatalog(host: DetectedHost, allowWrites: boolean): ToolDefinition[] {
  // 'browser' cannot reach this shell — the Office pane only ever detects Word
  // or Excel — but the Shell contract is shared, so the signature admits it and
  // the answer is the same as for 'none': no tools rather than Word's tools
  // pointed at something that is not a Word document.
  if (host !== 'word' && host !== 'excel') return []
  const read = host === 'word' ? READ_TOOLS_WORD : READ_TOOLS_EXCEL
  const write = host === 'word' ? WRITE_TOOLS_WORD : WRITE_TOOLS_EXCEL
  return allowWrites ? [...read, ...write] : [...read]
}

const WRITE_TOOL_NAMES = new Set([
  ...WRITE_TOOLS_WORD.map((t) => t.name),
  ...WRITE_TOOLS_EXCEL.map((t) => t.name),
  // No longer advertised, but still dispatched — see `format_paragraph` below.
  'format_paragraph',
])

export function isWriteTool(name: string): boolean {
  return WRITE_TOOL_NAMES.has(name)
}

/** Argument names that say *where* to act, so they are never mistaken for
 *  formatting properties when the rest of the bag is handed to the formatter. */
const TARGET_KEYS = new Set([
  'paragraph',
  'paragraphs',
  'from',
  'to',
  'find',
  'query',
  'search',
  'matchCase',
  'wholeWord',
  'firstOnly',
  'scope',
  'target',
])

/** Models fill unused parameters with explicit nulls rather than omitting them,
 *  and `null` coerces to 0 — so "no paragraph given" would silently become
 *  "paragraph 0" unless absence is checked for properly. */
function given(v: unknown): boolean {
  return v !== undefined && v !== null && v !== ''
}

function asIndexList(v: unknown): number[] | null {
  const list = (Array.isArray(v) ? v : [v]).filter(given)
  const out: number[] = []
  for (const item of list) {
    const n = typeof item === 'number' ? item : Number(item)
    if (!Number.isInteger(n) || n < 0) return null
    out.push(n)
  }
  return out.length > 0 ? out : null
}

/**
 * Work out what the model wants to act on.
 *
 * Order matters: an explicit phrase beats paragraph numbers, numbers beat the
 * scope keyword, and the selection is the fallback. Models routinely send both
 * `find` and a paragraph number when they mean "this phrase, which is in that
 * paragraph" — honouring the phrase is what they meant.
 */
function parseTarget(args: Record<string, unknown>): TextTarget | { error: string } {
  const find = args.find ?? args.query ?? args.search
  if (typeof find === 'string' && find.trim()) {
    return {
      kind: 'search',
      search: find,
      matchCase: args.matchCase === true,
      wholeWord: args.wholeWord === true,
      firstOnly: args.firstOnly === true,
    }
  }

  if (given(args.paragraphs) || given(args.paragraph)) {
    const indices = asIndexList(given(args.paragraphs) ? args.paragraphs : args.paragraph)
    if (!indices) return { error: 'Paragraph numbers must be whole numbers, zero or greater.' }
    return { kind: 'paragraphs', paragraphs: indices }
  }

  if (given(args.from) || given(args.to)) {
    const from = Number(given(args.from) ? args.from : args.to)
    const to = Number(given(args.to) ? args.to : args.from)
    if (!Number.isInteger(from) || !Number.isInteger(to) || from < 0 || to < 0) {
      return { error: '"from" and "to" must be whole paragraph numbers.' }
    }
    const lo = Math.min(from, to)
    const hi = Math.max(from, to)
    return {
      kind: 'paragraphs',
      paragraphs: Array.from({ length: hi - lo + 1 }, (_, i) => lo + i),
    }
  }

  // `target` is as natural a name as `scope`; accepting only one of them would
  // silently fall back to the selection when the model picked the other.
  const named = args.scope ?? args.target
  const scope = typeof named === 'string' ? named.trim().toLowerCase() : ''
  if (scope === 'document' || scope === 'all' || scope === 'body') return { kind: 'document' }
  if (scope && scope !== 'selection') {
    return { error: `Unknown scope "${String(args.scope)}". Use "selection" or "document".` }
  }
  return { kind: 'selection' }
}

/** Where an insertion goes. Defaults to the end of the document, which is the
 *  only choice that cannot silently overwrite or interleave with existing text. */
function parseAnchor(v: unknown): DocumentAnchor | { error: string } {
  if (v === undefined || v === null || v === '') return 'end'
  if (typeof v === 'number' || (typeof v === 'string' && /^\d+$/.test(v.trim()))) {
    const n = Number(v)
    if (!Number.isInteger(n) || n < 0) return { error: 'Paragraph numbers start at 0.' }
    return n
  }
  const word = String(v).trim().toLowerCase()
  if (word === 'start' || word === 'beginning' || word === 'top') return 'start'
  if (word === 'end' || word === 'bottom' || word === 'append') return 'end'
  if (word === 'selection' || word === 'cursor' || word === 'here') return 'selection'
  return {
    error: `Unknown position "${String(v)}". Use a paragraph number, "start", "end", or "selection".`,
  }
}

/**
 * Coerce whatever shape a model produced into a grid of cell text.
 *
 * Asked for a table of data, models emit an array of objects about as often as
 * an array of arrays. Rejecting that costs a round trip to say something the
 * caller could have worked out, so objects are turned into a header row plus
 * value rows using the keys of the first object.
 */
const PAGE_SIZES = ['letter', 'legal', 'tabloid', 'a3', 'a4', 'a5'] as const

/** Paragraph numbers for tools that act on paragraphs as whole units. Reuses
 *  the shared targeting vocabulary so `paragraphs: [2, 4]` and `from`/`to` both
 *  work, and refuses the phrase form, which cannot address a whole paragraph. */
function paragraphNumbers(
  args: Record<string, unknown>,
  tool: string,
): { list: number[] } | { error: string } {
  const target = parseTarget(args)
  if ('error' in target) return target
  if (target.kind !== 'paragraphs' || !target.paragraphs?.length) {
    return {
      error: `${tool} needs paragraph numbers: "paragraph", "paragraphs", or "from" and "to".`,
    }
  }
  return { list: target.paragraphs }
}

/** Phrasing for the proposal card and the tool result, so the user reads
 *  "Format 3 paragraphs" rather than a serialized target object. A bare noun
 *  phrase, so callers can put their own verb in front of it. */
function describeTarget(target: TextTarget): string {
  switch (target.kind) {
    case 'document':
      return 'the whole document'
    case 'search':
      return `${target.firstOnly ? 'the first' : 'every'} "${target.search}"`
    case 'paragraphs': {
      const list = target.paragraphs ?? []
      if (list.length === 1) return `paragraph ${list[0]}`
      if (list.length <= 4) return `paragraphs ${list.join(', ')}`
      return `${list.length} paragraphs`
    }
    default:
      return 'the selection'
  }
}

function describeAnchor(at: DocumentAnchor): string {
  if (at === 'start') return 'at the start'
  if (at === 'end') return 'at the end'
  if (at === 'selection') return 'at the cursor'
  return `after paragraph ${at}`
}

function normalizeTableRows(raw: unknown): string[][] | null {
  if (!Array.isArray(raw) || raw.length === 0) return null

  const first = raw[0]
  if (first && typeof first === 'object' && !Array.isArray(first)) {
    const columns = Object.keys(first as Record<string, unknown>)
    if (columns.length === 0) return null
    const body = raw.map((row) => {
      const record = (row ?? {}) as Record<string, unknown>
      return columns.map((c) => String(record[c] ?? ''))
    })
    return [columns, ...body]
  }

  const rows = raw.map((row) =>
    Array.isArray(row) ? row.map((cell) => String(cell ?? '')) : [String(row ?? '')],
  )
  return rows.length > 0 ? rows : null
}

function parseArgs(raw: string): Record<string, unknown> {
  try {
    const parsed = JSON.parse(raw || '{}')
    return parsed && typeof parsed === 'object' ? (parsed as Record<string, unknown>) : {}
  } catch {
    return {}
  }
}

function requireString(args: Record<string, unknown>, key: string): string | null {
  const v = args[key]
  return typeof v === 'string' && v.length > 0 ? v : null
}

function requireIndex(args: Record<string, unknown>, key: string): number | null {
  const v = args[key]
  const n = typeof v === 'number' ? v : Number(v)
  return Number.isInteger(n) && n >= 0 ? n : null
}

/**
 * Execute one document tool call.
 *
 * Errors come back as tool results, not exceptions: the model can read "no
 * paragraph 40" and correct itself, whereas a thrown error would kill the turn.
 */
export async function executeDocumentTool(
  name: string,
  argsJson: string,
  ctx: ToolContext,
): Promise<ToolOutcome> {
  const args = parseArgs(argsJson)
  const { adapter, mode } = ctx

  try {
    switch (name) {
      // ── Word reads ──
      case 'read_selection': {
        const c = await adapter.getContext('selection')
        return {
          content: c.text
            ? `(${c.scope})\n${c.text}`
            : 'The document is empty and nothing is selected.',
          isError: false,
        }
      }
      case 'read_document': {
        const c = await adapter.getContext('document')
        const outline = c.outline ? `Outline:\n${c.outline}\n\n` : ''
        return {
          content: c.text ? `${outline}${c.text}` : 'The document is empty.',
          isError: false,
        }
      }

      // ── Excel reads ──
      case 'read_range': {
        const c = await adapter.getContext('range')
        return { content: c.text || 'The selected range is empty.', isError: false }
      }
      case 'read_sheet': {
        const c = await adapter.getContext('sheet')
        const outline = c.outline ? `${c.outline}\n\n` : ''
        return { content: `${outline}${c.text || 'The sheet is empty.'}`, isError: false }
      }

      // ── Shared ──
      case 'search_document': {
        const query = requireString(args, 'query')
        if (!query) return { content: 'search_document requires a "query".', isError: true }
        const hits = await adapter.search(query)
        if (hits.length === 0) return { content: `No matches for "${query}".`, isError: false }
        return {
          content: hits.map((h) => `[${h.location}] ${h.text}`).join('\n'),
          isError: false,
        }
      }

      // ── Word writes ──
      case 'replace_selection': {
        const text = requireString(args, 'text')
        if (text === null) return { content: 'replace_selection requires "text".', isError: true }
        return stageOrApply(ctx, { kind: 'replaceSelection', text }, 'Replace the selection')
      }
      case 'insert_text': {
        const text = requireString(args, 'text')
        if (text === null) return { content: 'insert_text requires "text".', isError: true }
        const before = String(args.position ?? 'after') === 'before'
        return stageOrApply(
          ctx,
          before ? { kind: 'insertBefore', text } : { kind: 'insertAfter', text },
          `Insert text ${before ? 'before' : 'after'} the selection`,
        )
      }
      case 'replace_paragraph': {
        const idx = requireIndex(args, 'paragraph')
        const text = requireString(args, 'text')
        if (idx === null)
          return { content: 'replace_paragraph needs a "paragraph" index.', isError: true }
        if (text === null) return { content: 'replace_paragraph requires "text".', isError: true }
        return stageOrApply(
          ctx,
          { kind: 'replaceRange', rangeId: String(idx), text },
          `Replace paragraph ${idx}`,
        )
      }
      case 'add_comment': {
        const idx = requireIndex(args, 'paragraph')
        const text = requireString(args, 'text')
        if (idx === null || text === null) {
          return { content: 'add_comment needs "paragraph" and "text".', isError: true }
        }
        return stageOrApply(
          ctx,
          { kind: 'addComment', rangeId: String(idx), text },
          `Comment on paragraph ${idx}`,
        )
      }
      // `format_paragraph` was the whole of Word formatting before format_text.
      // Still dispatched, unadvertised: transcripts persisted by an earlier
      // build replay through here, and models that learned the old name keep
      // working instead of getting "unknown tool".
      case 'format_paragraph':
      case 'format_text': {
        const target = parseTarget(args)
        if ('error' in target) return { content: target.error, isError: true }

        const raw: Record<string, unknown> = {}
        for (const [k, v] of Object.entries(args)) if (!TARGET_KEYS.has(k)) raw[k] = v
        const { formatting, unknown, invalid } = normalizeFormatting(raw)

        const problems = [...invalid]
        if (unknown.length > 0) {
          problems.push(
            `${unknown.join(', ')} — not a formatting property. Supported: ${FORMATTING_KEYS.join(', ')}`,
          )
        }
        if (!hasAnyFormatting(formatting)) {
          return {
            content:
              problems.length > 0
                ? `Nothing was changed. ${problems.join('; ')}.`
                : `format_text needs at least one property to change (${FORMATTING_KEYS.join(', ')}).`,
            isError: true,
          }
        }

        const outcome = await Promise.resolve(
          stageOrApply(
            ctx,
            { kind: 'formatText', target, formatting },
            `Format ${describeTarget(target)}`,
          ),
        )
        // Applying what worked and naming what did not beats refusing the whole
        // call — but reporting only the success would teach the model that an
        // unsupported property is fine.
        return problems.length > 0
          ? {
              content: `${outcome.content}\nNot applied: ${problems.join('; ')}.`,
              isError: outcome.isError,
            }
          : outcome
      }

      case 'insert_paragraph': {
        const text = args.text
        if (typeof text !== 'string') {
          return {
            content: 'insert_paragraph requires "text" (use "" for a blank line).',
            isError: true,
          }
        }
        const at = parseAnchor(args.after ?? args.at ?? args.position)
        if (typeof at === 'object') return { content: at.error, isError: true }
        const style = typeof args.style === 'string' ? args.style : undefined
        return stageOrApply(
          ctx,
          { kind: 'insertParagraph', at, text, style },
          `Insert a paragraph ${describeAnchor(at)}`,
        )
      }

      case 'delete_paragraph': {
        const paragraphs = paragraphNumbers(args, 'delete_paragraph')
        if ('error' in paragraphs) return { content: paragraphs.error, isError: true }
        return stageOrApply(
          ctx,
          { kind: 'deleteParagraphs', paragraphs: paragraphs.list },
          `Delete ${describeTarget({ kind: 'paragraphs', paragraphs: paragraphs.list })}`,
        )
      }

      case 'set_list': {
        const paragraphs = paragraphNumbers(args, 'set_list')
        if ('error' in paragraphs) return { content: paragraphs.error, isError: true }
        const raw = String(args.type ?? args.listType ?? '').toLowerCase()
        const listType =
          raw === 'bullet' || raw === 'bulleted' || raw === 'unordered'
            ? ('bullet' as const)
            : raw === 'number' || raw === 'numbered' || raw === 'ordered'
              ? ('number' as const)
              : raw === 'none' || raw === 'remove'
                ? ('none' as const)
                : null
        if (!listType) {
          return { content: 'set_list "type" must be bullet, number, or none.', isError: true }
        }
        const level = Number(args.level)
        const where = describeTarget({ kind: 'paragraphs', paragraphs: paragraphs.list })
        return stageOrApply(
          ctx,
          {
            kind: 'setList',
            paragraphs: paragraphs.list,
            listType,
            level: Number.isFinite(level) ? level : undefined,
          },
          listType === 'none'
            ? `Remove list formatting from ${where}`
            : `Make ${where} a ${listType === 'bullet' ? 'bulleted' : 'numbered'} list`,
        )
      }

      case 'insert_table': {
        const rows = normalizeTableRows(args.rows ?? args.data ?? args.values)
        if (!rows) {
          return {
            content: 'insert_table needs "rows" as an array of arrays of cell text.',
            isError: true,
          }
        }
        const at = parseAnchor(args.after ?? args.at ?? args.position)
        if (typeof at === 'object') return { content: at.error, isError: true }
        const columns = Math.max(...rows.map((r) => r.length))
        return stageOrApply(
          ctx,
          {
            kind: 'insertTable',
            at,
            rows,
            headerRow: args.headerRow !== false,
            style: typeof args.style === 'string' ? args.style : undefined,
          },
          `Insert a ${rows.length}×${columns} table ${describeAnchor(at)}`,
        )
      }

      case 'insert_break': {
        const type = String(args.type ?? args.breakType ?? 'page')
        const at = parseAnchor(args.after ?? args.at ?? args.position)
        if (typeof at === 'object') return { content: at.error, isError: true }
        const breakType =
          type === 'sectionContinuous'
            ? ('sectionContinuous' as const)
            : type === 'section'
              ? ('section' as const)
              : type === 'line'
                ? ('line' as const)
                : ('page' as const)
        return stageOrApply(
          ctx,
          { kind: 'insertBreak', at, breakType },
          `Insert a ${breakType} break ${describeAnchor(at)}`,
        )
      }

      case 'insert_hyperlink': {
        const url = requireString(args, 'url')
        if (!url) return { content: 'insert_hyperlink requires a "url".', isError: true }
        // `text` is the link label here, not a target, so it must not be read
        // as one — parseTarget ignores it by design.
        const target = parseTarget(args)
        if ('error' in target) return { content: target.error, isError: true }
        const text = typeof args.text === 'string' && args.text ? args.text : undefined
        return stageOrApply(
          ctx,
          { kind: 'insertHyperlink', target, url, text },
          `Link ${describeTarget(target)} to ${url}`,
        )
      }

      case 'replace_all': {
        const find = requireString(args, 'find')
        const replace = args.replace
        if (!find) return { content: 'replace_all requires "find".', isError: true }
        if (typeof replace !== 'string') {
          return { content: 'replace_all requires "replace" (use "" to delete).', isError: true }
        }
        return stageOrApply(
          ctx,
          {
            kind: 'replaceAll',
            find,
            replace,
            matchCase: args.matchCase === true,
            wholeWord: args.wholeWord === true,
          },
          `Replace every "${find}" with "${replace}"`,
        )
      }

      case 'set_header_footer': {
        const part = String(args.part ?? '').toLowerCase() === 'header' ? 'header' : 'footer'
        const text = args.text
        if (typeof text !== 'string') {
          return {
            content: 'set_header_footer requires "text" (use "" for a page number alone).',
            isError: true,
          }
        }
        const alignment =
          typeof args.alignment === 'string' ? resolveAlignment(args.alignment) : null
        return stageOrApply(
          ctx,
          {
            kind: 'setHeaderFooter',
            part,
            text,
            alignment: (alignment ?? undefined) as TextAlignment | undefined,
            pageNumber: args.pageNumber === true,
          },
          `Set the ${part}`,
        )
      }

      case 'set_page_setup': {
        const all = given(args.margins) ? Number(args.margins) : Number.NaN
        const edge = (key: string): number | undefined => {
          if (given(args[key])) {
            const v = Number(args[key])
            if (Number.isFinite(v)) return v
          }
          return Number.isFinite(all) ? all : undefined
        }
        const margins = {
          top: edge('marginTop'),
          bottom: edge('marginBottom'),
          left: edge('marginLeft'),
          right: edge('marginRight'),
        }
        const named = Object.entries(margins).filter(([, v]) => v !== undefined)
        const orientation =
          args.orientation === 'landscape' || args.orientation === 'portrait'
            ? args.orientation
            : undefined
        const requested =
          typeof args.pageSize === 'string' ? args.pageSize.toLowerCase().trim() : ''
        if (requested && !PAGE_SIZES.includes(requested as (typeof PAGE_SIZES)[number])) {
          return {
            content: `Unknown pageSize "${requested}". Use ${PAGE_SIZES.join(', ')}.`,
            isError: true,
          }
        }
        const pageSize = requested ? (requested as (typeof PAGE_SIZES)[number]) : undefined
        if (!orientation && !pageSize && named.length === 0) {
          return {
            content: 'set_page_setup needs an orientation, a pageSize, or at least one margin.',
            isError: true,
          }
        }
        return stageOrApply(
          ctx,
          {
            kind: 'setPageSetup',
            setup: {
              orientation,
              pageSize,
              ...(named.length > 0 ? { margins: Object.fromEntries(named) } : {}),
            },
          },
          'Change the page setup',
        )
      }

      case 'read_formatting': {
        if (!adapter.readFormatting) {
          return { content: 'Formatting cannot be read in this application.', isError: true }
        }
        const target = parseTarget(args)
        if ('error' in target) return { content: target.error, isError: true }
        return { content: await adapter.readFormatting(target), isError: false }
      }

      // ── Excel writes ──
      case 'write_range': {
        const cells = normalizeValueCells(args.cells)
        if (!cells)
          return { content: 'write_range needs a "cells" array of {cell, value}.', isError: true }
        return stageOrApply(
          ctx,
          { kind: 'setCellValues', sheet: String(args.sheet ?? ''), cells },
          `Write ${cells.length} value(s)`,
        )
      }
      case 'write_formula': {
        const cells = normalizeFormulaCells(args.cells)
        if (!cells) {
          return {
            content: 'write_formula needs a "cells" array of {cell, formula}.',
            isError: true,
          }
        }
        return stageOrApply(
          ctx,
          { kind: 'setCellFormulas', sheet: String(args.sheet ?? ''), cells },
          `Write ${cells.length} formula(s)`,
        )
      }
      case 'format_range': {
        const range = requireString(args, 'range')
        if (!range) return { content: 'format_range requires a "range".', isError: true }
        const { range: _r, ...formatting } = args
        return stageOrApply(
          ctx,
          { kind: 'applyFormatting', rangeId: range, formatting },
          `Format ${range}`,
        )
      }

      default:
        return { content: `Unknown tool: ${name}`, isError: true }
    }
  } catch (e) {
    const message = (e as Error).message ?? String(e)
    return { content: `Tool "${name}" failed: ${message}`, isError: true }
  }

  function stageOrApply(
    toolCtx: ToolContext,
    edit: Edit,
    description: string,
  ): Promise<ToolOutcome> | ToolOutcome {
    if (mode === 'propose') {
      toolCtx.propose(edit, description)
      return {
        content: `Proposed: ${description}. It is staged for the user to review and has NOT been applied. Tell the user what you are proposing and why — do not say the document has changed.`,
        isError: false,
      }
    }
    // `applyEdits` is optional on the adapter contract because read-only hosts
    // exist. Word and Excel both implement it, so this is unreachable here —
    // but a thrown TypeError on `undefined` would surface as "Tool failed" with
    // no explanation, and a shell that ever does ship a read-only adapter
    // deserves the real reason.
    const apply = toolCtx.adapter.applyEdits
    if (!apply) {
      return { content: `This host cannot be edited, so "${name}" did nothing.`, isError: true }
    }
    return apply.call(toolCtx.adapter, [edit]).then((r) => ({
      content: r.ok ? r.summary : `Edit rejected: ${r.summary}`,
      isError: !r.ok,
    }))
  }
}

function normalizeValueCells(raw: unknown): CellValue[] | null {
  if (!Array.isArray(raw)) return null
  const out: CellValue[] = []
  for (const item of raw) {
    if (!item || typeof item !== 'object') return null
    const o = item as Record<string, unknown>
    if (typeof o.cell !== 'string') return null
    const v = o.value
    out.push({
      cell: o.cell,
      value:
        typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean' || v === null
          ? v
          : String(v ?? ''),
    })
  }
  return out.length > 0 ? out : null
}

function normalizeFormulaCells(raw: unknown): CellFormula[] | null {
  if (!Array.isArray(raw)) return null
  const out: CellFormula[] = []
  for (const item of raw) {
    if (!item || typeof item !== 'object') return null
    const o = item as Record<string, unknown>
    if (typeof o.cell !== 'string' || typeof o.formula !== 'string') return null
    // A model that omits the leading = writes a text cell that looks right and
    // computes nothing.
    const formula = o.formula.startsWith('=') ? o.formula : `=${o.formula}`
    out.push({ cell: o.cell, formula })
  }
  return out.length > 0 ? out : null
}
