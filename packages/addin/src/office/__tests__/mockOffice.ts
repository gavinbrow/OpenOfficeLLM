// Minimal Office.js fakes for adapter tests (P4.24).
//
// Models the proxy-object pattern the real API uses: you call `load()` to
// declare what you want, `sync()` resolves it, and reading a property before a
// sync gives you nothing. Getting that shape right is the point — a mock that
// returns values immediately would pass tests that the real Word rejects.
//
// Property writes are *recorded* rather than swallowed. An earlier version of
// this mock accepted `font.bold = true` into a black hole, which meant the
// formatting path could be broken in every direction and every test still
// passed — precisely the failure this file exists to catch.

export interface MockParagraph {
  text: string
  style: string
  /** Set via the locale-neutral `styleBuiltIn` setter. */
  styleBuiltIn?: string
  /** Records what insertText was called with, for assertions. */
  inserted?: { text: string; location: string }
  comments: string[]
  font: Record<string, unknown>
  /** Paragraph-level properties: alignment, spacing, indents, outline level. */
  props: Record<string, unknown>
  deleted?: boolean
  /** Set when the paragraph is attached to a list. */
  list?: { id: number; level: number }
  detachedFromList?: boolean
}

export function makeParagraph(text: string, style = 'Normal'): MockParagraph {
  return { text, style, comments: [], font: {}, props: {} }
}

/** A recording stand-in for Word.Font — every setter lands in `sink`. */
function fontProxy(sink: Record<string, unknown>) {
  return new Proxy({} as Record<string, unknown>, {
    set: (_t, key, value) => {
      sink[String(key)] = value
      return true
    },
    get: (_t, key) => (key === 'load' ? () => undefined : sink[String(key)]),
  })
}

interface RangeRecord {
  text: string
  font: Record<string, unknown>
  hyperlink?: string
  styleBuiltIn?: string
  fields: string[]
}

class ParagraphProxy {
  constructor(
    readonly p: MockParagraph,
    private pending: Array<() => void>,
    private state: WordDocState,
  ) {}

  get text() {
    return this.p.text
  }
  get style() {
    return this.p.style
  }
  set style(v: string) {
    this.p.style = v
  }
  set styleBuiltIn(v: string) {
    this.p.styleBuiltIn = v
  }
  get styleBuiltIn() {
    return this.p.styleBuiltIn ?? ''
  }
  get isListItem() {
    return this.p.list !== undefined
  }

  // A getter, not a field: class fields initialize before the constructor's
  // parameter properties exist under useDefineForClassFields.
  get font() {
    return fontProxy(this.p.font)
  }

  load() {
    return undefined
  }

  // Paragraph-level properties all land in `props`, so a test can assert on
  // exactly what the adapter set without the mock enumerating each one.
  set alignment(v: string) {
    this.p.props.alignment = v
  }
  set lineSpacing(v: number) {
    this.p.props.lineSpacing = v
  }
  set spaceBefore(v: number) {
    this.p.props.spaceBefore = v
  }
  set spaceAfter(v: number) {
    this.p.props.spaceAfter = v
  }
  set leftIndent(v: number) {
    this.p.props.leftIndent = v
  }
  set rightIndent(v: number) {
    this.p.props.rightIndent = v
  }
  set firstLineIndent(v: number) {
    this.p.props.firstLineIndent = v
  }
  set outlineLevel(v: number) {
    this.p.props.outlineLevel = v
  }

  insertText(text: string, location: string) {
    this.pending.push(() => {
      this.p.text = text
      this.p.inserted = { text, location }
    })
  }

  insertParagraph(text: string, location: string) {
    const created = makeParagraph(text)
    this.pending.push(() => {
      const at = this.state.paragraphs.indexOf(this.p)
      this.state.paragraphs.splice(location === 'Before' ? at : at + 1, 0, created)
    })
    return new ParagraphProxy(created, this.pending, this.state)
  }

  insertTable(rowCount: number, columnCount: number, location: string, values: string[][]) {
    return makeTableProxy(this.state, this.pending, {
      rowCount,
      columnCount,
      values,
      location,
    })
  }

  insertBreak(breakType: string, location: string) {
    this.pending.push(() => this.state.breaks.push({ type: breakType, location }))
  }

  delete() {
    this.pending.push(() => {
      this.p.deleted = true
      const at = this.state.paragraphs.indexOf(this.p)
      if (at >= 0) this.state.paragraphs.splice(at, 1)
    })
  }

  startNewList() {
    const id = ++this.state.nextListId
    this.pending.push(() => (this.p.list = { id, level: 0 }))
    return {
      id,
      load: () => undefined,
      setLevelBullet: (level: number, bullet: string) =>
        this.pending.push(() => this.state.lists.push({ id, level, kind: 'bullet', bullet })),
      setLevelNumbering: (level: number, numbering: string) =>
        this.pending.push(() => this.state.lists.push({ id, level, kind: 'number', numbering })),
    }
  }

  attachToList(listId: number, level: number) {
    this.pending.push(() => (this.p.list = { id: listId, level }))
  }

  detachFromList() {
    this.pending.push(() => {
      this.p.list = undefined
      this.p.detachedFromList = true
    })
  }

  getRange() {
    const record: RangeRecord = { text: this.p.text, font: this.p.font, fields: [] }
    this.state.ranges.push(record)
    return makeRangeProxy(record, this.pending, this.state, [this])
  }
}

function makeRangeProxy(
  record: RangeRecord,
  pending: Array<() => void>,
  state: WordDocState,
  paragraphs: ParagraphProxy[],
  onReplace?: (text: string) => void,
) {
  return {
    load: () => undefined,
    get text() {
      return record.text
    },
    font: fontProxy(record.font),
    get paragraphs() {
      return { load: () => undefined, items: paragraphs }
    },
    set hyperlink(v: string) {
      record.hyperlink = v
      state.hyperlinks.push({ text: record.text, url: v })
    },
    set styleBuiltIn(v: string) {
      record.styleBuiltIn = v
    },
    insertText: (text: string, location: string) => {
      pending.push(() => {
        record.text = text
        if (onReplace) onReplace(text)
        else if (paragraphs[0]) paragraphs[0].p.text = text
        state.rangeEdits.push({ text, location })
      })
    },
    insertComment: (t: string) => pending.push(() => paragraphs[0]?.p.comments.push(t)),
    insertField: (location: string, fieldType: string) =>
      pending.push(() => record.fields.push(`${fieldType}@${location}`)),
  }
}

function makeTableProxy(
  state: WordDocState,
  pending: Array<() => void>,
  spec: { rowCount: number; columnCount: number; values: string[][]; location: string },
) {
  const table: MockTable = {
    rowCount: spec.rowCount,
    columnCount: spec.columnCount,
    values: spec.values,
    location: spec.location,
    headerRowCount: 0,
  }
  pending.push(() => state.tables.push(table))
  return {
    set headerRowCount(v: number) {
      table.headerRowCount = v
    },
    set style(v: string) {
      table.style = v
    },
    set styleBuiltIn(v: string) {
      table.styleBuiltIn = v
    },
  }
}

export interface MockTable {
  rowCount: number
  columnCount: number
  values: string[][]
  location: string
  headerRowCount: number
  style?: string
  styleBuiltIn?: string
}

export interface MockHeaderFooter {
  /** 'header' or 'footer'. */
  part: string
  cleared: boolean
  text: string
  alignment?: string
  fields: string[]
}

export interface WordDocState {
  paragraphs: MockParagraph[]
  selectionText: string
  /** Set by insertText on the selection. */
  selectionEdits: Array<{ text: string; location: string }>
  /** Set by insertText on any non-selection range. */
  rangeEdits: Array<{ text: string; location: string }>
  ooxml: string
  tables: MockTable[]
  breaks: Array<{ type: string; location: string }>
  hyperlinks: Array<{ text: string; url: string }>
  lists: Array<{ id: number; level: number; kind: string; bullet?: string; numbering?: string }>
  headerFooters: MockHeaderFooter[]
  /** Every range handed out, so tests can inspect formatting applied to search
   *  hits and to the body as a whole. */
  ranges: RangeRecord[]
  sectionCount: number
  nextListId: number
}

/**
 * Install a fake `Word` global. Returns the mutable document state so a test
 * can assert on what the adapter actually did.
 */
export function installWordMock(initial: Partial<WordDocState> = {}): WordDocState {
  const state: WordDocState = {
    paragraphs: initial.paragraphs ?? [],
    selectionText: initial.selectionText ?? '',
    selectionEdits: [],
    rangeEdits: [],
    ooxml: initial.ooxml ?? '<w:document/>',
    tables: [],
    breaks: [],
    hyperlinks: [],
    lists: [],
    headerFooters: [],
    ranges: [],
    sectionCount: initial.sectionCount ?? 1,
    nextListId: 0,
  }

  const InsertLocation = {
    replace: 'Replace',
    after: 'After',
    before: 'Before',
    start: 'Start',
    end: 'End',
  }
  // Only the enums the adapter actually dereferences at runtime. Alignment,
  // underline, and break values are assigned as string literals — the Office.js
  // typings accept them, and a literal cannot go stale against a mock.
  const UnderlineType = { single: 'Single', none: 'None', double: 'Double' }
  const ChangeTrackingMode = { off: 'Off', trackAll: 'TrackAll' }
  const ListBullet = { solid: 'Solid', hollow: 'Hollow' }
  const ListNumbering = { arabic: 'Arabic', upperRoman: 'UpperRoman' }
  const FieldType = { page: 'Page', numPages: 'NumPages' }

  async function run<T>(cb: (context: unknown) => Promise<T>): Promise<T> {
    const pending: Array<() => void> = []

    const proxyFor = new Map<MockParagraph, ParagraphProxy>()
    const paragraphProxy = (p: MockParagraph) => {
      // Stable identity per run, so an adapter that resolves the same paragraph
      // twice sees one object — as it would against the real API.
      let existing = proxyFor.get(p)
      if (!existing) {
        existing = new ParagraphProxy(p, pending, state)
        proxyFor.set(p, existing)
      }
      return existing
    }

    const paragraphsCollection = (items = state.paragraphs) => ({
      load: () => undefined,
      get items() {
        return items.map(paragraphProxy)
      },
    })

    function searchCollection(
      needle: string,
      options?: { matchCase?: boolean; matchWholeWord?: boolean },
    ) {
      const hits: ReturnType<typeof makeRangeProxy>[] = []
      for (const p of state.paragraphs) {
        const flags = options?.matchCase ? 'g' : 'gi'
        const escaped = needle.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
        const pattern = options?.matchWholeWord ? `\\b${escaped}\\b` : escaped
        const matches = [...p.text.matchAll(new RegExp(pattern, flags))]
        for (const m of matches) {
          const record: RangeRecord = { text: m[0], font: {}, fields: [] }
          state.ranges.push(record)
          hits.push(
            makeRangeProxy(record, pending, state, [paragraphProxy(p)], (replacement) => {
              // Each queued hit replaces the leftmost remaining occurrence.
              p.text = p.text.replace(
                new RegExp(pattern, options?.matchCase ? '' : 'i'),
                replacement,
              )
            }),
          )
        }
      }
      return { load: () => undefined, items: hits }
    }

    const bodyRecord: RangeRecord = { text: '', font: {}, fields: [] }
    const body = {
      load: () => undefined,
      get text() {
        return state.paragraphs.map((p) => p.text).join('\r')
      },
      get paragraphs() {
        return paragraphsCollection()
      },
      font: fontProxy(bodyRecord.font),
      getRange: () => {
        state.ranges.push(bodyRecord)
        return makeRangeProxy(bodyRecord, pending, state, state.paragraphs.map(paragraphProxy))
      },
      search: searchCollection,
      getOoxml: () => ({ value: state.ooxml }),
      insertOoxml: (xml: string) => pending.push(() => (state.ooxml = xml)),
      insertParagraph: (text: string, location: string) => {
        const created = makeParagraph(text)
        pending.push(() => {
          if (location === 'Start') state.paragraphs.unshift(created)
          else state.paragraphs.push(created)
        })
        return paragraphProxy(created)
      },
      insertTable: (rowCount: number, columnCount: number, location: string, values: string[][]) =>
        makeTableProxy(state, pending, { rowCount, columnCount, values, location }),
      insertBreak: (breakType: string, location: string) =>
        pending.push(() => state.breaks.push({ type: breakType, location })),
    }

    const selection = {
      load: () => undefined,
      get text() {
        return state.selectionText
      },
      get paragraphs() {
        // With no selection Word still reports the paragraph holding the cursor.
        return paragraphsCollection(state.paragraphs.slice(0, 1))
      },
      font: fontProxy({}),
      insertText: (text: string, location: string) =>
        pending.push(() => state.selectionEdits.push({ text, location })),
    }

    function headerFooterBody(part: string) {
      const record: MockHeaderFooter = { part, cleared: false, text: '', fields: [] }
      pending.push(() => state.headerFooters.push(record))
      return {
        clear: () => pending.push(() => (record.cleared = true)),
        insertParagraph: (text: string, _location: string) => {
          pending.push(() => (record.text = text))
          const paragraph = {
            set alignment(v: string) {
              record.alignment = v
            },
            insertText: (t: string) => pending.push(() => (record.text += t)),
            getRange: () => ({
              insertField: (_loc: string, fieldType: string) =>
                pending.push(() => record.fields.push(fieldType)),
            }),
          }
          return paragraph
        },
      }
    }

    const sections = {
      load: () => undefined,
      items: Array.from({ length: state.sectionCount }, () => ({
        getHeader: () => headerFooterBody('header'),
        getFooter: () => headerFooterBody('footer'),
      })),
    }

    const context = {
      document: {
        body,
        sections,
        getSelection: () => selection,
        load: () => undefined,
        changeTrackingMode: ChangeTrackingMode.off,
      },
      sync: async () => {
        // Flush queued mutations, exactly as a real sync() would.
        for (const op of pending.splice(0)) op()
      },
    }
    return cb(context)
  }

  ;(globalThis as Record<string, unknown>).Word = {
    run,
    InsertLocation,
    UnderlineType,
    ChangeTrackingMode,
    ListBullet,
    ListNumbering,
    FieldType,
  }
  return state
}

export interface ExcelSheetState {
  name: string
  /** Row-major, including the header row. */
  values: unknown[][]
  formulas: unknown[][]
  /** Records writes so tests can assert. */
  writes: Array<{ address: string; kind: 'values' | 'formulas'; payload: unknown }>
}

export function installExcelMock(
  initial: Partial<ExcelSheetState> & {
    selectedAddress?: string
    selectedSize?: [number, number]
  } = {},
): ExcelSheetState {
  const state: ExcelSheetState = {
    name: initial.name ?? 'Sheet1',
    values: initial.values ?? [],
    formulas: initial.formulas ?? [],
    writes: [],
  }
  const selectedSize = initial.selectedSize ?? [1, 1]

  function makeRange(address: string, values: unknown[][], formulas: unknown[][]) {
    return {
      address,
      rowCount: values.length,
      columnCount: values[0]?.length ?? 0,
      rowIndex: 0,
      columnIndex: 0,
      load: () => undefined,
      get values() {
        return values
      },
      set values(v: unknown[][]) {
        state.writes.push({ address, kind: 'values', payload: v })
      },
      get formulas() {
        return formulas
      },
      set formulas(v: unknown[][]) {
        state.writes.push({ address, kind: 'formulas', payload: v })
      },
      set numberFormat(v: unknown) {
        state.writes.push({ address, kind: 'values', payload: v })
      },
      format: { font: {}, fill: {} },
      getRow: (i: number) => makeRange(`row${i}`, values.slice(i), formulas.slice(i)),
      getResizedRange: (rows: number) =>
        makeRange(address, values.slice(0, rows + 1), formulas.slice(0, rows + 1)),
    }
  }

  async function run<T>(cb: (context: unknown) => Promise<T>): Promise<T> {
    const sheet = {
      name: state.name,
      load: () => undefined,
      getUsedRange: () => makeRange('A1:E9', state.values, state.formulas),
      getRange: (addr: string) => {
        // A cell being read for the clobber check must report its own formula.
        const existing = lookupFormula(state, addr)
        return makeRange(addr, [[lookupValue(state, addr)]], [[existing]])
      },
    }
    const context = {
      workbook: {
        worksheets: {
          load: () => undefined,
          items: [{ name: state.name }],
          getActiveWorksheet: () => sheet,
          getItem: () => sheet,
        },
        getSelectedRange: () =>
          makeRange(
            initial.selectedAddress ?? 'A1',
            state.values.slice(0, selectedSize[0]),
            state.formulas.slice(0, selectedSize[0]),
          ),
      },
      sync: async () => undefined,
    }
    return cb(context)
  }

  ;(globalThis as Record<string, unknown>).Excel = { run }
  return state
}

function colIndex(letter: string): number {
  let n = 0
  for (const c of letter.toUpperCase()) n = n * 26 + (c.charCodeAt(0) - 64)
  return n - 1
}

function parseAddr(addr: string): { row: number; col: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(addr.trim())
  if (!m) return null
  return { row: Number(m[2]) - 1, col: colIndex(m[1]) }
}

function lookupFormula(state: ExcelSheetState, addr: string): unknown {
  const p = parseAddr(addr)
  if (!p) return ''
  return state.formulas[p.row]?.[p.col] ?? ''
}

function lookupValue(state: ExcelSheetState, addr: string): unknown {
  const p = parseAddr(addr)
  if (!p) return ''
  return state.values[p.row]?.[p.col] ?? ''
}

export function uninstallOfficeMocks(): void {
  delete (globalThis as Record<string, unknown>).Word
  delete (globalThis as Record<string, unknown>).Excel
}
