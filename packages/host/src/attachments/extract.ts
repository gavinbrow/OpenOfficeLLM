// Text extraction dispatch by MIME type / file extension.
//
// The host extracts text from every text-kind attachment at upload time, so
// prompt assembly can fold it into the system prompt without re-reading the
// file (and without the extraction library being loaded more than once per
// file). Image attachments are returned as `kind: 'image'` with empty text —
// their bytes are forwarded as vision content blocks by the prompt builder
// instead, which is a different phase.
//
// The extractor dispatch is intentionally a flat switch over extension and
// MIME prefix rather than a plugin registry: the set of supported types is
// small, stable, and each needs bespoke handling (utf-8 read, pdfjs-dist,
// mammoth, exceljs, jszip for pptx). A registry would obscure the fact that
// each branch is a one-off.
//
// All extractors are loaded via dynamic `import()` so the host binary does
// not pay for a PDF parser, an xlsx reader and a pptx unzipper at startup
// when the user is only ever going to attach a .txt. The dynamic imports are
// also what keeps the SEA bundle from pulling these into the single-file
// executable when they are unused.

import path from 'node:path'
import type { AttachmentKind } from './store.js'

/** Cap extracted text at the same character budget the prompt builder uses
 *  for document context. A token estimate that is wrong in the unsafe
 *  direction costs the user a failed request, and character counts are
 *  exact — so we cap on characters, not tokens. Mirrors prompt.ts:71. */
const MAX_CONTEXT_CHARS = 48_000

export class AttachmentError extends Error {
  readonly code: string
  constructor(code: string, message: string) {
    super(message)
    this.name = 'AttachmentError'
    this.code = code
  }
}

export interface ExtractResult {
  text: string
  kind: AttachmentKind
  truncated: boolean
}

/** Extensions read as utf-8 text without any parsing. Source code, config
 *  formats, and structured-but-human-readable files all live here. The list
 *  is intentionally permissive: a .log is just text, and treating a .json as
 *  text means the model sees the raw file rather than a re-serialized
 *  approximation of it. */
const TEXT_EXTENSIONS = new Set([
  '.txt', '.md', '.markdown', '.csv', '.tsv', '.json', '.xml', '.yaml', '.yml',
  '.html', '.htm', '.svg', '.log', '.ini', '.env', '.cfg', '.conf', '.toml',
  '.sql',
  // Source code.
  '.ts', '.tsx', '.js', '.jsx', '.mjs', '.cjs', '.py', '.rb', '.go', '.rs',
  '.java', '.c', '.h', '.cpp', '.cc', '.hpp', '.cs', '.php', '.sh', '.bash',
  '.zsh', '.ps1', '.bat', '.cmd', '.swift', '.kt', '.scala', '.lua', '.pl',
  '.r', '.dart', '.vue', '.svelte', '.astro', '.graphql', '.gql', '.proto',
  // Dotfiles without a suffix: `Makefile`, `Dockerfile`, `.gitignore`, etc.
  // are matched by lowercase basename below.
])

/** Lowercase basenames that count as text even though they have no extension. */
const TEXT_BASENAMES = new Set([
  'makefile', 'dockerfile', '.gitignore', '.editorconfig', '.env',
])

function extOf(fileName: string): string {
  return path.extname(fileName).toLowerCase()
}

function isTextByExtension(fileName: string): boolean {
  const ext = extOf(fileName)
  if (TEXT_EXTENSIONS.has(ext)) return true
  const base = path.basename(fileName).toLowerCase()
  return TEXT_BASENAMES.has(base)
}

function truncate(text: string): { text: string; truncated: boolean } {
  if (text.length <= MAX_CONTEXT_CHARS) return { text, truncated: false }
  return { text: `${text.slice(0, MAX_CONTEXT_CHARS)}\n…[truncated]`, truncated: true }
}

async function extractPdf(buffer: Buffer): Promise<string> {
  // pdfjs-dist (the maintained distribution of pdf.js) replaces `pdf-parse`,
  // whose bundled pdf.js v1.10.100 throws `bad XRef entry` on every PDF under
  // Node 22 (pdf-parse was last published 8 years ago and is a dead end). The
  // `legacy` build is the Node-friendly one — the main `build/pdf.mjs` entry
  // references browser globals like `DOMMatrix` and fails to even import on
  // Node. Dynamic import matches the rest of this file (mammoth, exceljs,
  // jszip) and keeps the parser out of the startup path until a PDF is
  // actually attached.
  //
  // Two globals the legacy build expects have to be set up before the import:
  //
  // 1. `DOMMatrix` / `Path2D`. The legacy build polyfills these from
  //    `@napi-rs/canvas` at module init via
  //    `require(createRequire(import.meta.url))`. That works in the dev/ESM
  //    path, but the host is bundled to a CJS SEA blob by tsup, which
  //    rewrites `import.meta.url` to an empty object — so the polyfill's
  //    `createRequire(undefined)` throws, the canvas require is skipped, and
  //    pdf.js's own top-level `new DOMMatrix()` then throws `DOMMatrix is not
  //    defined` before extraction even starts. We install permissive stubs
  //    ourselves. The host only does text extraction — no rendering — so a
  //    stub that satisfies the top-level `new DOMMatrix()` / `new Path2D()`
  //    is all pdf.js needs to import; the stubs accept any constructor args
  //    and otherwise do nothing. Depending on `@napi-rs/canvas` instead would
  //    pull a native `.node` module into the tsup bundle (which
  //    `noExternal: [/.*/]` forces) and break the build with unresolved
  //    platform-specific binaries.
  //
  // 2. `globalThis.pdfjsWorker.WorkerMessageHandler`. pdf.js parses PDFs in a
  //    worker thread in the browser; on Node it runs a "fake worker" on the
  //    main thread. The fake-worker loader does `await import(workerSrc)`
  //    where `workerSrc` defaults to "./pdf.worker.mjs" — a path relative to
  //    the bundle, which does not exist in the CJS SEA output. pdf.js
  //    short-circuits that loader if `globalThis.pdfjsWorker.WorkerMessage
  //    Handler` is already set, so we import the worker module ourselves and
  //    inject it. Both `pdf.mjs` and `pdf.worker.mjs` are dynamic-imported
  //    with string-literal specifiers, so tsup bundles them into the CJS blob
  //    (no runtime file resolution needed).
  //
  // The `any`-cast on `globalThis` is because Node's `@types/node` does not
  // declare `DOMMatrix`/`Path2D`/`pdfjsWorker` (they are DOM-lib / pdf.js
  // internals) and the host tsconfig does not pull in the DOM lib, so a typed
  // reference would not compile.
  const g = globalThis as unknown as Record<string, unknown>
  if (typeof g.DOMMatrix === 'undefined') {
    class DOMMatrixStub {
      constructor(..._args: unknown[]) {}
    }
    g.DOMMatrix = DOMMatrixStub
  }
  if (typeof g.Path2D === 'undefined') {
    class Path2DStub {
      constructor(..._args: unknown[]) {}
    }
    g.Path2D = Path2DStub
  }
  if (!g.pdfjsWorker) {
    const worker = await import('pdfjs-dist/legacy/build/pdf.worker.mjs')
    g.pdfjsWorker = worker
  }
  const { getDocument, VerbosityLevel } = await import('pdfjs-dist/legacy/build/pdf.mjs')
  // `data` must be a plain Uint8Array: Buffer is a Uint8Array subclass, but
  // pdf.js transfers the buffer to the (fake, main-thread) worker and a Node
  // Buffer's extra methods survive the transfer as dead weight. Wrapping in a
  // fresh Uint8Array view costs nothing and sidesteps that. `verbosity: ERRORS`
  // silences the "Indexing all PDF objects" / "standardFontDataUrl" notices
  // pdf.js emits for minimal PDFs without embedded fonts — they are not
  // errors and would otherwise pollute the host log on every upload.
  const loadingTask = getDocument({
    data: new Uint8Array(buffer),
    verbosity: VerbosityLevel.ERRORS,
  })
  const doc = await loadingTask.promise
  const parts: string[] = []
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i)
    const content = await page.getTextContent()
    // Each text item is a `{ str, dir, transform, ... }`; marked-content items
    // from `includeMarkedContent` (off by default) carry no `str`, so the
    // `it.str ?? ''` guard keeps a future flag flip from crashing extraction.
    const line = content.items
      .map((it) => ('str' in it ? (it as { str: string }).str : ''))
      .join('')
    parts.push(line)
  }
  await loadingTask.destroy()
  return parts.join('\n')
}

async function extractDocx(buffer: Buffer): Promise<string> {
  const mammoth = await import('mammoth')
  // mammoth's Node input takes `{ buffer }` (its Browser input takes
  // `{ arrayBuffer }`). We are on Node, so the buffer form is the native one
  // and avoids an ArrayBuffer copy.
  const result = await mammoth.extractRawText({ buffer })
  return result.value ?? ''
}

async function extractXlsx(buffer: Buffer): Promise<string> {
  const exceljs = await import('exceljs')
  const workbook = new exceljs.Workbook()
  // exceljs's .d.ts declares its OWN global `interface Buffer extends ArrayBuffer`
  // (line 1 of index.d.ts) and types `xlsx.load(buffer: Buffer)` against it.
  // That fake `Buffer` requires `slice` to return an `ArrayBuffer`, which
  // Node's real `Buffer` (whose `slice` returns `Uint8Array`) never satisfies
  // — so no Node Buffer can be passed without a cast. The runtime accepts a
  // Node Buffer fine; cast through `unknown` to the exceljs-declared `Buffer`.
  await workbook.xlsx.load(buffer as unknown as Parameters<typeof workbook.xlsx.load>[0])
  const sheets: string[] = []
  workbook.eachSheet((sheet) => {
    const name = sheet.name ?? 'Sheet'
    const rows: string[] = []
    sheet.eachRow({ includeEmpty: false }, (row) => {
      // `row.values` is `CellValue[] | { [key]: CellValue }` and is 1-indexed:
      // index 0 is always undefined. The previous `Object.values(values)`
      // (or an un-sliced array) included that dummy slot, which stringified
      // to '' and joined into a leading comma on every row. Slice it off so
      // the first real cell is the first column. Complex cell values (rich
      // text, formula results) fall back to `String(v)`, which is ugly but
      // never empty.
      const values = row.values
      const arr = Array.isArray(values) ? values.slice(1) : Object.values(values)
      const cells = arr.map((v) => (v == null ? '' : String(v)))
      rows.push(cells.join(','))
    })
    sheets.push(`Sheet: ${name}\n${rows.join('\n')}`)
  })
  return sheets.join('\n\n')
}

async function extractPptx(buffer: Buffer): Promise<string> {
  const JSZip = (await import('jszip')).default
  const zip = await JSZip.loadAsync(buffer)
  // Slide XML lives at `ppt/slides/slideN.xml` (zero-padded or not). Sort by
  // the numeric slide index parsed out of the filename so the text comes
  // back in presentation order rather than zip insertion order.
  const slideEntries = Object.keys(zip.files)
    .filter((p) => /^ppt\/slides\/slide\d+\.xml$/i.test(p))
    .map((p) => ({ path: p, index: parseInt(/\d+/.exec(path.basename(p))?.[0] ?? '0', 10) }))
    .sort((a, b) => a.index - b.index)
  const slides: string[] = []
  for (const entry of slideEntries) {
    const xml = await zip.files[entry.path].async('string')
    // `<a:t>` is the DrawingML run-text element — the actual visible text on
    // the slide. A full XML walk would also pull speaker notes and table
    // cells, but the regex is robust against the namespace prefix variants
    // Office emits and avoids pulling in an XML parser for one tag. The
    // opening tag can carry attributes (notably `xml:space="preserve"`), so
    // `\b[^>]*` admits any attribute set before the `>`; the text capture is
    // a second pass on each full match so the attribute payload does not leak
    // into the extracted run text.
    const matches = xml.match(/<a:t\b[^>]*>([^<]*)<\/a:t>/g) ?? []
    const text = matches
      .map((m) => {
        const inner = m.match(/<a:t\b[^>]*>([^<]*)<\/a:t>/)
        return inner ? inner[1] : ''
      })
      .join('')
    slides.push(text)
  }
  return slides.join('\n\n')
}

export async function extractText(
  buffer: Buffer,
  fileName: string,
  mimeType: string,
): Promise<ExtractResult> {
  // Images never get text extraction here — they are forwarded as vision
  // content blocks by the prompt builder, which is a separate path.
  if (mimeType.toLowerCase().startsWith('image/')) {
    return { text: '', kind: 'image', truncated: false }
  }

  // Any `text/*` MIME type is readable as utf-8 regardless of extension.
  const isTextMime = mimeType.toLowerCase().startsWith('text/')
  const ext = extOf(fileName)

  let raw: string
  if (isTextMime || isTextByExtension(fileName)) {
    raw = buffer.toString('utf-8')
  } else if (ext === '.pdf') {
    raw = await extractPdf(buffer)
  } else if (ext === '.docx') {
    raw = await extractDocx(buffer)
  } else if (ext === '.xlsx') {
    raw = await extractXlsx(buffer)
  } else if (ext === '.pptx') {
    raw = await extractPptx(buffer)
  } else {
    throw new AttachmentError(
      'unsupported_file_type',
      `Unsupported file type: ${ext || '(no extension)'}`,
    )
  }

  const { text, truncated } = truncate(raw)
  return { text, kind: 'text', truncated }
}