// Word host adapter (P4.4–P4.10).
//
// Everything routes through a single Word.run per operation. Batching matters
// beyond performance: Word groups the proxy operations in one run into one undo
// record, so a multi-part edit reverts with a single Ctrl+Z. Splitting a logical
// change across two runs leaves the user pressing Ctrl+Z repeatedly and getting
// a half-edited document in between.

import type {
  ContextScope,
  DocumentAnchor,
  DocumentContext,
  Edit,
  TextFormatting,
  TextTarget,
} from '@openofficellm/shared'
import {
  estimateTokens,
  type ApplyResult,
  type HostAdapter,
  type SearchHit,
  type SnapshotPayload,
} from '@openofficellm/ui'
import {
  describeFormatting,
  needsParagraphProps,
  patchPageSetup,
  resolveBreak,
  resolveStyle,
  resolveUnderline,
  wordAlignment,
} from './wordFormat'

/** Past this, sending the whole body wastes context and usually overflows the
 *  model. Beyond it we send the outline plus the head of the document and let
 *  the model pull specific parts with search_document. */
const FULL_BODY_CHAR_LIMIT = 60_000

/** getOoxml on a very large document produces tens of megabytes and can wedge
 *  the pane. Past this we snapshot text only and say so. */
const OOXML_SNAPSHOT_CHAR_LIMIT = 400_000

/** Word's own find box refuses a longer needle, and so does body.search. */
const MAX_SEARCH_LENGTH = 255

/** A search that hits this many places is almost never what the model meant,
 *  and formatting them all is a document-wide change dressed up as a small one. */
const MAX_SEARCH_HITS = 400

function nowIso(): string {
  return new Date().toISOString()
}

async function readOutline(context: Word.RequestContext): Promise<string> {
  const paragraphs = context.document.body.paragraphs
  paragraphs.load('items/text,items/style,items/styleBuiltIn')
  await context.sync()
  const lines: string[] = []
  paragraphs.items.forEach((p, i) => {
    const style = String(p.style ?? '')
    // Heading paragraphs are the outline; body text is not.
    if (!/^Heading\s*\d/i.test(style) && style !== 'Title') return
    const level = style === 'Title' ? 0 : Number(style.replace(/\D+/g, '')) || 1
    const text = (p.text ?? '').trim()
    if (!text) return
    lines.push(`${'  '.repeat(level)}- [¶${i}] ${text}`)
  })
  return lines.join('\n')
}

/**
 * The body's paragraphs, loaded at most once per run.
 *
 * Almost every edit needs the paragraph collection, and each load costs a
 * `sync()` — which also flushes the mutations queued so far, splitting what
 * should be one undo record into several. Caching keeps a batch of formatting
 * edits down to a single flush. Structural edits invalidate it, because after
 * inserting or deleting a paragraph every index past the change has moved.
 */
class ParagraphCache {
  private items: Word.Paragraph[] | null = null

  constructor(private readonly context: Word.RequestContext) {}

  async all(): Promise<Word.Paragraph[]> {
    if (!this.items) {
      const collection = this.context.document.body.paragraphs
      collection.load('items')
      await this.context.sync()
      this.items = collection.items
    }
    return this.items
  }

  invalidate(): void {
    this.items = null
  }
}

/** Ranges to apply character formatting to, and the paragraphs behind them. */
interface FormatScope {
  ranges: Word.Range[]
  /** Empty when the caller said it only needs character formatting. */
  paragraphs: Word.Paragraph[]
  /** Human phrasing for the tool result, e.g. `3 occurrences of "budget"`. */
  label: string
}

type Resolved<T> = T | { error: string }

function isError<T>(v: Resolved<T>): v is { error: string } {
  return typeof v === 'object' && v !== null && 'error' in v
}

function listIndices(indices: number[]): string {
  if (indices.length === 1) return `paragraph ${indices[0]}`
  if (indices.length <= 4) return `paragraphs ${indices.join(', ')}`
  return `${indices.length} paragraphs`
}

/** Collect the paragraphs each range sits in, for paragraph-level properties. */
async function paragraphsOf(
  context: Word.RequestContext,
  ranges: Word.Range[],
): Promise<Word.Paragraph[]> {
  const collections = ranges.map((r) => {
    const paragraphs = r.paragraphs
    paragraphs.load('items')
    return paragraphs
  })
  await context.sync()
  return collections.flatMap((c) => c.items)
}

/**
 * Turn a target the model described into concrete Word objects.
 *
 * `wantParagraphs` is not an optimization detail — resolving paragraphs costs a
 * `sync()`, and a request that only sets bold has no business paying for it.
 */
async function resolveScope(
  context: Word.RequestContext,
  cache: ParagraphCache,
  target: TextTarget,
  wantParagraphs: boolean,
): Promise<Resolved<FormatScope>> {
  switch (target.kind) {
    case 'document': {
      const body = context.document.body
      return {
        ranges: [body.getRange()],
        paragraphs: wantParagraphs ? await cache.all() : [],
        label: 'the whole document',
      }
    }

    case 'paragraphs': {
      const indices = target.paragraphs ?? []
      if (indices.length === 0) return { error: 'No paragraph numbers were given.' }
      const all = await cache.all()
      const missing = indices.filter((i) => !Number.isInteger(i) || i < 0 || i >= all.length)
      if (missing.length > 0) {
        return {
          error: `Paragraph ${missing.join(', ')} does not exist — the document has ${all.length} (0–${all.length - 1}). Re-read the document to get current numbers.`,
        }
      }
      const paragraphs = indices.map((i) => all[i])
      return {
        ranges: paragraphs.map((p) => p.getRange()),
        paragraphs,
        label: listIndices(indices),
      }
    }

    case 'search': {
      const needle = (target.search ?? '').trim()
      if (!needle) return { error: 'Searching requires some text to look for.' }
      if (needle.length > MAX_SEARCH_LENGTH) {
        return {
          error: `Search text is too long (${needle.length} chars, max ${MAX_SEARCH_LENGTH}).`,
        }
      }
      const results = context.document.body.search(needle, {
        matchCase: target.matchCase === true,
        matchWholeWord: target.wholeWord === true,
      })
      results.load('items')
      await context.sync()
      let hits = results.items
      if (hits.length === 0) {
        return { error: `No text matching "${needle}" — nothing was changed.` }
      }
      if (target.firstOnly) hits = hits.slice(0, 1)
      if (hits.length > MAX_SEARCH_HITS) {
        return {
          error: `"${needle}" matches ${hits.length} places. Narrow the search, or target the document if that is really what you mean.`,
        }
      }
      return {
        ranges: hits,
        paragraphs: wantParagraphs ? await paragraphsOf(context, hits) : [],
        label: `${hits.length} occurrence${hits.length === 1 ? '' : 's'} of "${needle}"`,
      }
    }

    default: {
      const selection = context.document.getSelection()
      selection.load('text')
      await context.sync()
      if ((selection.text ?? '').trim()) {
        return {
          ranges: [selection],
          paragraphs: wantParagraphs ? await paragraphsOf(context, [selection]) : [],
          label: 'the selection',
        }
      }
      // An empty selection is an insertion point, not an error. Word itself
      // applies formatting to the paragraph holding the cursor, so doing the
      // same beats reporting success while changing nothing.
      const paragraphs = selection.paragraphs
      paragraphs.load('items')
      await context.sync()
      const first = paragraphs.items[0]
      if (!first) {
        return { error: 'Nothing is selected and the cursor is not in a paragraph.' }
      }
      return {
        ranges: [first.getRange()],
        paragraphs: [first],
        label: 'the paragraph at the cursor',
      }
    }
  }
}

function applyFont(font: Word.Font, f: TextFormatting): void {
  if (f.bold !== undefined) font.bold = f.bold
  if (f.italic !== undefined) font.italic = f.italic
  if (f.underline !== undefined) {
    const underline = resolveUnderline(f.underline)
    if (underline) font.underline = underline
  }
  if (f.strikeThrough !== undefined) font.strikeThrough = f.strikeThrough
  if (f.doubleStrikeThrough !== undefined) font.doubleStrikeThrough = f.doubleStrikeThrough
  if (f.superscript !== undefined) font.superscript = f.superscript
  if (f.subscript !== undefined) font.subscript = f.subscript
  if (f.smallCaps !== undefined) font.smallCaps = f.smallCaps
  if (f.allCaps !== undefined) font.allCaps = f.allCaps
  if (f.color !== undefined) font.color = f.color
  if (f.highlightColor !== undefined) {
    // Word documents null as "remove the highlight", but the typings only admit
    // a string. The cast is the API's own contract, not a shortcut.
    font.highlightColor = f.highlightColor as string
  }
  if (f.size !== undefined) font.size = f.size
  if (f.font !== undefined) font.name = f.font
}

function applyParagraphProps(p: Word.Paragraph, f: TextFormatting): void {
  if (f.style !== undefined) {
    const style = resolveStyle(f.style)
    // styleBuiltIn is locale-neutral; `style` takes the localized display name
    // and throws on anything else, which is what breaks formatting on a
    // non-English Word.
    if (style.builtIn) p.styleBuiltIn = style.builtIn
    else if (style.custom) p.style = style.custom
  }
  if (f.alignment !== undefined) p.alignment = wordAlignment(f.alignment)
  if (f.lineSpacing !== undefined) p.lineSpacing = f.lineSpacing
  if (f.spaceBefore !== undefined) p.spaceBefore = f.spaceBefore
  if (f.spaceAfter !== undefined) p.spaceAfter = f.spaceAfter
  if (f.leftIndent !== undefined) p.leftIndent = f.leftIndent
  if (f.rightIndent !== undefined) p.rightIndent = f.rightIndent
  if (f.firstLineIndent !== undefined) p.firstLineIndent = f.firstLineIndent
  if (f.outlineLevel !== undefined) p.outlineLevel = f.outlineLevel
}

/** Look up a list of paragraph indices, refusing the whole edit if any of them
 *  is out of range — a partially applied structural change is worse than none. */
async function pickParagraphs(
  cache: ParagraphCache,
  indices: number[],
  verb: string,
): Promise<Resolved<Word.Paragraph[]>> {
  if (indices.length === 0) return { error: `No paragraphs given to ${verb}.` }
  const all = await cache.all()
  const missing = indices.filter((i) => !Number.isInteger(i) || i < 0 || i >= all.length)
  if (missing.length > 0) {
    return {
      error: `Cannot ${verb} paragraph ${missing.join(', ')}: the document has ${all.length} (0–${all.length - 1}).`,
    }
  }
  return indices.map((i) => all[i])
}

/** Resolve an insertion point to the paragraph it attaches to, or null for
 *  "the body itself" (start/end of the document). */
async function resolveAnchor(
  context: Word.RequestContext,
  cache: ParagraphCache,
  at: DocumentAnchor,
): Promise<Resolved<{ paragraph: Word.Paragraph | null; atStart: boolean; label: string }>> {
  if (at === 'start') return { paragraph: null, atStart: true, label: 'at the start' }
  if (at === 'end') return { paragraph: null, atStart: false, label: 'at the end' }

  if (at === 'selection') {
    const selection = context.document.getSelection()
    const paragraphs = selection.paragraphs
    paragraphs.load('items')
    await context.sync()
    const items = paragraphs.items
    const last = items[items.length - 1]
    if (!last) return { error: 'The cursor is not in a paragraph.' }
    return { paragraph: last, atStart: false, label: 'after the selection' }
  }

  const all = await cache.all()
  if (!Number.isInteger(at) || at < 0 || at >= all.length) {
    return {
      error: `Paragraph ${at} does not exist — the document has ${all.length} (0–${all.length - 1}).`,
    }
  }
  return { paragraph: all[at], atStart: false, label: `after paragraph ${at}` }
}

export class WordAdapter implements HostAdapter {
  readonly host = 'word' as const

  async getContext(scope: ContextScope): Promise<DocumentContext> {
    if (scope === 'none') {
      return { host: 'word', scope, text: '', tokenEstimate: 0 }
    }

    return Word.run(async (context) => {
      const body = context.document.body
      const selection = context.document.getSelection()
      selection.load('text')
      body.load('text')
      await context.sync()

      const selectionText = (selection.text ?? '').trim()
      const bodyText = body.text ?? ''

      // A "selection" scope with nothing selected is the single most common
      // case — the user clicks into the pane and types without selecting
      // anything first. Silently falling back to the document is what makes
      // "summarize this" work rather than answering about an empty string.
      let effectiveScope: ContextScope = scope
      let text: string
      if (scope === 'selection') {
        if (selectionText) {
          text = selectionText
        } else {
          effectiveScope = 'document'
          text = bodyText
        }
      } else if (scope === 'paragraph') {
        const paras = selection.paragraphs
        paras.load('items/text')
        await context.sync()
        const around = paras.items.map((p) => p.text ?? '').join('\n')
        if (around.trim()) {
          text = around
        } else {
          effectiveScope = 'document'
          text = bodyText
        }
      } else {
        text = bodyText
      }

      const outline = await readOutline(context)

      let finalText = text
      let truncatedNote = ''
      if (effectiveScope === 'document' && finalText.length > FULL_BODY_CHAR_LIMIT) {
        finalText = finalText.slice(0, FULL_BODY_CHAR_LIMIT)
        truncatedNote = '\n\n[document truncated — use search_document to find specific passages]'
      }
      finalText += truncatedNote

      return {
        host: 'word' as const,
        scope: effectiveScope,
        text: finalText,
        outline: outline || undefined,
        tokenEstimate: estimateTokens(finalText + outline),
      }
    })
  }

  async applyEdits(edits: Edit[]): Promise<ApplyResult> {
    if (edits.length === 0) return { ok: true, summary: 'No changes to apply.' }

    return Word.run(async (context) => {
      const cache = new ParagraphCache(context)
      const applied: string[] = []

      for (const edit of edits) {
        const outcome = await applyOne(context, cache, edit)
        if (isError(outcome)) return { ok: false, summary: outcome.error }
        applied.push(outcome)
      }

      await context.sync()
      return { ok: true, summary: `Applied: ${applied.join('; ')}.` }
    })
  }

  async snapshot(): Promise<SnapshotPayload> {
    return Word.run(async (context) => {
      const body = context.document.body
      body.load('text')
      await context.sync()
      const textLength = (body.text ?? '').length

      if (textLength > OOXML_SNAPSHOT_CHAR_LIMIT) {
        return {
          id: `snap_${Date.now().toString(36)}`,
          host: 'word' as const,
          createdAt: nowIso(),
          sizeBytes: textLength,
          // Explicitly marked so restore() can refuse rather than silently
          // replacing a formatted document with plain text.
          data: { kind: 'too-large' as const },
        }
      }

      const ooxml = body.getOoxml()
      await context.sync()
      return {
        id: `snap_${Date.now().toString(36)}`,
        host: 'word' as const,
        createdAt: nowIso(),
        sizeBytes: ooxml.value.length,
        data: { kind: 'ooxml' as const, ooxml: ooxml.value },
      }
    })
  }

  async restore(snapshot: SnapshotPayload): Promise<void> {
    const data = snapshot.data as { kind: string; ooxml?: string }
    if (data?.kind !== 'ooxml' || typeof data.ooxml !== 'string') {
      throw new Error(
        'This document was too large to snapshot, so the run cannot be reverted automatically. Use Ctrl+Z.',
      )
    }
    await Word.run(async (context) => {
      context.document.body.insertOoxml(data.ooxml!, Word.InsertLocation.replace)
      await context.sync()
    })
  }

  async search(query: string, limit = 20): Promise<SearchHit[]> {
    if (!query.trim()) return []
    return Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs
      paragraphs.load('items/text')
      await context.sync()
      const needle = query.toLowerCase()
      const hits: SearchHit[] = []
      for (let i = 0; i < paragraphs.items.length && hits.length < limit; i++) {
        const text = paragraphs.items[i].text ?? ''
        if (text.toLowerCase().includes(needle)) {
          hits.push({ location: String(i), text: text.slice(0, 500) })
        }
      }
      return hits
    })
  }

  /**
   * Report the formatting already in force on a target.
   *
   * Without this the model can only write formatting, never read it, so it
   * cannot match a document's existing look or check that an edit landed — it
   * has to guess and then claim success.
   */
  async readFormatting(target: TextTarget): Promise<string> {
    return Word.run(async (context) => {
      const cache = new ParagraphCache(context)
      const scope = await resolveScope(context, cache, target, true)
      if (isError(scope)) return scope.error

      const range = scope.ranges[0]
      const paragraph = scope.paragraphs[0]
      range.load('text')
      range.font.load(
        'bold,italic,underline,color,highlightColor,name,size,strikeThrough,superscript,subscript',
      )
      if (paragraph) {
        paragraph.load(
          'style,styleBuiltIn,alignment,lineSpacing,spaceBefore,spaceAfter,leftIndent,rightIndent,firstLineIndent,outlineLevel,isListItem',
        )
      }
      await context.sync()

      const font = range.font
      const lines = [
        `Target: ${scope.label}`,
        `Text: ${JSON.stringify((range.text ?? '').slice(0, 120))}`,
        `Font: ${font.name ?? '(inherited)'} ${font.size ?? '?'}pt`,
        `Bold: ${String(font.bold)}  Italic: ${String(font.italic)}  Underline: ${String(font.underline)}`,
        `Strikethrough: ${String(font.strikeThrough)}  Superscript: ${String(font.superscript)}  Subscript: ${String(font.subscript)}`,
        `Colour: ${font.color ?? '(inherited)'}  Highlight: ${font.highlightColor ?? 'none'}`,
      ]
      if (paragraph) {
        lines.push(
          `Style: ${paragraph.style ?? '(none)'} (built-in id: ${paragraph.styleBuiltIn ?? 'none'})`,
          `Alignment: ${paragraph.alignment ?? '(inherited)'}  Outline level: ${paragraph.outlineLevel ?? '?'}`,
          `Spacing: ${paragraph.spaceBefore ?? 0}pt before, ${paragraph.spaceAfter ?? 0}pt after, line ${paragraph.lineSpacing ?? '?'}pt`,
          `Indents: left ${paragraph.leftIndent ?? 0}pt, right ${paragraph.rightIndent ?? 0}pt, first line ${paragraph.firstLineIndent ?? 0}pt`,
          `List item: ${String(paragraph.isListItem)}`,
        )
      }
      if (scope.ranges.length > 1) {
        lines.push(`(${scope.ranges.length} places match; the formatting above is the first one.)`)
      }
      return lines.join('\n')
    })
  }

  /** Whether track-changes is on. Surfaced so the pane can tell the user their
   *  edits will land as tracked revisions (P4.9). */
  async isTrackingChanges(): Promise<boolean> {
    try {
      return await Word.run(async (context) => {
        const doc = context.document
        doc.load('changeTrackingMode')
        await context.sync()
        return doc.changeTrackingMode !== Word.ChangeTrackingMode.off
      })
    } catch {
      // Requires WordApi 1.4; older hosts simply do not report it.
      return false
    }
  }
}

/**
 * Apply one edit, returning a description of what it did.
 *
 * Split out of `applyEdits` so each kind stays readable; the caller owns the
 * single trailing `sync()` that makes the batch one undo record.
 */
async function applyOne(
  context: Word.RequestContext,
  cache: ParagraphCache,
  edit: Edit,
): Promise<Resolved<string>> {
  switch (edit.kind) {
    case 'replaceSelection': {
      const range = context.document.getSelection()
      // insertText('Replace') on the range, not body-level replacement:
      // it keeps the surrounding paragraph's formatting intact.
      range.insertText(edit.text, Word.InsertLocation.replace)
      return `replaced the selection (${edit.text.length} chars)`
    }

    case 'insertAfter': {
      context.document.getSelection().insertText(edit.text, Word.InsertLocation.after)
      return `inserted ${edit.text.length} chars after the selection`
    }

    case 'insertBefore': {
      context.document.getSelection().insertText(edit.text, Word.InsertLocation.before)
      return `inserted ${edit.text.length} chars before the selection`
    }

    case 'replaceRange': {
      // rangeId is a paragraph index — the same id search_document and the
      // outline hand out, so the model can round-trip them.
      const idx = Number(edit.rangeId)
      if (!Number.isInteger(idx) || idx < 0) {
        return { error: `Invalid paragraph reference: ${edit.rangeId}` }
      }
      const all = await cache.all()
      const target = all[idx]
      if (!target) {
        return { error: `Paragraph ${idx} does not exist (document has ${all.length}).` }
      }
      target.insertText(edit.text, Word.InsertLocation.replace)
      return `replaced paragraph ${idx}`
    }

    case 'addComment': {
      const idx = Number(edit.rangeId)
      const all = await cache.all()
      const target = Number.isInteger(idx) ? all[idx] : undefined
      if (!target) return { error: `Cannot comment: paragraph ${edit.rangeId} not found.` }
      target.getRange().insertComment(edit.text)
      return `commented on paragraph ${idx}`
    }

    case 'applyFormatting': {
      // The pre-formatText shape, still reachable from proposals persisted by an
      // earlier build. Route it onto the same code path rather than keeping two
      // formatters that drift apart.
      const idx = Number(edit.rangeId)
      if (!Number.isInteger(idx) || idx < 0) {
        return { error: `Cannot format: paragraph ${edit.rangeId} not found.` }
      }
      return applyOne(context, cache, {
        kind: 'formatText',
        target: { kind: 'paragraphs', paragraphs: [idx] },
        formatting: edit.formatting as TextFormatting,
      })
    }

    case 'formatText': {
      const scope = await resolveScope(
        context,
        cache,
        edit.target,
        needsParagraphProps(edit.formatting),
      )
      if (isError(scope)) return scope
      for (const range of scope.ranges) applyFont(range.font, edit.formatting)
      for (const paragraph of scope.paragraphs) applyParagraphProps(paragraph, edit.formatting)
      const what = describeFormatting(edit.formatting)
      return `formatted ${scope.label}${what ? ` (${what})` : ''}`
    }

    case 'insertParagraph': {
      const anchor = await resolveAnchor(context, cache, edit.at)
      if (isError(anchor)) return anchor
      const created = anchor.paragraph
        ? anchor.paragraph.insertParagraph(edit.text, Word.InsertLocation.after)
        : context.document.body.insertParagraph(
            edit.text,
            anchor.atStart ? Word.InsertLocation.start : Word.InsertLocation.end,
          )
      if (edit.style) {
        const style = resolveStyle(edit.style)
        if (style.builtIn) created.styleBuiltIn = style.builtIn
        else if (style.custom) created.style = style.custom
      }
      // Every index past the insertion point has shifted.
      cache.invalidate()
      return `inserted a paragraph ${anchor.label}`
    }

    case 'deleteParagraphs': {
      // Descending, so each delete leaves the indices below it untouched.
      const ordered = [...edit.paragraphs].sort((a, b) => b - a)
      const paragraphs = await pickParagraphs(cache, ordered, 'delete')
      if (isError(paragraphs)) return paragraphs
      for (const p of paragraphs) p.delete()
      cache.invalidate()
      return `deleted ${listIndices([...edit.paragraphs].sort((a, b) => a - b))}`
    }

    case 'setList': {
      const paragraphs = await pickParagraphs(cache, edit.paragraphs, 'make a list from')
      if (isError(paragraphs)) return paragraphs
      if (edit.listType === 'none') {
        for (const p of paragraphs) p.detachFromList()
        return `removed list formatting from ${listIndices(edit.paragraphs)}`
      }
      // Word's list levels are zero-based; the tool takes 1 for the top level.
      const level = Math.max(0, Math.min(8, (edit.level ?? 1) - 1))
      const list = paragraphs[0].startNewList()
      list.load('id')
      await context.sync()
      if (edit.listType === 'bullet') list.setLevelBullet(level, Word.ListBullet.solid)
      else list.setLevelNumbering(level, Word.ListNumbering.arabic)
      // The rest join the list the first one started, so they number as one
      // list rather than restarting at 1 apiece.
      for (const p of paragraphs.slice(1)) p.attachToList(list.id, level)
      const kind = edit.listType === 'bullet' ? 'bulleted' : 'numbered'
      return `made ${listIndices(edit.paragraphs)} a ${kind} list`
    }

    case 'insertTable': {
      const rows = edit.rows.filter((r) => Array.isArray(r))
      if (rows.length === 0) return { error: 'A table needs at least one row.' }
      const columnCount = Math.max(...rows.map((r) => r.length))
      if (columnCount === 0) return { error: 'A table needs at least one column.' }
      // Word rejects a values grid that is not exactly rowCount × columnCount.
      const values = rows.map((r) =>
        Array.from({ length: columnCount }, (_, i) => String(r[i] ?? '')),
      )
      const anchor = await resolveAnchor(context, cache, edit.at)
      if (isError(anchor)) return anchor
      const table = anchor.paragraph
        ? anchor.paragraph.insertTable(
            values.length,
            columnCount,
            Word.InsertLocation.after,
            values,
          )
        : context.document.body.insertTable(
            values.length,
            columnCount,
            anchor.atStart ? Word.InsertLocation.start : Word.InsertLocation.end,
            values,
          )
      table.headerRowCount = edit.headerRow === false ? 0 : 1
      const style = resolveStyle(edit.style ?? 'GridTable4_Accent1')
      if (style.builtIn) table.styleBuiltIn = style.builtIn
      else if (style.custom) table.style = style.custom
      cache.invalidate()
      return `inserted a ${values.length}×${columnCount} table ${anchor.label}`
    }

    case 'insertBreak': {
      const breakType = resolveBreak(edit.breakType)
      if (!breakType) {
        return { error: `Unknown break type "${edit.breakType}". Use page, line, or section.` }
      }
      const anchor = await resolveAnchor(context, cache, edit.at)
      if (isError(anchor)) return anchor
      if (anchor.paragraph) {
        anchor.paragraph.insertBreak(breakType, Word.InsertLocation.after)
      } else {
        context.document.body.insertBreak(
          breakType,
          anchor.atStart ? Word.InsertLocation.start : Word.InsertLocation.end,
        )
      }
      cache.invalidate()
      return `inserted a ${edit.breakType} break ${anchor.label}`
    }

    case 'insertHyperlink': {
      const url = edit.url.trim()
      if (!/^(https?:|mailto:|file:)/i.test(url)) {
        return { error: `"${url}" is not a link. Use an http(s):// or mailto: address.` }
      }
      const scope = await resolveScope(context, cache, edit.target, false)
      if (isError(scope)) return scope
      for (const range of scope.ranges) {
        // insertText returns the range covering the new text; the original
        // proxy no longer describes what is on the page, so the link has to go
        // on the returned one or it lands on stale coordinates.
        const linked = edit.text ? range.insertText(edit.text, Word.InsertLocation.replace) : range
        linked.hyperlink = url
        linked.styleBuiltIn = 'Hyperlink'
      }
      return `linked ${scope.label} to ${url}`
    }

    case 'replaceAll': {
      const needle = edit.find.trim()
      if (!needle) return { error: 'Find text is required.' }
      if (needle.length > MAX_SEARCH_LENGTH) {
        return {
          error: `Find text is too long (${needle.length} chars, max ${MAX_SEARCH_LENGTH}).`,
        }
      }
      const results = context.document.body.search(needle, {
        matchCase: edit.matchCase === true,
        matchWholeWord: edit.wholeWord === true,
      })
      results.load('items')
      await context.sync()
      const hits = results.items
      if (hits.length === 0)
        return { error: `No occurrences of "${needle}" — nothing was changed.` }
      for (const hit of hits) hit.insertText(edit.replace, Word.InsertLocation.replace)
      cache.invalidate()
      return `replaced ${hits.length} occurrence${hits.length === 1 ? '' : 's'} of "${needle}"`
    }

    case 'setHeaderFooter': {
      const sections = context.document.sections
      sections.load('items')
      await context.sync()
      if (sections.items.length === 0) return { error: 'This document has no sections.' }
      for (const section of sections.items) {
        const part =
          edit.part === 'header' ? section.getHeader('Primary') : section.getFooter('Primary')
        part.clear()
        const paragraph = part.insertParagraph(edit.text, Word.InsertLocation.start)
        if (edit.alignment) paragraph.alignment = wordAlignment(edit.alignment)
        if (edit.pageNumber) {
          // A literal page number would be wrong on every page but one, so this
          // has to be a field. Needs a recent Word build (WordApi 1.5).
          if (edit.text) paragraph.insertText(' ', Word.InsertLocation.end)
          paragraph.getRange().insertField(Word.InsertLocation.end, Word.FieldType.page)
        }
      }
      const count = sections.items.length
      return `set the ${edit.part} on ${count} section${count === 1 ? '' : 's'}`
    }

    case 'setPageSetup': {
      const body = context.document.body
      body.load('text')
      await context.sync()
      if ((body.text ?? '').length > OOXML_SNAPSHOT_CHAR_LIMIT) {
        return {
          error:
            'This document is too large to change page settings safely. Change them in Word directly.',
        }
      }
      const ooxml = body.getOoxml()
      await context.sync()
      const patched = patchPageSetup(ooxml.value, edit.setup)
      if ('error' in patched) return { error: patched.error }
      body.insertOoxml(patched.xml, Word.InsertLocation.replace)
      cache.invalidate()
      return `set ${patched.changed.join(' and ')}`
    }

    default:
      return { error: `${(edit as Edit).kind} is not supported in Word.` }
  }
}
