// Builds the system prompt for a chat turn.
//
// Before this existed the `context` field of ChatRequest was accepted, typed,
// carried across the wire — and then dropped on the floor by every adapter. The
// model was told nothing about the document, so "what is in this document"
// truthfully answered that no document had been provided.

import type { ChatMessage, ChatRequest, ContentBlock, DocumentContext, EditMode } from '@openofficellm/shared'
import { getMeta, getBytes, extractTextOcr } from './attachments/index.js'
import { logger } from './logging.js'

const MODE_GUIDANCE: Record<EditMode, string> = {
  propose: [
    'EDIT MODE: propose.',
    'When the user asks for a change, call the appropriate edit tool. Edits are',
    'staged for the user to review and approve — they are NOT applied yet, so do',
    'not claim you have changed the document. Say what you are proposing.',
  ].join(' '),
  direct: [
    'EDIT MODE: direct.',
    'Edit tools apply immediately to the live document. The user can undo with',
    'Ctrl+Z. Make the change the user asked for and nothing more — do not',
    'reformat, retitle, or "improve" material you were not asked about.',
  ].join(' '),
  agentic: [
    'EDIT MODE: agentic.',
    'You may chain tool calls to complete a multi-step task. Read before you',
    'write. Work in small verifiable steps and stop as soon as the task is done.',
    'The whole run can be reverted as one unit, so prefer a coherent sequence',
    'over many unrelated edits.',
  ].join(' '),
}

const BASE_IDENTITY = [
  'You are OpenOfficeLLM, an AI assistant embedded in the application the user is working in.',
  'You help them read, understand, and where the host allows it, edit what is in front of them.',
].join(' ')

/** Identity for the browser side panel.
 *
 *  Replaces the base identity rather than adding to it: told it is in an Office
 *  task pane *and* looking at a web page, a model reliably offers to edit the
 *  page, which it cannot do. */
const BROWSER_IDENTITY = [
  'You are OpenOfficeLLM, an AI assistant in a browser side panel.',
  'You help the user read and understand the web page they are looking at.',
  "You can read the page but cannot change it: it is not the user's document.",
  'Never offer to edit, fix, or rewrite the page itself — offer the text instead.',
].join(' ')

const ANSWER_STYLE = [
  'Answer directly. Do not narrate your reasoning, restate the question, list',
  'the steps you plan to take, or describe your own response before giving it.',
  'The user sees only your final answer, so a preamble reads as noise.',
  'Use markdown for structure when it helps; keep prose tight.',
].join(' ')

/** The two things about the Word tools that are not obvious from their own
 *  descriptions, and that cost a wasted edit each when a model gets them wrong:
 *  how to reach text inside a paragraph, and that indices move. */
const WORD_EDITING_NOTES = [
  'You can change anything about this document that the user could: text, fonts,',
  'colours, highlighting, styles, alignment, spacing, indents, lists, tables,',
  'links, breaks, headers and footers, and page setup. To format a phrase inside a',
  'paragraph rather than the whole paragraph, target it with `find`. Paragraph',
  'numbers shift whenever you insert or delete a paragraph, so re-read the document',
  'before making further edits by number.',
].join(' ')

/** Word truncates on characters rather than tokens deliberately: a token
 *  estimate that is wrong in the unsafe direction costs the user a failed
 *  request, and character counts are exact. ~4 chars/token is the usual ratio. */
const MAX_CONTEXT_CHARS = 48_000

function truncate(text: string, limit = MAX_CONTEXT_CHARS): { text: string; truncated: boolean } {
  if (text.length <= limit) return { text, truncated: false }
  return { text: `${text.slice(0, limit)}\n…[truncated]`, truncated: true }
}

/** Flatten a message's content to a plain string, regardless of whether it
 *  arrived as a string or as typed content blocks. Image blocks contribute
 *  nothing to flat text (their content is the image bytes, which have no
 *  text representation here) — only text blocks are joined. Used where the
 *  system-prompt assembly needs a string: inheriting caller system messages,
 *  and anywhere else that historically assumed `content: string`. */
function textOf(content: string | ContentBlock[]): string {
  if (typeof content === 'string') return content
  return content
    .filter((b): b is Extract<ContentBlock, { type: 'text' }> => b.type === 'text')
    .map((b) => b.text)
    .join('')
}

function describeScope(ctx: DocumentContext): string {
  switch (ctx.scope) {
    case 'selection':
      return 'the text the user currently has selected'
    case 'paragraph':
      return 'the paragraph containing the cursor, with surrounding context'
    case 'document':
      return 'the full document body'
    case 'sheet':
      return 'the active worksheet'
    case 'range':
      return 'the currently selected cell range'
    case 'page':
      return 'the readable content of the page the user is viewing'
    default:
      return 'the current document context'
  }
}

/**
 * Render the attached document context as a system-prompt section.
 *
 * Returns null when there is nothing useful to say, so we never emit an empty
 * "here is the document:" header — a model shown an empty section reliably
 * concludes the document is empty rather than that none was attached.
 */
export function renderContextSection(ctx: DocumentContext | undefined): string | null {
  if (!ctx) return null
  const hasText = typeof ctx.text === 'string' && ctx.text.trim().length > 0
  const hasOutline = typeof ctx.outline === 'string' && ctx.outline.trim().length > 0
  const hasSchema = Array.isArray(ctx.schema) && ctx.schema.length > 0
  if (!hasText && !hasOutline && !hasSchema) return null

  const parts: string[] = []
  const appName = ctx.host === 'excel' ? 'Excel' : 'Word'
  parts.push(
    ctx.isAttachment && ctx.fileName
      ? `The user attached a file "${ctx.fileName}". The following is its contents, extracted just now.`
      : ctx.host === 'browser'
        ? `The user is reading a web page. The following is ${describeScope(ctx)}, captured just now.`
        : `The user is working in ${appName}. The following is ${describeScope(ctx)}, captured just now.`,
  )

  if (hasOutline) {
    const { text } = truncate(ctx.outline!, 8_000)
    const outlineHeading =
      ctx.host === 'excel'
        ? 'Workbook structure'
        : ctx.host === 'browser'
          ? 'Page outline'
          : 'Document outline'
    parts.push(`## ${outlineHeading}\n${text}`)
  }

  if (hasSchema) {
    const cols = ctx
      .schema!.map((c) => {
        const samples = c.sampleValues.slice(0, 3).join(', ')
        return `- ${c.name} (${c.type})${samples ? ` e.g. ${samples}` : ''}`
      })
      .join('\n')
    const rows = typeof ctx.totalRows === 'number' ? `\nTotal rows: ${ctx.totalRows}` : ''
    parts.push(`## Columns\n${cols}${rows}`)
  }

  if (hasText) {
    const { text, truncated } = truncate(ctx.text)
    const fence = ctx.host === 'excel' ? 'csv' : 'text'
    parts.push(`## Content\n\`\`\`${fence}\n${text}\n\`\`\``)
    if (truncated) {
      parts.push(
        ctx.host === 'browser'
          ? 'This content was truncated because the page is large. Use search_page to find specific parts if you need more.'
          : 'This content was truncated because the document is large. Use the read tools to fetch specific parts if you need more.',
      )
    }
  }

  if (ctx.host === 'excel' && typeof ctx.totalRows === 'number' && ctx.totalRows > 0) {
    parts.push(
      'Rows shown may be a sample of a larger range. Never assume the sample is the whole sheet when computing totals — read the range you need, or write a formula and let Excel compute it.',
    )
  }

  return parts.join('\n\n')
}

export interface BuildPromptOptions {
  req: ChatRequest
  /** Resolved skill prompt, if the turn invoked one. */
  skillPrompt?: string
  /** True if the turn was given any tools at all. */
  hasTools: boolean
  /** Rendered text from file attachments, appended to the system prompt
   *  after the skill section and before the document context section. */
  attachmentSection?: string | null
}

/**
 * Build the system prompt text for a turn. Order matters: identity, then how to
 * answer, then mode rules, then the skill, then the document. The document goes
 * last because it is the longest section and models weight the tail of a long
 * system prompt more reliably than its middle.
 */
export function buildSystemPrompt(opts: BuildPromptOptions): string {
  const { req, skillPrompt, hasTools, attachmentSection } = opts
  const browser = req.context?.host === 'browser'
  const sections: string[] = [browser ? BROWSER_IDENTITY : BASE_IDENTITY, ANSWER_STYLE]

  if (hasTools) {
    if (browser) {
      // No mode guidance and no editing vocabulary: this host has no write
      // tools, and telling a model how to propose an edit it cannot make is how
      // you get an assistant that announces changes nobody asked for and nobody
      // received.
      sections.push(
        [
          'You have tools for reading the page. Prefer reading it over asking the',
          'user to paste it. Never claim you cannot see the page — call read_page',
          'instead. On a long page, search_page is much cheaper than reading the',
          'whole thing.',
        ].join(' '),
      )
    } else {
      sections.push(
        [
          'You have tools for reading and editing the document. Prefer reading the',
          'document over asking the user to paste it. Never claim you cannot see the',
          'document — call a read tool instead. Only call an edit tool when the user',
          'has actually asked for a change.',
        ].join(' '),
      )
      sections.push(MODE_GUIDANCE[req.mode] ?? MODE_GUIDANCE.propose)
      if (req.context?.host === 'word') sections.push(WORD_EDITING_NOTES)
    }
  }

  if (skillPrompt && skillPrompt.trim()) sections.push(skillPrompt.trim())
  if (req.systemPrompt && req.systemPrompt.trim()) sections.push(req.systemPrompt.trim())
  if (attachmentSection) sections.push(attachmentSection)

  const ctxSection = renderContextSection(req.context)
  if (ctxSection) sections.push(ctxSection)

  return sections.join('\n\n')
}

/**
 * Return the message list to send upstream: the built system prompt followed by
 * the conversation, with any caller-supplied system messages folded in.
 *
 * Reasoning is stripped from assistant history on the way out. Feeding a
 * model's own scratchpad back to it degrades the next turn and burns tokens,
 * and several providers reject a `thinking` block that is not accompanied by
 * its original signature.
 */
export function withSystemPrompt(req: ChatRequest, systemPrompt: string): ChatMessage[] {
  const inherited = req.messages
    .filter((m) => m.role === 'system')
    .map((m) => textOf(m.content))
    .filter((c) => c.trim().length > 0)

  const combined = [systemPrompt, ...inherited].filter(Boolean).join('\n\n')
  const rest = req.messages
    .filter((m) => m.role !== 'system')
    .map(({ reasoning: _reasoning, ...m }) => m)

  return combined ? [{ role: 'system', content: combined }, ...rest] : rest
}

export interface AttachmentPayload {
  /** Text from text attachments and OCR'd images, rendered as a system-prompt
   *  section. Null when there are no text-contributing attachments. */
  systemSection: string | null
  /** Image content blocks for vision-capable models. Empty when the model
   *  can't take images, or when there are no image attachments. */
  imageBlocks: ContentBlock[]
}

interface BuildAttachmentOptions {
  req: ChatRequest
  /** True if the resolved model can accept image content blocks. When false,
   *  image attachments are OCR'd into text instead. */
  visionCapable: boolean
}

/**
 * Resolve the attachments on a ChatRequest into prompt material.
 *
 * Text attachments are read from the host's attachment store and rendered into
 * a single system-prompt section. Image attachments are either returned as
 * content blocks (for vision-capable models) or run through OCR and folded
 * into the same system-prompt section as text attachments.
 *
 * Called by the chat route after `buildSystemPrompt` so the attachment section
 * can be appended to the system prompt, and before `withSystemPrompt` so
 * image blocks can be injected into the first user message.
 */
export async function buildAttachmentPayload(opts: BuildAttachmentOptions): Promise<AttachmentPayload> {
  const { req, visionCapable } = opts
  const refs = req.attachments ?? []
  if (refs.length === 0) return { systemSection: null, imageBlocks: [] }

  const textSections: string[] = []
  const imageBlocks: ContentBlock[] = []

  for (const ref of refs) {
    const meta = getMeta(ref.id)
    if (!meta) {
      // The attachment was deleted or the host restarted. Skip it rather than
      // failing the whole turn — a missing attachment is not a reason to
      // refuse to answer.
      logger.warn({ msg: 'attachment not found', id: ref.id, fileName: ref.fileName })
      continue
    }

    if (meta.kind === 'image') {
      // Use the host's classified kind, not the client's ref.kind — the host
      // inspected the actual file bytes at upload time, while ref.kind is just
      // a label the pane forwarded. A stale or mislabeled client ref would
      // otherwise route an image through the text branch and drop it (images
      // have no meta.text), or push text through the vision branch.
      if (visionCapable) {
        const bytes = getBytes(ref.id)
        if (bytes) {
          imageBlocks.push({
            type: 'image',
            mimeType: meta.mimeType,
            data: bytes.buffer.toString('base64'),
          })
        }
      } else {
        // Non-vision model: OCR the image and fold the text into the system
        // prompt. This is the "OCR skill" path — the user gets text either way.
        const bytes = getBytes(ref.id)
        if (bytes) {
          try {
            const ocrText = await extractTextOcr(bytes.buffer, meta.mimeType)
            if (ocrText.trim()) {
              textSections.push(renderAttachmentFileSection(ref.fileName, ocrText))
            } else {
              textSections.push(renderAttachmentFileSection(ref.fileName, '(OCR found no text in this image.)'))
            }
          } catch (e) {
            logger.warn({ msg: 'ocr failed', id: ref.id, error: String((e as Error).message ?? e) })
            textSections.push(renderAttachmentFileSection(ref.fileName, '(OCR failed on this image.)'))
          }
        }
      }
    } else {
      // Text attachment: use the extracted text stored on the meta.
      if (meta.text) {
        textSections.push(renderAttachmentFileSection(ref.fileName, meta.text))
      }
    }
  }

  const systemSection = textSections.length > 0 ? textSections.join('\n\n') : null
  return { systemSection, imageBlocks }
}

/** Render a single attachment's text as a labeled section. */
function renderAttachmentFileSection(fileName: string, text: string): string {
  const { text: truncated, truncated: didTruncate } = truncate(text)
  const header = `## Attached file: ${fileName}${didTruncate ? ' (truncated)' : ''}`
  return `${header}\n\`\`\`text\n${truncated}\n\`\`\``
}
