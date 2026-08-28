import { describe, it, expect, vi, beforeEach } from 'vitest'
import type { AttachmentRef, ChatRequest, DocumentContext } from '@openofficellm/shared'

// buildAttachmentPayload reaches into the on-disk attachment store (getMeta /
// getBytes) and the OCR worker (extractTextOcr). Mock the whole barrel so the
// prompt tests never touch the filesystem or spawn tesseract. renderContextSection
// and buildSystemPrompt do not import this module, so the existing tests are
// unaffected by the mock.
vi.mock('../attachments/index.js', () => ({
  getMeta: vi.fn(),
  getBytes: vi.fn(),
  extractTextOcr: vi.fn(),
  saveAttachment: vi.fn(),
  deleteAttachment: vi.fn(),
  cleanup: vi.fn(),
}))

import {
  buildSystemPrompt,
  renderContextSection,
  withSystemPrompt,
  buildAttachmentPayload,
} from '../prompt.js'
import { getMeta, getBytes, extractTextOcr } from '../attachments/index.js'

function req(overrides: Partial<ChatRequest> = {}): ChatRequest {
  return {
    messages: [{ role: 'user', content: 'hi' }],
    model: 'ollama/test',
    mode: 'propose',
    ...overrides,
  }
}

function ref(overrides: Partial<AttachmentRef> = {}): AttachmentRef {
  return {
    id: 'att-1',
    fileName: 'notes.txt',
    kind: 'text',
    mimeType: 'text/plain',
    tokenEstimate: 10,
    ...overrides,
  }
}

describe('renderContextSection', () => {
  it('returns null when there is no context', () => {
    expect(renderContextSection(undefined)).toBeNull()
  })

  // An empty section is worse than none: a model shown "here is the document:"
  // followed by nothing concludes the document is empty.
  it('returns null when the context carries no content', () => {
    const ctx: DocumentContext = { host: 'word', scope: 'document', text: '   ' }
    expect(renderContextSection(ctx)).toBeNull()
  })

  it('includes the document text and names the host', () => {
    const ctx: DocumentContext = { host: 'word', scope: 'document', text: 'Hello world' }
    const out = renderContextSection(ctx)!
    expect(out).toContain('Word')
    expect(out).toContain('Hello world')
  })

  it('renders the Excel column schema and row count', () => {
    const ctx: DocumentContext = {
      host: 'excel',
      scope: 'sheet',
      text: 'a,b',
      schema: [{ name: 'Amount', type: 'number', sampleValues: ['1', '2'] }],
      totalRows: 5000,
    }
    const out = renderContextSection(ctx)!
    expect(out).toContain('Amount (number)')
    expect(out).toContain('Total rows: 5000')
    // Guarding against the model summing a 60-row sample and calling it the total.
    expect(out).toContain('may be a sample')
  })

  it('marks truncated content and says how to get more', () => {
    const ctx: DocumentContext = { host: 'word', scope: 'document', text: 'x'.repeat(60_000) }
    const out = renderContextSection(ctx)!
    expect(out).toContain('[truncated]')
    expect(out).toContain('read tools')
  })
})

describe('buildSystemPrompt', () => {
  it('omits tool guidance when the turn has no tools', () => {
    const out = buildSystemPrompt({ req: req(), hasTools: false })
    expect(out).not.toContain('EDIT MODE')
  })

  it('includes mode guidance matching the request', () => {
    expect(buildSystemPrompt({ req: req({ mode: 'direct' }), hasTools: true })).toContain(
      'EDIT MODE: direct',
    )
    expect(buildSystemPrompt({ req: req({ mode: 'agentic' }), hasTools: true })).toContain(
      'EDIT MODE: agentic',
    )
  })

  it('tells the model not to narrate its reasoning', () => {
    expect(buildSystemPrompt({ req: req(), hasTools: false })).toContain('Do not narrate')
  })

  it('places the document context last', () => {
    const out = buildSystemPrompt({
      req: req({ context: { host: 'word', scope: 'document', text: 'DOCBODY' } }),
      hasTools: true,
    })
    expect(out.indexOf('DOCBODY')).toBeGreaterThan(out.indexOf('EDIT MODE'))
  })

  it('appends the skill prompt', () => {
    const out = buildSystemPrompt({ req: req(), hasTools: false, skillPrompt: 'BE TERSE' })
    expect(out).toContain('BE TERSE')
  })
})

describe('withSystemPrompt', () => {
  it('prepends a single system message and keeps the rest in order', () => {
    const messages = withSystemPrompt(req(), 'SYS')
    expect(messages[0]).toEqual({ role: 'system', content: 'SYS' })
    expect(messages).toHaveLength(2)
  })

  it('folds caller-supplied system messages into the one system message', () => {
    const messages = withSystemPrompt(
      req({
        messages: [
          { role: 'system', content: 'EXTRA' },
          { role: 'user', content: 'hi' },
        ],
      }),
      'SYS',
    )
    expect(messages).toHaveLength(2)
    expect(messages[0].content).toBe('SYS\n\nEXTRA')
  })

  // Reasoning is a per-turn artifact. Sending it back degrades the next turn,
  // and some providers reject a thinking block without its signature.
  it('strips reasoning from assistant history', () => {
    const messages = withSystemPrompt(
      req({
        messages: [
          { role: 'assistant', content: 'answer', reasoning: 'scratchpad' },
          { role: 'user', content: 'again' },
        ],
      }),
      'SYS',
    )
    expect(messages.some((m) => 'reasoning' in m)).toBe(false)
    expect(messages[1].content).toBe('answer')
  })
})

// The browser host is read-only, and the prompt is where that has to land. A
// model told it is in an Office task pane will offer to edit the page, and an
// assistant that offers a change it cannot make is worse than one that admits
// the limit up front.
describe('browser host', () => {
  const pageCtx: DocumentContext = {
    host: 'browser',
    scope: 'page',
    text: 'The quick brown fox.',
    outline: '# Heading',
  }

  it('says the user is reading a web page, not working in Word', () => {
    const out = renderContextSection(pageCtx)!
    expect(out).toContain('reading a web page')
    expect(out).not.toContain('working in Word')
    expect(out).not.toContain('working in Excel')
  })

  it('labels the outline as a page outline', () => {
    expect(renderContextSection(pageCtx)!).toContain('## Page outline')
  })

  it('describes the page scope in its own terms', () => {
    expect(renderContextSection(pageCtx)!).toContain('page the user is viewing')
  })

  it('uses the browser identity instead of the Office one', () => {
    const out = buildSystemPrompt({ req: req({ context: pageCtx }), hasTools: true })
    expect(out).toContain('browser side panel')
    expect(out).toContain('cannot change it')
    expect(out).not.toContain('Microsoft Office task pane')
  })

  it('withholds edit-mode guidance, which has nothing to gate here', () => {
    for (const mode of ['propose', 'direct', 'agentic'] as const) {
      const out = buildSystemPrompt({ req: req({ mode, context: pageCtx }), hasTools: true })
      expect(out).not.toContain('EDIT MODE')
    }
  })

  it('withholds the Word formatting notes', () => {
    const out = buildSystemPrompt({ req: req({ context: pageCtx }), hasTools: true })
    expect(out).not.toContain('colours, highlighting, styles')
  })

  it('still tells Word to use its edit tools', () => {
    const wordCtx: DocumentContext = { host: 'word', scope: 'document', text: 'x' }
    const out = buildSystemPrompt({ req: req({ context: wordCtx }), hasTools: true })
    expect(out).toContain('EDIT MODE')
    expect(out).toContain('reading and editing the document')
  })

  it('points at search_page rather than the read tools when a page is truncated', () => {
    const long: DocumentContext = { host: 'browser', scope: 'page', text: 'x'.repeat(60_000) }
    expect(renderContextSection(long)!).toContain('search_page')
  })
})

// An attachment and a live document render with different headers: the model
// is told "the user attached a file named X" rather than "the user is working
// in Word", so it does not assume it can edit the attachment the way it can
// edit the open document.
describe('renderContextSection (attachment)', () => {
  it('says the user attached a file when isAttachment and fileName are set', () => {
    const ctx: DocumentContext = {
      host: 'word',
      scope: 'document',
      text: 'extracted body',
      fileName: 'report.pdf',
      isAttachment: true,
    }
    const out = renderContextSection(ctx)!
    expect(out).toContain('user attached a file')
    expect(out).toContain('report.pdf')
    expect(out).not.toContain('working in Word')
    expect(out).toContain('extracted body')
  })

  it('falls back to the host line when isAttachment is false or undefined', () => {
    const attachedFalse: DocumentContext = {
      host: 'word',
      scope: 'document',
      text: 'body',
      fileName: 'ignored.pdf',
      isAttachment: false,
    }
    expect(renderContextSection(attachedFalse)!).toContain('working in Word')
    expect(renderContextSection(attachedFalse)!).not.toContain('user attached a file')

    const noFlag: DocumentContext = { host: 'word', scope: 'document', text: 'body' }
    expect(renderContextSection(noFlag)!).toContain('working in Word')
  })

  it('falls back to the host line when isAttachment is true but fileName is missing', () => {
    const ctx: DocumentContext = { host: 'word', scope: 'document', text: 'body', isAttachment: true }
    expect(renderContextSection(ctx)!).toContain('working in Word')
  })
})

describe('buildAttachmentPayload', () => {
  beforeEach(() => {
    vi.clearAllMocks()
  })

  it('returns null section and empty image blocks when there are no attachments', async () => {
    const out = await buildAttachmentPayload({ req: req(), visionCapable: true })
    expect(out.systemSection).toBeNull()
    expect(out.imageBlocks).toEqual([])
    expect(getMeta).not.toHaveBeenCalled()
  })

  it('folds a text attachment into the system section', async () => {
    vi.mocked(getMeta).mockReturnValue({
      id: 'att-1',
      fileName: 'notes.txt',
      mimeType: 'text/plain',
      kind: 'text',
      tokenEstimate: 5,
      text: 'the extracted text',
      path: '/tmp/att-1/notes.txt',
    })
    const out = await buildAttachmentPayload({
      req: req({ attachments: [ref()] }),
      visionCapable: false,
    })
    expect(out.imageBlocks).toEqual([])
    expect(out.systemSection).toContain('notes.txt')
    expect(out.systemSection).toContain('the extracted text')
  })

  it('returns no image blocks for a text attachment even when vision is available', async () => {
    vi.mocked(getMeta).mockReturnValue({
      id: 'att-1',
      fileName: 'notes.txt',
      mimeType: 'text/plain',
      kind: 'text',
      tokenEstimate: 5,
      text: 'text',
      path: '/tmp/att-1/notes.txt',
    })
    const out = await buildAttachmentPayload({
      req: req({ attachments: [ref()] }),
      visionCapable: true,
    })
    expect(out.imageBlocks).toEqual([])
    expect(out.systemSection).toContain('text')
  })

  it('inlines an image attachment as a vision content block when visionCapable', async () => {
    vi.mocked(getMeta).mockReturnValue({
      id: 'att-img',
      fileName: 'photo.png',
      mimeType: 'image/png',
      kind: 'image',
      tokenEstimate: 0,
      path: '/tmp/att-img/photo.png',
    })
    vi.mocked(getBytes).mockReturnValue({
      buffer: Buffer.from('binary', 'utf-8'),
      mimeType: 'image/png',
    })
    const out = await buildAttachmentPayload({
      req: req({ attachments: [ref({ id: 'att-img', fileName: 'photo.png', kind: 'image', mimeType: 'image/png' })] }),
      visionCapable: true,
    })
    expect(out.imageBlocks).toHaveLength(1)
    const block = out.imageBlocks[0]
    expect(block.type).toBe('image')
    if (block.type !== 'image') throw new Error('expected image block')
    expect(block.mimeType).toBe('image/png')
    expect(block.data).toBe(Buffer.from('binary', 'utf-8').toString('base64'))
    expect(out.systemSection).toBeNull()
    expect(extractTextOcr).not.toHaveBeenCalled()
  })

  it('OCR-folds an image into the system section when the model is not vision-capable', async () => {
    vi.mocked(getMeta).mockReturnValue({
      id: 'att-img',
      fileName: 'scan.png',
      mimeType: 'image/png',
      kind: 'image',
      tokenEstimate: 0,
      path: '/tmp/att-img/scan.png',
    })
    vi.mocked(getBytes).mockReturnValue({ buffer: Buffer.from('px'), mimeType: 'image/png' })
    vi.mocked(extractTextOcr).mockResolvedValue('ocr result text')
    const out = await buildAttachmentPayload({
      req: req({ attachments: [ref({ id: 'att-img', fileName: 'scan.png', kind: 'image', mimeType: 'image/png' })] }),
      visionCapable: false,
    })
    expect(out.imageBlocks).toEqual([])
    expect(extractTextOcr).toHaveBeenCalledTimes(1)
    expect(out.systemSection).toContain('scan.png')
    expect(out.systemSection).toContain('ocr result text')
  })

  it('notes when OCR finds no text in an image', async () => {
    vi.mocked(getMeta).mockReturnValue({
      id: 'att-img',
      fileName: 'blank.png',
      mimeType: 'image/png',
      kind: 'image',
      tokenEstimate: 0,
      path: '/tmp/blank.png',
    })
    vi.mocked(getBytes).mockReturnValue({ buffer: Buffer.from('px'), mimeType: 'image/png' })
    vi.mocked(extractTextOcr).mockResolvedValue('')
    const out = await buildAttachmentPayload({
      req: req({ attachments: [ref({ id: 'att-img', fileName: 'blank.png', kind: 'image', mimeType: 'image/png' })] }),
      visionCapable: false,
    })
    expect(out.systemSection).toContain('no text')
  })

  it('records an OCR failure as a labeled section rather than throwing', async () => {
    vi.mocked(getMeta).mockReturnValue({
      id: 'att-img',
      fileName: 'broken.png',
      mimeType: 'image/png',
      kind: 'image',
      tokenEstimate: 0,
      path: '/tmp/broken.png',
    })
    vi.mocked(getBytes).mockReturnValue({ buffer: Buffer.from('px'), mimeType: 'image/png' })
    vi.mocked(extractTextOcr).mockRejectedValue(new Error('ocr crash'))
    const out = await buildAttachmentPayload({
      req: req({ attachments: [ref({ id: 'att-img', fileName: 'broken.png', kind: 'image', mimeType: 'image/png' })] }),
      visionCapable: false,
    })
    expect(out.systemSection).toContain('OCR failed')
  })

  it('skips a missing attachment (getMeta returns null) without throwing', async () => {
    vi.mocked(getMeta).mockReturnValue(null)
    const out = await buildAttachmentPayload({
      req: req({ attachments: [ref({ id: 'gone' })] }),
      visionCapable: true,
    })
    expect(out.systemSection).toBeNull()
    expect(out.imageBlocks).toEqual([])
    expect(getBytes).not.toHaveBeenCalled()
  })

  it('skips a text attachment whose meta has no extracted text', async () => {
    vi.mocked(getMeta).mockReturnValue({
      id: 'att-1',
      fileName: 'empty.txt',
      mimeType: 'text/plain',
      kind: 'text',
      tokenEstimate: 0,
      text: '',
      path: '/tmp/empty.txt',
    })
    const out = await buildAttachmentPayload({
      req: req({ attachments: [ref({ id: 'att-1', fileName: 'empty.txt' })] }),
      visionCapable: false,
    })
    expect(out.systemSection).toBeNull()
  })
})
