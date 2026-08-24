import { describe, it, expect } from 'vitest'
import type { ChatRequest, DocumentContext } from '@openofficellm/shared'
import { buildSystemPrompt, renderContextSection, withSystemPrompt } from '../prompt.js'

function req(overrides: Partial<ChatRequest> = {}): ChatRequest {
  return {
    messages: [{ role: 'user', content: 'hi' }],
    model: 'ollama/test',
    mode: 'propose',
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
