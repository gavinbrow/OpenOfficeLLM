import { describe, it, expect } from 'vitest'
import type { StreamEvent } from '@openofficellm/shared'
import { ReasoningSplitter, splitReasoning } from '../reasoning.js'

/** Feed chunks through a splitter and return the concatenated channels. */
function run(chunks: string[]): { content: string; reasoning: string } {
  const s = new ReasoningSplitter()
  let content = ''
  let reasoning = ''
  for (const c of chunks) {
    const out = s.push(c)
    content += out.content
    reasoning += out.reasoning
  }
  const tail = s.flush()
  return { content: content + tail.content, reasoning: reasoning + tail.reasoning }
}

async function collect(events: StreamEvent[]): Promise<StreamEvent[]> {
  async function* source() {
    for (const e of events) yield e
  }
  const out: StreamEvent[] = []
  for await (const e of splitReasoning(source())) out.push(e)
  return out
}

describe('ReasoningSplitter', () => {
  it('passes text through untouched when there are no tags', () => {
    expect(run(['Hello ', 'world'])).toEqual({ content: 'Hello world', reasoning: '' })
  })

  it('separates a complete think block', () => {
    const r = run(['<think>planning</think>The answer.'])
    expect(r.reasoning).toBe('planning')
    expect(r.content).toBe('The answer.')
  })

  // The case that actually bit: a tag straddling two SSE deltas. A naive
  // per-chunk regex sees neither half and emits the tag as prose.
  it('handles a tag split across chunk boundaries', () => {
    const r = run(['<thi', 'nk>secret</thi', 'nk>visible'])
    expect(r.reasoning).toBe('secret')
    expect(r.content).toBe('visible')
  })

  it('handles the tag split one character at a time', () => {
    const r = run('<think>abc</think>xyz'.split(''))
    expect(r.reasoning).toBe('abc')
    expect(r.content).toBe('xyz')
  })

  it('treats a closing tag with no opener as the end of reasoning', () => {
    // Models prefilled with an opening tag emit only the closer.
    const r = run(['reasoned aloud</think>the answer'])
    expect(r.reasoning).toBe('')
    expect(r.content).toBe('reasoned aloudthe answer')
  })

  it('does not stall on prose containing a less-than sign', () => {
    const s = new ReasoningSplitter()
    // `<` followed by something that cannot start a tag must emit immediately,
    // otherwise every comparison operator delays the stream by a chunk.
    expect(s.push('if a < b then').content).toBe('if a < b then')
  })

  it('withholds only a genuine partial tag', () => {
    const s = new ReasoningSplitter()
    const first = s.push('answer <thi')
    expect(first.content).toBe('answer ')
    const second = s.push('nk>hidden')
    expect(second.content).toBe('')
    expect(second.reasoning).toBe('hidden')
  })

  it('flushes an unterminated partial tag as literal text', () => {
    expect(run(['done <thin'])).toEqual({ content: 'done <thin', reasoning: '' })
  })

  it('recognises alternate tag names case-insensitively', () => {
    const r = run(['<Thinking>x</Thinking>y'])
    expect(r.reasoning).toBe('x')
    expect(r.content).toBe('y')
  })

  it('keeps reasoning across multiple blocks', () => {
    const r = run(['a<think>1</think>b<think>2</think>c'])
    expect(r.reasoning).toBe('12')
    expect(r.content).toBe('abc')
  })
})

describe('splitReasoning', () => {
  it('converts inline tags into reasoning events', async () => {
    const out = await collect([
      { type: 'delta', text: '<think>plan' },
      { type: 'delta', text: 'ning</think>Answer' },
    ])
    expect(out).toEqual([
      { type: 'reasoning', text: 'plan' },
      { type: 'reasoning', text: 'ning' },
      { type: 'delta', text: 'Answer' },
    ])
  })

  it('passes non-delta events through unchanged', async () => {
    const usage: StreamEvent = {
      type: 'usage',
      promptTokens: 1,
      completionTokens: 2,
      totalTokens: 3,
    }
    const out = await collect([usage])
    expect(out).toEqual([usage])
  })

  it('drains held-back text before a terminal done event', async () => {
    const out = await collect([
      { type: 'delta', text: 'tail <thin' },
      { type: 'done', requestId: 'r1' },
    ])
    // The partial tag was never completed, so it is literal text — and it must
    // arrive before `done`, since consumers stop reading at the terminal event.
    expect(out).toEqual([
      { type: 'delta', text: 'tail ' },
      { type: 'delta', text: '<thin' },
      { type: 'done', requestId: 'r1' },
    ])
  })

  it('does not emit empty events', async () => {
    const out = await collect([
      { type: 'delta', text: '<think>' },
      { type: 'done', requestId: 'r' },
    ])
    expect(out).toEqual([{ type: 'done', requestId: 'r' }])
  })
})
