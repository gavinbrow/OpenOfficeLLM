import { describe, it, expect } from 'vitest'
import { openChatStream } from '../stream'

// Pure-logic test: we can't easily exercise fetch + ReadableStream in jsdom
// without a heavy harness, so we verify the public surface and the abort path.

describe('openChatStream', () => {
  it('returns an async iterable and an abort function', () => {
    const handle = openChatStream({ model: 'x', messages: [], mode: 'propose' })
    expect(typeof handle.events[Symbol.asyncIterator]).toBe('function')
    expect(typeof handle.abort).toBe('function')
  })

  it('abort() is safe to call without consuming the iterable', async () => {
    const handle = openChatStream({ model: 'x', messages: [], mode: 'propose' })
    expect(() => handle.abort()).not.toThrow()
    // Consuming after abort should terminate without throwing.
    const events: unknown[] = []
    for await (const ev of handle.events) {
      events.push(ev)
    }
    expect(events).toHaveLength(0)
  })

  it('respects an already-aborted signal', async () => {
    const controller = new AbortController()
    controller.abort()
    const handle = openChatStream(
      { model: 'x', messages: [], mode: 'propose' },
      { signal: controller.signal },
    )
    const events: unknown[] = []
    for await (const ev of handle.events) {
      events.push(ev)
    }
    expect(events).toHaveLength(0)
  })
})
