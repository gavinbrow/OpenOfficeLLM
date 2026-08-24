import { describe, it, expect } from 'vitest'
import { parseSseStream, parseNdjsonStream, isDoneSentinel, rechunk } from '../sse-parser.js'

function toStream(chunks: string[]): ReadableStream<Uint8Array> {
  const encoder = new TextEncoder()
  return new ReadableStream<Uint8Array>({
    start(controller) {
      for (const c of chunks) {
        controller.enqueue(encoder.encode(c))
      }
      controller.close()
    },
  })
}

function toBytes(chunks: string[]): AsyncIterable<Uint8Array> {
  const encoder = new TextEncoder()
  return (async function* () {
    for (const c of chunks) {
      yield encoder.encode(c)
    }
  })()
}

describe('parseSseStream', () => {
  it('parses a single data event', async () => {
    const events = []
    for await (const ev of parseSseStream(toStream(['data: {"a":1}\n\n']))) {
      events.push(ev)
    }
    expect(events).toHaveLength(1)
    expect(events[0].data).toBe('{"a":1}')
  })

  it('parses multiple events', async () => {
    const src = 'data: {"a":1}\n\ndata: {"a":2}\n\n'
    const events = []
    for await (const ev of parseSseStream(toStream([src]))) {
      events.push(ev)
    }
    expect(events).toHaveLength(2)
    expect(events[0].data).toBe('{"a":1}')
    expect(events[1].data).toBe('{"a":2}')
  })

  it('handles event-typed lines (Anthropic)', async () => {
    const src = 'event: message_delta\ndata: {"usage":{"output_tokens":5}}\n\n'
    const events = []
    for await (const ev of parseSseStream(toStream([src]))) {
      events.push(ev)
    }
    expect(events).toHaveLength(1)
    expect(events[0].event).toBe('message_delta')
    expect(events[0].data).toBe('{"usage":{"output_tokens":5}}')
  })

  it('handles multi-line data fields', async () => {
    const src = 'data: line1\ndata: line2\n\n'
    const events = []
    for await (const ev of parseSseStream(toStream([src]))) {
      events.push(ev)
    }
    expect(events).toHaveLength(1)
    expect(events[0].data).toBe('line1\nline2')
  })

  it('handles [DONE] sentinel', async () => {
    expect(isDoneSentinel('[DONE]')).toBe(true)
    expect(isDoneSentinel(' [DONE] ')).toBe(true)
    expect(isDoneSentinel('{"a":1}')).toBe(false)
  })

  it('handles chunked transfer splitting a single event across chunks', async () => {
    const events = []
    for await (const ev of parseSseStream(toStream(['data: {"a"', ':1}\n\n']))) {
      events.push(ev)
    }
    expect(events).toHaveLength(1)
    expect(events[0].data).toBe('{"a":1}')
  })

  it('handles partial JSON lines split mid-stream', async () => {
    const events = []
    for await (const ev of parseSseStream(toStream(['data: {"a":1}\n\nda', 'ta: {"b":2}\n\n']))) {
      events.push(ev)
    }
    expect(events).toHaveLength(2)
    expect(events[0].data).toBe('{"a":1}')
    expect(events[1].data).toBe('{"b":2}')
  })

  it('handles CRLF line endings', async () => {
    const events = []
    for await (const ev of parseSseStream(toStream(['data: {"a":1}\r\n\r\n']))) {
      events.push(ev)
    }
    expect(events).toHaveLength(1)
    expect(events[0].data).toBe('{"a":1}')
  })

  it('ignores comment lines', async () => {
    const events = []
    for await (const ev of parseSseStream(toStream([': keepalive\ndata: {"a":1}\n\n']))) {
      events.push(ev)
    }
    expect(events).toHaveLength(1)
    expect(events[0].data).toBe('{"a":1}')
  })

  it('handles empty body', async () => {
    const events = []
    for await (const ev of parseSseStream(toStream([]))) {
      events.push(ev)
    }
    expect(events).toHaveLength(0)
  })

  it('handles null body', async () => {
    const events = []
    for await (const ev of parseSseStream(null)) {
      events.push(ev)
    }
    expect(events).toHaveLength(0)
  })

  it('fuzz: truncated and interleaved chunks', async () => {
    const full = 'data: {"a":1}\n\ndata: {"b":2}\n\ndata: {"c":3}\n\n'
    const chunkSizes = [3, 7, 1, 15, 4]
    const events = []
    for await (const ev of parseSseStream(toStream(sliceString(full, chunkSizes)))) {
      events.push(ev)
    }
    expect(events).toHaveLength(3)
    expect(events.map((e) => e.data)).toEqual(['{"a":1}', '{"b":2}', '{"c":3}'])
  })
})

describe('parseNdjsonStream', () => {
  it('parses one JSON object per line', async () => {
    const src = '{"a":1}\n{"b":2}\n'
    const events = []
    for await (const ev of parseNdjsonStream(toStream([src]))) {
      events.push(ev)
    }
    expect(events).toHaveLength(2)
    expect(events[0].obj).toEqual({ a: 1 })
    expect(events[1].obj).toEqual({ b: 2 })
  })

  it('handles a trailing object without newline', async () => {
    const src = '{"a":1}\n{"b":2}'
    const events = []
    for await (const ev of parseNdjsonStream(toStream([src]))) {
      events.push(ev)
    }
    expect(events).toHaveLength(2)
  })

  it('handles chunked splits within a line', async () => {
    const events = []
    for await (const ev of parseNdjsonStream(toStream(['{"a":', '1}\n{"b":2}\n']))) {
      events.push(ev)
    }
    expect(events).toHaveLength(2)
    expect(events[0].obj).toEqual({ a: 1 })
  })

  it('skips malformed lines', async () => {
    const src = 'not json\n{"a":1}\n'
    const events = []
    for await (const ev of parseNdjsonStream(toStream([src]))) {
      events.push(ev)
    }
    expect(events).toHaveLength(1)
    expect(events[0].obj).toEqual({ a: 1 })
  })

  it('fuzz: split across many small chunks', async () => {
    const full = '{"a":1}\n{"done":true}\n'
    const events = []
    for await (const ev of parseNdjsonStream(toStream(sliceString(full, [2, 5, 1, 8, 3])))) {
      events.push(ev)
    }
    expect(events).toHaveLength(2)
  })

  it('skips blank lines', async () => {
    const src = '{"a":1}\n\n{"b":2}\n'
    const events = []
    for await (const ev of parseNdjsonStream(toStream([src]))) {
      events.push(ev)
    }
    expect(events).toHaveLength(2)
  })
})

describe('rechunk', () => {
  it('rechunks a byte stream into the requested sizes', async () => {
    const out = []
    for await (const c of rechunk(toBytes(['hello world!!']), [3, 5, 100])) {
      out.push(Buffer.from(c).toString('utf8'))
    }
    expect(out).toEqual(['hel', 'lo wo', 'rld!!'])
  })
})

function sliceString(s: string, sizes: number[]): string[] {
  const chunks: string[] = []
  let i = 0
  for (const size of sizes) {
    chunks.push(s.slice(i, i + size))
    i += size
  }
  if (i < s.length) chunks.push(s.slice(i))
  return chunks.filter((c) => c.length > 0)
}
