export interface ParsedSseEvent {
  event?: string
  data: string
  id?: string
  retry?: number
}

const DATA_DONE_SENTINEL = '[DONE]'

export async function* parseSseStream(
  body: ReadableStream<Uint8Array> | null,
): AsyncGenerator<ParsedSseEvent> {
  if (!body) return
  const reader = body.getReader()
  const decoder = new TextDecoder('utf-8')
  let buffer = ''
  try {
    while (true) {
      const { done, value } = await reader.read()
      if (done) break
      buffer += decoder.decode(value, { stream: true })
      const events = drainBuffer(buffer)
      buffer = events.remaining
      for (const ev of events.events) {
        yield ev
      }
    }
    const tail = decoder.decode()
    if (tail) buffer += tail
    if (buffer.trim()) {
      const ev = parseEventBlock(buffer)
      if (ev) yield ev
    }
  } finally {
    try {
      reader.releaseLock()
    } catch {
      // ignore
    }
  }
}

function drainBuffer(buffer: string): { events: ParsedSseEvent[]; remaining: string } {
  const events: ParsedSseEvent[] = []
  let remaining = buffer
  while (true) {
    const idx = findEventBoundary(remaining)
    if (idx < 0) break
    const block = remaining.slice(0, idx)
    remaining = remaining.slice(idx).replace(/^(\r?\n){2}/, '')
    const ev = parseEventBlock(block)
    if (ev) events.push(ev)
  }
  return { events, remaining }
}

function findEventBoundary(s: string): number {
  const lf = s.indexOf('\n\n')
  const crlf = s.indexOf('\r\n\r\n')
  if (lf === -1 && crlf === -1) return -1
  if (crlf === -1) return lf
  if (lf === -1) return crlf
  return Math.min(lf, crlf)
}

function parseEventBlock(block: string): ParsedSseEvent | null {
  const lines = block.split(/\r?\n/)
  const event: ParsedSseEvent = { data: '' }
  const dataParts: string[] = []
  let sawAny = false
  for (const rawLine of lines) {
    if (rawLine === '') continue
    if (rawLine.startsWith(':')) continue
    const colon = rawLine.indexOf(':')
    const field = colon === -1 ? rawLine : rawLine.slice(0, colon)
    let value = colon === -1 ? '' : rawLine.slice(colon + 1)
    if (value.startsWith(' ')) value = value.slice(1)
    sawAny = true
    switch (field) {
      case 'event':
        event.event = value
        break
      case 'data':
        dataParts.push(value)
        break
      case 'id':
        event.id = value
        break
      case 'retry':
        {
          const n = Number(value)
          if (Number.isFinite(n)) event.retry = n
        }
        break
      default:
        break
    }
  }
  if (!sawAny) return null
  event.data = dataParts.join('\n')
  return event
}

export function isDoneSentinel(data: string): boolean {
  return data.trim() === DATA_DONE_SENTINEL
}

export interface NdjsonEvent {
  obj: unknown
  raw: string
}

export async function* parseNdjsonStream(
  body: ReadableStream<Uint8Array> | null,
): AsyncGenerator<NdjsonEvent> {
  if (!body) return
  const reader = body.getReader()
  const decoder = new TextDecoder('utf-8')
  let buffer = ''
  try {
    while (true) {
      const { done, value } = await reader.read()
      if (done) break
      buffer += decoder.decode(value, { stream: true })
      const lines = drainLines(buffer)
      buffer = lines.remaining
      for (const line of lines.lines) {
        const trimmed = line.trim()
        if (!trimmed) continue
        try {
          const obj = JSON.parse(trimmed)
          yield { obj, raw: trimmed }
        } catch {
          // skip malformed line
        }
      }
    }
    const tail = decoder.decode()
    if (tail) buffer += tail
    if (buffer.trim()) {
      try {
        const obj = JSON.parse(buffer.trim())
        yield { obj, raw: buffer.trim() }
      } catch {
        // ignore trailing partial
      }
    }
  } finally {
    try {
      reader.releaseLock()
    } catch {
      // ignore
    }
  }
}

function drainLines(buffer: string): { lines: string[]; remaining: string } {
  const lines: string[] = []
  let remaining = buffer
  while (true) {
    const nl = remaining.indexOf('\n')
    if (nl === -1) break
    lines.push(remaining.slice(0, nl))
    remaining = remaining.slice(nl + 1)
  }
  return { lines, remaining }
}

export async function* rechunk(
  chunks: AsyncIterable<Uint8Array>,
  sizes: number[],
): AsyncGenerator<Uint8Array> {
  const queue: number[] = [...sizes]
  let pending: Uint8Array = new Uint8Array(0)
  for await (const chunk of chunks) {
    pending = concatU8(pending, chunk)
    while (queue.length > 0 && pending.length >= queue[0]) {
      const size = queue.shift()!
      yield pending.subarray(0, size)
      pending = pending.subarray(size)
    }
  }
  if (pending.length > 0) {
    yield pending
  }
}

function concatU8(a: Uint8Array, b: Uint8Array): Uint8Array {
  const out = new Uint8Array(a.length + b.length)
  out.set(a, 0)
  out.set(b, a.length)
  return out
}
