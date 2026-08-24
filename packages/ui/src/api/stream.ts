// SSE stream client over fetch + ReadableStream. EventSource cannot POST or
// set headers, so we parse the `text/event-stream` wire format manually.
//
// Handles: partial line buffering, `event:`/`data:` pairs separated by blank
// lines, mid-stream reconnect with capped backoff, and abort via AbortController.

import type { StreamEvent } from '@openofficellm/shared'
import { getAuthToken, apiUrl, baseHeaders } from './client'

export interface StreamOptions {
  signal?: AbortSignal
  /** Max reconnect attempts after a mid-stream drop. Default 2. */
  maxReconnects?: number
  /** Base backoff ms; doubles per attempt. Default 500. */
  baseBackoffMs?: number
}

export interface StreamHandle {
  /** Async iterable of parsed StreamEvents. */
  events: AsyncIterable<StreamEvent>
  /** Abort the in-flight request (wired to the Stop button). */
  abort: () => void
}

interface ParsedMessage {
  event: string | null
  data: string | null
}

function parseEventChunk(buffer: string): { messages: ParsedMessage[]; remainder: string } {
  const messages: ParsedMessage[] = []
  const lines = buffer.split('\n')
  // The final element is whatever follows the last newline — it may be a
  // partial line, so keep it as the remainder for the next chunk.
  const remainder = lines.pop() ?? ''
  let current: ParsedMessage = { event: null, data: null }
  let haveFields = false
  const flush = () => {
    if (haveFields) {
      messages.push(current)
      current = { event: null, data: null }
      haveFields = false
    }
  }
  for (const rawLine of lines) {
    const line = rawLine.replace(/\r$/, '')
    if (line === '') {
      flush()
      continue
    }
    if (line.startsWith(':')) continue
    const colon = line.indexOf(':')
    const field = colon === -1 ? line : line.slice(0, colon)
    let value = colon === -1 ? '' : line.slice(colon + 1)
    if (value.startsWith(' ')) value = value.slice(1)
    if (field === 'event') {
      current.event = value
      haveFields = true
    } else if (field === 'data') {
      current.data = current.data === null ? value : `${current.data}\n${value}`
      haveFields = true
    }
  }
  return { messages, remainder }
}

function toStreamEvent(msg: ParsedMessage): StreamEvent | null {
  const data = msg.data
  if (!data) return null
  try {
    const parsed = JSON.parse(data) as StreamEvent
    if (
      parsed &&
      typeof parsed === 'object' &&
      typeof (parsed as { type: unknown }).type === 'string'
    ) {
      return parsed
    }
  } catch {
    return null
  }
  return null
}

function sleep(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) {
      reject(new DOMException('Aborted', 'AbortError'))
      return
    }
    const t = setTimeout(() => {
      signal?.removeEventListener('abort', onAbort)
      resolve()
    }, ms)
    const onAbort = () => {
      clearTimeout(t)
      reject(new DOMException('Aborted', 'AbortError'))
    }
    signal?.addEventListener('abort', onAbort, { once: true })
  })
}

export interface StreamCallbacks {
  onReconnect?: (attempt: number, backoffMs: number) => void
  onReconnectFailed?: (attempts: number) => void
}

/**
 * Open an SSE chat stream. Returns a handle whose `events` is an async
 * iterable consumed by the chat store; call `abort()` from the Stop button.
 *
 * The body payload is the ChatRequest; we re-POST on reconnect. Because chat
 * is not idempotent, reconnect only retries if no events have yet been
 * delivered for the current turn — once deltas have streamed, a drop surfaces
 * as an error event instead of a retry.
 */
export function openChatStream(
  payload: unknown,
  opts: StreamOptions = {},
  callbacks: StreamCallbacks = {},
): StreamHandle {
  const controller = new AbortController()
  if (opts.signal) {
    if (opts.signal.aborted) controller.abort()
    else opts.signal.addEventListener('abort', () => controller.abort(), { once: true })
  }
  const maxReconnects = opts.maxReconnects ?? 2
  const baseBackoffMs = opts.baseBackoffMs ?? 500

  const token = getAuthToken()
  const headers: Record<string, string> = {
    ...baseHeaders(),
    'Content-Type': 'application/json',
    Accept: 'text/event-stream',
  }
  if (token) headers['Authorization'] = `Bearer ${token}`

  const body = JSON.stringify(payload)

  async function* events(): AsyncIterable<StreamEvent> {
    let attempt = 0
    let sawAnyEvent = false
    while (true) {
      if (controller.signal.aborted) return
      let res: Response
      try {
        res = await fetch(apiUrl('/api/chat'), {
          method: 'POST',
          headers,
          body,
          signal: controller.signal,
        })
      } catch (e) {
        if (controller.signal.aborted || (e as Error).name === 'AbortError') return
        if (attempt < maxReconnects && !sawAnyEvent) {
          const backoff = baseBackoffMs * 2 ** attempt
          callbacks.onReconnect?.(attempt + 1, backoff)
          try {
            await sleep(backoff, controller.signal)
          } catch {
            return
          }
          attempt++
          continue
        }
        callbacks.onReconnectFailed?.(attempt)
        yield { type: 'error', code: 'network', message: (e as Error).message, retryable: true }
        return
      }

      if (res.status === 401 || res.status === 403) {
        yield {
          type: 'error',
          code: 'forbidden',
          message: 'Authentication failed. Relaunch the add-in from the ribbon.',
        }
        return
      }
      if (!res.ok || !res.body) {
        let message = `HTTP ${res.status}`
        try {
          const errBody = (await res.json()) as { message?: string }
          if (errBody?.message) message = errBody.message
        } catch {
          // ignore
        }
        const retryable = res.status >= 500 || res.status === 429
        if (retryable && attempt < maxReconnects && !sawAnyEvent) {
          const backoff = baseBackoffMs * 2 ** attempt
          callbacks.onReconnect?.(attempt + 1, backoff)
          try {
            await sleep(backoff, controller.signal)
          } catch {
            return
          }
          attempt++
          continue
        }
        yield { type: 'error', code: `http_${res.status}`, message, retryable }
        return
      }

      const reader = res.body.getReader()
      const decoder = new TextDecoder('utf-8')
      let buffer = ''
      // `sawTerminalEvent` is true only if the host explicitly sent `done` or
      // `error`. A clean connection close (reader.read() returns done=true)
      // without a terminal event is a NORMAL end-of-stream for providers that
      // don't synthesize `done` (Ollama/OpenAI/Anthropic). Only flag a
      // stream_interrupted error if the close was due to an exception or if no
      // events were ever received (the request never really got started).
      let sawTerminalEvent = false
      let sawAnyEventThisAttempt = sawAnyEvent

      try {
        while (true) {
          if (controller.signal.aborted) {
            try {
              await reader.cancel()
            } catch {
              // ignore
            }
            return
          }
          const { done, value } = await reader.read()
          if (done) {
            const { messages } = parseEventChunk(buffer + '\n')
            for (const m of messages) {
              const ev = toStreamEvent(m)
              if (ev) {
                sawAnyEvent = true
                sawAnyEventThisAttempt = true
                if (ev.type === 'done' || ev.type === 'error') sawTerminalEvent = true
                yield ev
                if (ev.type === 'done' || ev.type === 'error') return
              }
            }
            buffer = ''
            break
          }
          buffer += decoder.decode(value, { stream: true })
          const parsed = parseEventChunk(buffer)
          buffer = parsed.remainder
          for (const m of parsed.messages) {
            const ev = toStreamEvent(m)
            if (ev) {
              sawAnyEvent = true
              sawAnyEventThisAttempt = true
              if (ev.type === 'done' || ev.type === 'error') sawTerminalEvent = true
              yield ev
              if (ev.type === 'done' || ev.type === 'error') return
            }
          }
        }
      } catch (e) {
        if (controller.signal.aborted || (e as Error).name === 'AbortError') {
          try {
            await reader.cancel()
          } catch {
            // ignore
          }
          return
        }
        // A network exception mid-stream is a genuine interruption. Only
        // surface it as an error if we already saw events (otherwise retry).
        if (attempt < maxReconnects && !sawAnyEventThisAttempt) {
          const backoff = baseBackoffMs * 2 ** attempt
          callbacks.onReconnect?.(attempt + 1, backoff)
          try {
            await sleep(backoff, controller.signal)
          } catch {
            return
          }
          attempt++
          continue
        }
        callbacks.onReconnectFailed?.(attempt)
        yield {
          type: 'error',
          code: 'stream_interrupted',
          message: (e as Error).message ?? 'stream dropped mid-flight',
          retryable: true,
        }
        return
      }

      // The connection closed cleanly (reader.read() returned done=true).
      // If the host sent an explicit `done`/`error`, we already returned.
      // If events were received and the connection closed without a terminal
      // event, treat it as a successful end-of-stream (the host now
      // synthesizes `done`, but older hosts may not — be tolerant).
      if (!sawTerminalEvent) {
        if (!sawAnyEventThisAttempt && attempt < maxReconnects) {
          // No events at all — the request never produced output. Retry.
          const backoff = baseBackoffMs * 2 ** attempt
          callbacks.onReconnect?.(attempt + 1, backoff)
          try {
            await sleep(backoff, controller.signal)
          } catch {
            return
          }
          attempt++
          continue
        }
        // Events were received and the stream closed cleanly — success.
      }
      return
    }
  }

  return {
    events: events(),
    abort: () => controller.abort(),
  }
}
