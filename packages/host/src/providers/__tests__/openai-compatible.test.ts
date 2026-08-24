import { describe, it, expect, vi, afterEach } from 'vitest'
import type { ChatRequest, StreamEvent, ToolCall } from '@openofficellm/shared'
import { OpenAiCompatibleAdapter } from '../openai-compatible.js'

/** An SSE response body built from raw `data:` frames. */
function sseResponse(frames: string[]): Response {
  const body = new ReadableStream<Uint8Array>({
    start(controller) {
      const enc = new TextEncoder()
      for (const f of frames) controller.enqueue(enc.encode(`data: ${f}\n\n`))
      controller.enqueue(enc.encode('data: [DONE]\n\n'))
      controller.close()
    },
  })
  return new Response(body, { status: 200, headers: { 'content-type': 'text/event-stream' } })
}

function chunk(delta: unknown, finish?: string): string {
  return JSON.stringify({ choices: [{ delta, ...(finish ? { finish_reason: finish } : {}) }] })
}

const REQ: ChatRequest = {
  model: 'lms/qwen',
  messages: [{ role: 'user', content: 'hi' }],
  mode: 'propose',
}

function adapter() {
  return new OpenAiCompatibleAdapter({
    id: 'lms',
    name: 'LMS',
    baseUrl: 'http://127.0.0.1:1234/v1',
    kind: 'local',
    authHeaderStyle: 'none',
  })
}

async function collect(res: Response, req: ChatRequest = REQ): Promise<StreamEvent[]> {
  vi.stubGlobal('fetch', vi.fn().mockResolvedValue(res))
  const out: StreamEvent[] = []
  for await (const ev of adapter().stream(req, new AbortController().signal)) out.push(ev)
  return out
}

function toolCalls(events: StreamEvent[]): ToolCall[] {
  return events.flatMap((e) => (e.type === 'tool_call' ? [e.toolCall] : []))
}

afterEach(() => vi.unstubAllGlobals())

describe('OpenAiCompatibleAdapter tool-call buffering', () => {
  it('does not duplicate the first argument fragment of a single-chunk tool call', async () => {
    // Ollama and LM Studio both deliver small tool calls whole, in one delta.
    // Double-counting that fragment yields `{...}{...}`, which the next request
    // sends straight back upstream — Ollama rejects it with HTTP 400
    // "invalid tool call arguments".
    const events = await collect(
      sseResponse([
        chunk({
          tool_calls: [
            {
              index: 0,
              id: 'call_1',
              function: { name: 'search_document', arguments: '{"query":"budget"}' },
            },
          ],
        }),
        chunk({}, 'tool_calls'),
      ]),
    )
    const calls = toolCalls(events)
    expect(calls).toHaveLength(1)
    expect(calls[0].arguments).toBe('{"query":"budget"}')
    expect(() => JSON.parse(calls[0].arguments)).not.toThrow()
  })

  it('concatenates fragmented arguments exactly once each', async () => {
    const events = await collect(
      sseResponse([
        chunk({
          tool_calls: [
            { index: 0, id: 'c1', function: { name: 'replace_selection', arguments: '{"te' } },
          ],
        }),
        chunk({ tool_calls: [{ index: 0, function: { arguments: 'xt":"h' } }] }),
        chunk({ tool_calls: [{ index: 0, function: { arguments: 'i"}' } }] }),
        chunk({}, 'tool_calls'),
      ]),
    )
    const calls = toolCalls(events)
    expect(calls).toHaveLength(1)
    expect(calls[0].arguments).toBe('{"text":"hi"}')
  })

  it('keeps parallel tool calls apart when the provider omits `index`', async () => {
    // Without an index every fragment lands in the same buffer, so two calls
    // merge into one with `{"a":1}{"b":2}` as its arguments.
    const events = await collect(
      sseResponse([
        chunk({ tool_calls: [{ id: 'c1', function: { name: 'read_document', arguments: '{}' } }] }),
        chunk({
          tool_calls: [
            { id: 'c2', function: { name: 'search_document', arguments: '{"query":"x"}' } },
          ],
        }),
        chunk({}, 'tool_calls'),
      ] as string[]),
    )
    const calls = toolCalls(events)
    expect(calls.map((c) => c.name)).toEqual(['read_document', 'search_document'])
    expect(calls.map((c) => c.arguments)).toEqual(['{}', '{"query":"x"}'])
  })

  it('emits buffered tool calls when the stream ends without a finish_reason', async () => {
    const events = await collect(
      sseResponse([
        chunk({
          tool_calls: [
            { index: 0, id: 'c1', function: { name: 'read_document', arguments: '{}' } },
          ],
        }),
      ]),
    )
    expect(toolCalls(events).map((c) => c.name)).toEqual(['read_document'])
  })

  it('reports a model-not-found error instead of ending the stream silently', async () => {
    const events = await collect(
      new Response(JSON.stringify({ error: { message: 'model not found' } }), { status: 404 }),
    )
    expect(events).toHaveLength(1)
    expect(events[0]).toMatchObject({ type: 'error' })
  })
})

describe('OpenAiCompatibleAdapter request body', () => {
  async function bodySentFor(req: ChatRequest): Promise<Record<string, unknown>> {
    const spy = vi.fn().mockResolvedValue(sseResponse([]))
    vi.stubGlobal('fetch', spy)
    for await (const _ of adapter().stream(req, new AbortController().signal)) {
      void _
    }
    const init = spy.mock.calls[0][1] as RequestInit
    return JSON.parse(String(init.body)) as Record<string, unknown>
  }

  it('re-serializes assistant tool-call arguments as a JSON object string', async () => {
    // Whatever mangling reached the transcript, the wire format has to be a
    // string that `JSON.parse` yields an object for, or strict servers 400.
    const body = await bodySentFor({
      ...REQ,
      messages: [
        { role: 'user', content: 'hi' },
        {
          role: 'assistant',
          content: '',
          toolCalls: [{ id: 'c1', name: 'search_document', arguments: '{"q":1}{"q":1}' }],
        },
        { role: 'tool', content: 'none', toolCallId: 'c1', toolName: 'search_document' },
      ],
    })
    const messages = body.messages as Array<{
      tool_calls?: Array<{ function: { arguments: string } }>
    }>
    const args = messages[1].tool_calls![0].function.arguments
    expect(typeof args).toBe('string')
    expect(JSON.parse(args)).toEqual({})
  })

  it('passes through valid arguments untouched', async () => {
    const body = await bodySentFor({
      ...REQ,
      messages: [
        {
          role: 'assistant',
          content: '',
          toolCalls: [{ id: 'c1', name: 'search_document', arguments: '{"query":"budget"}' }],
        },
      ],
    })
    const messages = body.messages as Array<{
      tool_calls?: Array<{ function: { arguments: string } }>
    }>
    expect(JSON.parse(messages[0].tool_calls![0].function.arguments)).toEqual({
      query: 'budget',
    })
  })
})
