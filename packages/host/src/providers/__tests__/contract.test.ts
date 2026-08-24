import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import os from 'node:os'
import { OpenAiCompatibleAdapter } from '../openai-compatible.js'
import { OllamaAdapter } from '../ollama.js'
import { AnthropicAdapter } from '../anthropic.js'
import { GoogleAdapter } from '../google.js'
import { clearProviders } from '../registry.js'
import { setSecret } from '../../secrets.js'
import type { StreamEvent } from '@openofficellm/shared'
import type { ProviderAdapter } from '../types.js'

const here = path.dirname(fileURLToPath(import.meta.url))
const FIXTURES = path.join(here, 'fixtures')
const TMP_APPDATA = path.join(os.tmpdir(), `ool-contract-${process.pid}-${Date.now()}`)

function readFixture(name: string): string {
  return fs.readFileSync(path.join(FIXTURES, name), 'utf8')
}

function fixtureStream(name: string): ReadableStream<Uint8Array> {
  const text = readFixture(name)
  const encoder = new TextEncoder()
  return new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(encoder.encode(text))
      controller.close()
    },
  })
}

function mockFetch(body: ReadableStream<Uint8Array> | string, status = 200): typeof fetch {
  return (async () =>
    new Response(typeof body === 'string' ? body : body, {
      status,
      headers: { 'content-type': status === 200 ? 'text/event-stream' : 'application/json' },
    })) as unknown as typeof fetch
}

async function collectStream(
  adapter: ProviderAdapter,
  req: Parameters<ProviderAdapter['stream']>[0],
): Promise<StreamEvent[]> {
  const ctrl = new AbortController()
  const out: StreamEvent[] = []
  for await (const ev of adapter.stream(req, ctrl.signal)) {
    out.push(ev)
    if (ev.type === 'done' || ev.type === 'error') break
  }
  return out
}

const baseReq = {
  messages: [{ role: 'user' as const, content: 'hi' }],
  model: '',
  mode: 'propose' as const,
}

beforeEach(() => {
  clearProviders()
  fs.mkdirSync(TMP_APPDATA, { recursive: true })
  process.env.APPDATA = TMP_APPDATA
})

afterEach(() => {
  try {
    fs.rmSync(TMP_APPDATA, { recursive: true, force: true })
  } catch {
    // ignore
  }
})

describe('contract: openai-compatible streaming deltas', () => {
  it('emits delta events and a usage event', async () => {
    const original = globalThis.fetch
    globalThis.fetch = mockFetch(fixtureStream('openai-streaming-deltas.txt'))
    try {
      const adapter = new OpenAiCompatibleAdapter({
        id: 'test-oai',
        name: 'Test OAI',
        baseUrl: 'http://127.0.0.1:9999/v1',
        kind: 'local',
        authHeaderStyle: 'none',
      })
      const events = await collectStream(adapter, { ...baseReq, model: 'test-oai/test' })
      const deltas = events.filter((e) => e.type === 'delta')
      const usage = events.find((e) => e.type === 'usage')
      expect(deltas.map((d) => (d as { text: string }).text).join('')).toBe('Hello world')
      expect(usage).toBeDefined()
      expect((usage as { promptTokens: number }).promptTokens).toBe(5)
    } finally {
      globalThis.fetch = original
    }
  })
})

describe('contract: openai-compatible tool call', () => {
  it('accumulates a tool_call event across chunks', async () => {
    const original = globalThis.fetch
    globalThis.fetch = mockFetch(fixtureStream('openai-tool-call.txt'))
    try {
      const adapter = new OpenAiCompatibleAdapter({
        id: 'test-oai',
        name: 'Test OAI',
        baseUrl: 'http://127.0.0.1:9999/v1',
        kind: 'local',
        authHeaderStyle: 'none',
      })
      const events = await collectStream(adapter, { ...baseReq, model: 'test-oai/test' })
      const toolCalls = events.filter((e) => e.type === 'tool_call')
      expect(toolCalls).toHaveLength(1)
      const tc = (toolCalls[0] as { toolCall: { name: string; arguments: string } }).toolCall
      expect(tc.name).toBe('search_document')
      expect(JSON.parse(tc.arguments)).toEqual({ q: 'foo' })
    } finally {
      globalThis.fetch = original
    }
  })
})

describe('contract: auth failure (401)', () => {
  it('emits an auth error event', async () => {
    setSecret('test-oai', 'sk-test-key')
    const original = globalThis.fetch
    globalThis.fetch = mockFetch('{"error":"unauthorized"}', 401)
    try {
      const adapter = new OpenAiCompatibleAdapter({
        id: 'test-oai',
        name: 'Test OAI',
        baseUrl: 'http://127.0.0.1:9999/v1',
        kind: 'cloud',
      })
      const events = await collectStream(adapter, { ...baseReq, model: 'test-oai/test' })
      const err = events.find((e) => e.type === 'error')
      expect(err).toBeDefined()
      expect((err as { code: string }).code).toBe('auth')
    } finally {
      globalThis.fetch = original
    }
  })
})

describe('contract: rate limit (429)', () => {
  it('emits a rate_limit error event', async () => {
    setSecret('test-oai', 'sk-test-key')
    const original = globalThis.fetch
    globalThis.fetch = mockFetch('{"error":"rate limited"}', 429)
    try {
      const adapter = new OpenAiCompatibleAdapter({
        id: 'test-oai',
        name: 'Test OAI',
        baseUrl: 'http://127.0.0.1:9999/v1',
        kind: 'cloud',
      })
      const events = await collectStream(adapter, { ...baseReq, model: 'test-oai/test' })
      const err = events.find((e) => e.type === 'error')
      expect(err).toBeDefined()
      expect((err as { code: string }).code).toBe('rate_limit')
      expect((err as { retryable: boolean }).retryable).toBe(true)
    } finally {
      globalThis.fetch = original
    }
  })
})

describe('contract: mid-stream abort', () => {
  it('stops emitting after abort', async () => {
    const original = globalThis.fetch
    globalThis.fetch = mockFetch(fixtureStream('openai-streaming-deltas.txt'))
    try {
      const adapter = new OpenAiCompatibleAdapter({
        id: 'test-oai',
        name: 'Test OAI',
        baseUrl: 'http://127.0.0.1:9999/v1',
        kind: 'local',
        authHeaderStyle: 'none',
      })
      const ctrl = new AbortController()
      const out: StreamEvent[] = []
      let count = 0
      for await (const ev of adapter.stream({ ...baseReq, model: 'test-oai/test' }, ctrl.signal)) {
        out.push(ev)
        count += 1
        if (count === 1) ctrl.abort()
      }
      const deltas = out.filter((e) => e.type === 'delta')
      expect(deltas.length).toBeLessThanOrEqual(2)
    } finally {
      globalThis.fetch = original
    }
  })
})

describe('contract: malformed chunk', () => {
  it('skips invalid JSON without throwing', async () => {
    const original = globalThis.fetch
    const encoder = new TextEncoder()
    const body = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(
          encoder.encode(
            'data: not-json\n\ndata: {"choices":[{"delta":{"content":"ok"}}]}\n\ndata: [DONE]\n\n',
          ),
        )
        controller.close()
      },
    })
    globalThis.fetch = mockFetch(body)
    try {
      const adapter = new OpenAiCompatibleAdapter({
        id: 'test-oai',
        name: 'Test OAI',
        baseUrl: 'http://127.0.0.1:9999/v1',
        kind: 'local',
        authHeaderStyle: 'none',
      })
      const events = await collectStream(adapter, { ...baseReq, model: 'test-oai/test' })
      const deltas = events.filter((e) => e.type === 'delta')
      expect(deltas.map((d) => (d as { text: string }).text).join('')).toBe('ok')
    } finally {
      globalThis.fetch = original
    }
  })
})

describe('contract: ollama streaming', () => {
  it('parses NDJSON deltas and usage', async () => {
    const original = globalThis.fetch
    globalThis.fetch = mockFetch(fixtureStream('ollama-streaming.txt'))
    try {
      const adapter = new OllamaAdapter()
      const events = await collectStream(adapter, { ...baseReq, model: 'ollama/test' })
      const deltas = events.filter((e) => e.type === 'delta')
      expect(deltas.map((d) => (d as { text: string }).text).join('')).toBe('Hello from Ollama')
      const usage = events.find((e) => e.type === 'usage')
      expect(usage).toBeDefined()
      expect((usage as { promptTokens: number }).promptTokens).toBe(5)
    } finally {
      globalThis.fetch = original
    }
  })
})

describe('contract: anthropic streaming', () => {
  it('parses typed SSE events', async () => {
    setSecret('anthropic', 'sk-ant-test')
    const original = globalThis.fetch
    globalThis.fetch = mockFetch(fixtureStream('anthropic-streaming.txt'))
    try {
      const adapter = new AnthropicAdapter()
      const events = await collectStream(adapter, { ...baseReq, model: 'anthropic/test' })
      const deltas = events.filter((e) => e.type === 'delta')
      expect(deltas.map((d) => (d as { text: string }).text).join('')).toBe('Hello from Claude')
      const usage = events.find((e) => e.type === 'usage')
      expect(usage).toBeDefined()
    } finally {
      globalThis.fetch = original
    }
  })
})

describe('contract: google adapter constructs', () => {
  it('can be instantiated', () => {
    const adapter = new GoogleAdapter()
    expect(adapter.id).toBe('google')
    expect(adapter.kind).toBe('cloud')
  })
})
