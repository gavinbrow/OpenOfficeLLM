import type {
  ChatMessage,
  ChatRequest,
  ModelInfo,
  ProviderCapabilities,
  StreamEvent,
  ToolCall,
  ImportedProviderConfig,
} from '@openofficellm/shared'
import { providerApiKey } from './credentials.js'
import {
  AuthError,
  NetworkError,
  ProviderError,
  classifyHttpStatus,
  toStreamEvent,
  type ProviderAdapter,
} from './types.js'
import { logger } from '../logging.js'
import { fetchWithRetry } from './retry.js'
import { parseSseStream, isDoneSentinel } from './sse-parser.js'
import { estimateCost } from './pricing.js'

export interface OpenAiCompatibleConfig {
  id: string
  name: string
  baseUrl: string
  kind: 'local' | 'cloud'
  capabilities?: Partial<ProviderCapabilities>
  authHeaderStyle?: 'bearer' | 'api-key-header' | 'none'
  apiKeyHeader?: string
  extraHeaders?: Record<string, string>
  modelsHint?: string[]
  isConfiguredOverride?: () => boolean
  requiresKey?: boolean
  /** Optional API key supplied by an importer; used when providerApiKey is not set. */
  apiKey?: string
}

const DEFAULT_CAPS: ProviderCapabilities = { tools: true, vision: false, streaming: true }

export class OpenAiCompatibleAdapter implements ProviderAdapter {
  readonly id: string
  readonly name: string
  readonly kind: 'local' | 'cloud'
  readonly capabilities: ProviderCapabilities
  private cfg: OpenAiCompatibleConfig

  constructor(cfg: OpenAiCompatibleConfig) {
    this.id = cfg.id
    this.name = cfg.name
    this.kind = cfg.kind
    this.capabilities = { ...DEFAULT_CAPS, ...cfg.capabilities }
    this.cfg = cfg
  }

  isConfigured(): boolean {
    if (this.cfg.isConfiguredOverride) return this.cfg.isConfiguredOverride()
    if (this.kind === 'local') return true
    if (this.cfg.requiresKey === false) return true
    return providerApiKey(this.id) !== null
  }

  private apiKey(): string | null {
    return providerApiKey(this.id) ?? this.cfg.apiKey ?? null
  }

  private authHeaders(): Record<string, string> {
    const style = this.cfg.authHeaderStyle ?? 'bearer'
    if (style === 'none') return {}
    const key = this.apiKey()
    if (!key) {
      if (this.kind === 'local') return {}
      if (this.cfg.requiresKey === false) return {}
      throw new AuthError(`no API key configured for ${this.id}`)
    }
    if (style === 'api-key-header') {
      const header = this.cfg.apiKeyHeader ?? 'api-key'
      return { [header]: key }
    }
    return { Authorization: `Bearer ${key}` }
  }

  private buildHeaders(extra?: Record<string, string>): Record<string, string> {
    return {
      'content-type': 'application/json',
      ...this.authHeaders(),
      ...(this.cfg.extraHeaders ?? {}),
      ...(extra ?? {}),
    }
  }

  async isReachable(): Promise<boolean> {
    try {
      const url = joinUrl(this.cfg.baseUrl, '/models')
      const res = await fetchWithRetry(
        () => fetch(url, { method: 'GET', headers: this.buildHeaders() }),
        { maxAttempts: 1, logTag: this.id },
      )
      return res.ok
    } catch {
      return false
    }
  }

  async listModels(): Promise<ModelInfo[]> {
    const url = joinUrl(this.cfg.baseUrl, '/models')
    const res = await fetchWithRetry(() => fetch(url, { headers: this.buildHeaders() }), {
      logTag: this.id,
      isIdempotent: true,
    })
    if (!res.ok) {
      if (res.status === 404) {
        return fallbackModels(this.id, this.name, this.kind, this.capabilities, this.cfg.modelsHint)
      }
      const body = await res.text().catch(() => '')
      throw classifyHttpStatus(res.status, body)
    }
    const json = (await res.json().catch(() => ({}))) as {
      data?: Array<{ id: string; owned_by?: string }>
    }
    const list = json.data ?? []
    if (list.length === 0) {
      return fallbackModels(this.id, this.name, this.kind, this.capabilities, this.cfg.modelsHint)
    }
    return list.map((m) => ({
      id: `${this.id}/${m.id}`,
      name: m.id,
      providerId: this.id,
      providerName: this.name,
      kind: this.kind,
      capabilities: this.capabilities,
    }))
  }

  async *stream(req: ChatRequest, signal: AbortSignal): AsyncIterable<StreamEvent> {
    const modelLocal = stripProviderPrefix(req.model, this.id)
    const url = joinUrl(this.cfg.baseUrl, '/chat/completions')
    const body = buildChatCompletionsBody(req, modelLocal, this.capabilities)
    let res: Response
    try {
      res = await fetchWithRetry(
        () =>
          fetch(url, {
            method: 'POST',
            headers: this.buildHeaders({ accept: 'text/event-stream' }),
            body: JSON.stringify(body),
            signal,
          }),
        { signal, logTag: this.id, isIdempotent: false },
      )
    } catch (e) {
      if (signal.aborted) return
      if ((e as Error).message === 'aborted') return
      throw e instanceof ProviderError ? e : new NetworkError(String((e as Error).message ?? e))
    }
    if (!res.ok) {
      const text = await res.text().catch(() => '')
      if (signal.aborted) return
      yield yieldStreamError(classifyHttpStatus(res.status, text))
      return
    }
    if (!res.body) {
      yield { type: 'error', code: 'upstream', message: 'no response body', retryable: true }
      return
    }
    let promptTokens = 0
    let completionTokens = 0
    const toolCalls = new ToolCallAccumulator()
    try {
      for await (const ev of parseSseStream(res.body)) {
        if (signal.aborted) return
        if (ev.data && isDoneSentinel(ev.data)) break
        if (!ev.data) continue
        let obj: unknown
        try {
          obj = JSON.parse(ev.data)
        } catch {
          continue
        }
        const chunk = obj as {
          choices?: Array<{
            delta?: {
              content?: string
              // DeepSeek and Groq use `reasoning_content`; OpenRouter and xAI
              // use `reasoning`. Neither is part of the OpenAI spec, so both
              // have to be probed for.
              reasoning_content?: string
              reasoning?: string
              tool_calls?: ToolCallDelta[]
            }
            finish_reason?: string
          }>
          usage?: { prompt_tokens?: number; completion_tokens?: number; total_tokens?: number }
        }
        if (chunk.usage) {
          promptTokens = chunk.usage.prompt_tokens ?? promptTokens
          completionTokens = chunk.usage.completion_tokens ?? completionTokens
        }
        const choice = chunk.choices?.[0]
        if (choice) {
          const delta = choice.delta
          const reasoning = delta?.reasoning_content ?? delta?.reasoning
          if (reasoning) {
            yield { type: 'reasoning', text: reasoning }
          }
          if (delta?.content) {
            yield { type: 'delta', text: delta.content }
          }
          if (delta?.tool_calls) toolCalls.absorb(delta.tool_calls)
          if (choice.finish_reason) yield* toolCalls.drain()
        }
      }
      // Not every server sends a terminal `finish_reason` — some close the
      // stream after `[DONE]`, or after the last content delta. Draining here
      // as well is what keeps a tool call from being silently dropped.
      yield* toolCalls.drain()
      if (promptTokens > 0 || completionTokens > 0) {
        yield {
          type: 'usage',
          promptTokens,
          completionTokens,
          totalTokens: promptTokens + completionTokens,
        }
        const cost = estimateCost(req.model, promptTokens, completionTokens)
        if (cost > 0) yield { type: 'cost', estimatedCostUsd: cost }
      }
    } catch (e) {
      if (signal.aborted) return
      yield toStreamEvent(e)
    }
  }
}

function yieldStreamError(err: ProviderError): StreamEvent {
  return { type: 'error', code: err.code, message: err.message, retryable: err.retryable }
}

export interface ToolCallDelta {
  index?: number
  id?: string
  function?: { name?: string; arguments?: string }
}

/**
 * Reassembles streamed tool calls.
 *
 * Arguments arrive as a run of string fragments that have to be concatenated in
 * order and exactly once. Getting that wrong is not a display bug: the joined
 * string goes straight back upstream on the next round, and a server that
 * parses it — Ollama does — answers HTTP 400 "invalid tool call arguments".
 *
 * Fragments are correlated by `index` per the OpenAI spec. Servers that omit it
 * fall back to `id`, and a fragment carrying neither continues the call it was
 * most recently appended to, which is the only reading available.
 */
export class ToolCallAccumulator {
  private buffers = new Map<string, { id: string; name: string; args: string }>()
  private lastKey: string | null = null

  absorb(deltas: ToolCallDelta[]): void {
    for (const tc of deltas) {
      const key = this.keyFor(tc)
      let buf = this.buffers.get(key)
      if (!buf) {
        buf = { id: tc.id ?? '', name: '', args: '' }
        this.buffers.set(key, buf)
      }
      if (tc.id) buf.id = tc.id
      if (tc.function?.name) buf.name = tc.function.name
      if (tc.function?.arguments) buf.args += tc.function.arguments
      this.lastKey = key
    }
  }

  private keyFor(tc: ToolCallDelta): string {
    if (typeof tc.index === 'number') return `i:${tc.index}`
    if (tc.id) return `id:${tc.id}`
    return this.lastKey ?? 'i:0'
  }

  /** Emit and forget everything buffered so far. Safe to call repeatedly. */
  *drain(): Generator<StreamEvent> {
    for (const buf of this.buffers.values()) {
      // A call with no name is not a call — some servers emit an empty
      // tool_calls frame alongside the finish reason.
      if (!buf.name) continue
      const toolCall: ToolCall = {
        id: buf.id || `call_${Math.random().toString(36).slice(2, 10)}`,
        name: buf.name,
        arguments: normalizeToolArguments(buf.args),
      }
      yield { type: 'tool_call', toolCall }
    }
    this.buffers.clear()
    this.lastKey = null
  }
}

/**
 * Coerce tool-call arguments to a string that parses to a JSON object.
 *
 * Applied both on the way out of a stream and on the way back upstream: a
 * transcript loaded from localStorage may predate this fix and still hold a
 * doubled fragment, and one bad string in history poisons every later turn of
 * that conversation with a 400.
 */
export function normalizeToolArguments(raw: string | undefined): string {
  const trimmed = (raw ?? '').trim()
  if (trimmed === '') return '{}'
  try {
    const parsed: unknown = JSON.parse(trimmed)
    if (parsed && typeof parsed === 'object' && !Array.isArray(parsed))
      return JSON.stringify(parsed)
  } catch {
    // fall through
  }
  logger.warn({ msg: 'discarding unparseable tool call arguments', length: trimmed.length })
  return '{}'
}

function stripProviderPrefix(model: string, providerId: string): string {
  if (model.startsWith(`${providerId}/`)) return model.slice(providerId.length + 1)
  return model
}

function joinUrl(base: string, path: string): string {
  const b = base.replace(/\/+$/, '')
  const p = path.startsWith('/') ? path : `/${path}`
  return b + p
}

function buildChatCompletionsBody(
  req: ChatRequest,
  model: string,
  caps: ProviderCapabilities,
): Record<string, unknown> {
  const messages = req.messages.map(toOpenAiMessage)
  const body: Record<string, unknown> = {
    model,
    messages,
    stream: true,
  }
  if (typeof req.temperature === 'number') body.temperature = req.temperature
  if (typeof req.maxTokens === 'number') body.max_tokens = req.maxTokens
  if (caps.tools && req.tools && req.tools.length > 0) {
    // `{function: {name}}` with no schema is not a valid tool definition — the
    // model is told the tool exists but not what it takes, so it either refuses
    // to call it or calls it with invented arguments.
    body.tools = req.tools.map((t) => ({
      type: 'function',
      function: { name: t.name, description: t.description, parameters: t.parameters },
    }))
  }
  return body
}

function toOpenAiMessage(m: ChatMessage): Record<string, unknown> {
  const out: Record<string, unknown> = { role: m.role, content: m.content }
  if (m.toolCalls && m.toolCalls.length > 0) {
    out.tool_calls = m.toolCalls.map((tc) => ({
      id: tc.id,
      type: 'function',
      function: { name: tc.name, arguments: normalizeToolArguments(tc.arguments) },
    }))
  }
  if (m.toolCallId) out.tool_call_id = m.toolCallId
  return out
}

function fallbackModels(
  providerId: string,
  providerName: string,
  kind: 'local' | 'cloud',
  caps: ProviderCapabilities,
  hints?: string[],
): ModelInfo[] {
  if (!hints || hints.length === 0) return []
  return hints.map((name) => ({
    id: `${providerId}/${name}`,
    name,
    providerId,
    providerName,
    kind,
    capabilities: caps,
  }))
}

export function createImportedAdapter(cfg: ImportedProviderConfig): OpenAiCompatibleAdapter {
  return new OpenAiCompatibleAdapter({
    id: cfg.id,
    name: cfg.name,
    baseUrl: cfg.baseUrl,
    kind: 'cloud',
    modelsHint: cfg.models,
  })
}

export function presetAdapters(): OpenAiCompatibleAdapter[] {
  return [
    new OpenAiCompatibleAdapter({
      id: 'lm-studio',
      name: 'LM Studio',
      baseUrl: 'http://127.0.0.1:1234/v1',
      kind: 'local',
      authHeaderStyle: 'none',
    }),
    new OpenAiCompatibleAdapter({
      id: 'llama-cpp',
      name: 'llama.cpp',
      baseUrl: 'http://127.0.0.1:8080/v1',
      kind: 'local',
      authHeaderStyle: 'none',
    }),
    new OpenAiCompatibleAdapter({
      id: 'openai',
      name: 'OpenAI',
      baseUrl: 'https://api.openai.com/v1',
      kind: 'cloud',
      capabilities: { tools: true, vision: true, streaming: true },
    }),
    new OpenAiCompatibleAdapter({
      id: 'openrouter',
      name: 'OpenRouter',
      baseUrl: 'https://openrouter.ai/api/v1',
      kind: 'cloud',
      extraHeaders: { 'HTTP-Referer': 'https://openofficellm.local', 'X-Title': 'OpenOfficeLLM' },
    }),
    new OpenAiCompatibleAdapter({
      id: 'groq',
      name: 'Groq',
      baseUrl: 'https://api.groq.com/openai/v1',
      kind: 'cloud',
    }),
    new OpenAiCompatibleAdapter({
      id: 'deepseek',
      name: 'DeepSeek',
      baseUrl: 'https://api.deepseek.com/v1',
      kind: 'cloud',
    }),
    new OpenAiCompatibleAdapter({
      id: 'together',
      name: 'Together',
      baseUrl: 'https://api.together.xyz/v1',
      kind: 'cloud',
    }),
    new OpenAiCompatibleAdapter({
      id: 'xai',
      name: 'xAI',
      baseUrl: 'https://api.x.ai/v1',
      kind: 'cloud',
    }),
  ]
}
