import type {
  ChatMessage,
  ChatRequest,
  ModelInfo,
  ProviderCapabilities,
  StreamEvent,
  ToolCall,
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
import { fetchWithRetry } from './retry.js'
import { parseSseStream } from './sse-parser.js'
import { estimateCost } from './pricing.js'

const BASE = 'https://api.anthropic.com/v1'
const ANTHROPIC_VERSION = '2023-06-01'
const CAPS: ProviderCapabilities = { tools: true, vision: true, streaming: true }

/** Fallback list used only when GET /v1/models is unreachable. The dated 3.x
 *  ids that used to live here have been retired by Anthropic and now 404, so a
 *  user whose network briefly hiccuped was offered three models that could not
 *  answer. Aliases (no date suffix) always resolve to the current snapshot. */
const KNOWN_MODELS = [
  { id: 'claude-opus-5', name: 'Claude Opus 5', contextWindow: 1000000 },
  { id: 'claude-sonnet-5', name: 'Claude Sonnet 5', contextWindow: 1000000 },
  { id: 'claude-haiku-4-5', name: 'Claude Haiku 4.5', contextWindow: 200000 },
]

export class AnthropicAdapter implements ProviderAdapter {
  readonly id = 'anthropic'
  readonly name = 'Anthropic'
  readonly kind = 'cloud' as const
  readonly capabilities: ProviderCapabilities = CAPS

  isConfigured(): boolean {
    return providerApiKey(this.id) !== null
  }

  private apiKey(): string {
    const k = providerApiKey(this.id)
    if (!k) throw new AuthError('no Anthropic API key configured')
    return k
  }

  private headers(extra?: Record<string, string>): Record<string, string> {
    return {
      'content-type': 'application/json',
      'x-api-key': this.apiKey(),
      'anthropic-version': ANTHROPIC_VERSION,
      ...(extra ?? {}),
    }
  }

  async isReachable(): Promise<boolean> {
    try {
      const res = await fetchWithRetry(
        () => fetch(`${BASE}/models`, { method: 'GET', headers: this.headers() }),
        { maxAttempts: 1, logTag: this.id },
      )
      return res.ok
    } catch {
      return false
    }
  }

  async listModels(): Promise<ModelInfo[]> {
    try {
      const res = await fetchWithRetry(() => fetch(`${BASE}/models`, { headers: this.headers() }), {
        logTag: this.id,
      })
      if (res.ok) {
        const json = (await res.json().catch(() => ({}))) as {
          data?: Array<{ id: string; display_name?: string }>
        }
        if (json.data && json.data.length > 0) {
          return json.data.map((m) => ({
            id: `${this.id}/${m.id}`,
            name: m.display_name ?? m.id,
            providerId: this.id,
            providerName: this.name,
            kind: 'cloud',
            capabilities: CAPS,
          }))
        }
      }
    } catch {
      // fall through to static list
    }
    return KNOWN_MODELS.map((m) => ({
      id: `${this.id}/${m.id}`,
      name: m.name,
      providerId: this.id,
      providerName: this.name,
      kind: 'cloud',
      capabilities: CAPS,
      contextWindow: m.contextWindow,
    }))
  }

  async *stream(req: ChatRequest, signal: AbortSignal): AsyncIterable<StreamEvent> {
    const modelLocal = stripProviderPrefix(req.model, this.id)
    const url = `${BASE}/messages`
    const body = buildAnthropicBody(req, modelLocal)
    let res: Response
    try {
      res = await fetchWithRetry(
        () =>
          fetch(url, {
            method: 'POST',
            headers: this.headers({ accept: 'text/event-stream' }),
            body: JSON.stringify(body),
            signal,
          }),
        { signal, logTag: this.id, isIdempotent: false },
      )
    } catch (e) {
      if (signal.aborted || (e as Error).message === 'aborted') return
      throw e instanceof ProviderError ? e : new NetworkError(String((e as Error).message ?? e))
    }
    if (!res.ok) {
      const text = await res.text().catch(() => '')
      if (signal.aborted) return
      const err = classifyHttpStatus(res.status, text)
      yield { type: 'error', code: err.code, message: err.message, retryable: err.retryable }
      return
    }
    if (!res.body) {
      yield { type: 'error', code: 'upstream', message: 'no response body', retryable: true }
      return
    }
    let promptTokens = 0
    let completionTokens = 0
    const toolBuffers = new Map<number, { id: string; name: string; args: string }>()
    try {
      for await (const ev of parseSseStream(res.body)) {
        if (signal.aborted) return
        if (!ev.event || !ev.data) continue
        let obj: unknown
        try {
          obj = JSON.parse(ev.data)
        } catch {
          continue
        }
        switch (ev.event) {
          case 'message_start': {
            const m = obj as { message?: { usage?: { input_tokens?: number } } }
            promptTokens = m.message?.usage?.input_tokens ?? promptTokens
            break
          }
          case 'content_block_start': {
            const m = obj as {
              index?: number
              content_block?: { type?: string; id?: string; name?: string }
            }
            if (m.content_block?.type === 'tool_use' && typeof m.index === 'number') {
              toolBuffers.set(m.index, {
                id: m.content_block.id ?? '',
                name: m.content_block.name ?? '',
                args: '',
              })
            }
            break
          }
          case 'content_block_delta': {
            const m = obj as {
              index?: number
              delta?: { type?: string; text?: string; partial_json?: string; thinking?: string }
            }
            if (m.delta?.type === 'text_delta' && m.delta.text) {
              yield { type: 'delta', text: m.delta.text }
            } else if (m.delta?.type === 'thinking_delta' && m.delta.thinking) {
              yield { type: 'reasoning', text: m.delta.thinking }
            } else if (m.delta?.type === 'input_json_delta' && typeof m.index === 'number') {
              const buf = toolBuffers.get(m.index)
              if (buf && m.delta.partial_json) buf.args += m.delta.partial_json
            }
            break
          }
          case 'content_block_stop': {
            const m = obj as { index?: number }
            if (typeof m.index === 'number' && toolBuffers.has(m.index)) {
              const buf = toolBuffers.get(m.index)!
              const tc: ToolCall = {
                id: buf.id,
                name: buf.name,
                arguments: buf.args || '{}',
              }
              yield { type: 'tool_call', toolCall: tc }
              toolBuffers.delete(m.index)
            }
            break
          }
          case 'message_delta': {
            const m = obj as { usage?: { output_tokens?: number } }
            completionTokens = m.usage?.output_tokens ?? completionTokens
            break
          }
          case 'message_stop': {
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
            break
          }
          default:
            break
        }
      }
    } catch (e) {
      if (signal.aborted) return
      yield toStreamEvent(e)
    }
  }
}

function stripProviderPrefix(model: string, providerId: string): string {
  if (model.startsWith(`${providerId}/`)) return model.slice(providerId.length + 1)
  return model
}

function buildAnthropicBody(req: ChatRequest, model: string): Record<string, unknown> {
  const systemMessages = req.messages.filter((m) => m.role === 'system')
  const nonSystem = req.messages.filter((m) => m.role !== 'system')
  const systemText = [systemMessages.map((m) => m.content).join('\n'), req.systemPrompt ?? '']
    .filter(Boolean)
    .join('\n\n')
  const body: Record<string, unknown> = {
    model,
    messages: nonSystem.map(toAnthropicMessage),
    max_tokens: req.maxTokens ?? 4096,
    stream: true,
  }
  if (systemText) body.system = systemText
  if (typeof req.temperature === 'number') body.temperature = req.temperature
  if (req.tools && req.tools.length > 0) {
    body.tools = req.tools.map((t) => ({
      name: t.name,
      description: t.description,
      input_schema: t.parameters,
    }))
  }
  return body
}

function toAnthropicMessage(m: ChatMessage): Record<string, unknown> {
  if (m.role === 'tool' && m.toolCallId) {
    return {
      role: 'user',
      content: [{ type: 'tool_result', tool_use_id: m.toolCallId, content: m.content }],
    }
  }
  if (m.toolCalls && m.toolCalls.length > 0) {
    return {
      role: m.role,
      content: [
        ...(m.content ? [{ type: 'text', text: m.content }] : []),
        ...m.toolCalls.map((tc) => ({
          type: 'tool_use',
          id: tc.id,
          name: tc.name,
          input: safeParseArgs(tc.arguments),
        })),
      ],
    }
  }
  return { role: m.role, content: m.content }
}

function safeParseArgs(args: string): unknown {
  try {
    return JSON.parse(args)
  } catch {
    return {}
  }
}
