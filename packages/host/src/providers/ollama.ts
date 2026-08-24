import type {
  ChatMessage,
  ChatRequest,
  ModelInfo,
  ProviderCapabilities,
  StreamEvent,
  ToolCall,
} from '@openofficellm/shared'
import {
  NetworkError,
  ProviderError,
  classifyHttpStatus,
  toStreamEvent,
  type ProviderAdapter,
} from './types.js'
import { fetchWithRetry } from './retry.js'
import { parseNdjsonStream } from './sse-parser.js'

const DEFAULT_BASE = 'http://127.0.0.1:11434'
const CAPS: ProviderCapabilities = { tools: true, vision: true, streaming: true }

export interface OllamaConfig {
  baseUrl?: string
  id?: string
  name?: string
}

export class OllamaAdapter implements ProviderAdapter {
  readonly id: string
  readonly name: string
  readonly kind = 'local' as const
  readonly capabilities: ProviderCapabilities = CAPS
  private baseUrl: string

  constructor(cfg: OllamaConfig = {}) {
    this.id = cfg.id ?? 'ollama'
    this.name = cfg.name ?? 'Ollama'
    this.baseUrl = (cfg.baseUrl ?? process.env.OLLAMA_HOST ?? DEFAULT_BASE).replace(/\/+$/, '')
  }

  /** Allow the settings UI to re-point local Ollama after the adapter is created. */
  setBaseUrl(url: string): void {
    this.baseUrl = url.replace(/\/+$/, '')
  }

  isConfigured(): boolean {
    return true
  }

  async isReachable(): Promise<boolean> {
    try {
      const res = await fetchWithRetry(() => fetch(`${this.baseUrl}/api/tags`, { method: 'GET' }), {
        maxAttempts: 1,
        logTag: this.id,
      })
      return res.ok
    } catch {
      return false
    }
  }

  async listModels(): Promise<ModelInfo[]> {
    const res = await fetchWithRetry(() => fetch(`${this.baseUrl}/api/tags`), {
      logTag: this.id,
      isIdempotent: true,
    })
    if (!res.ok) {
      const body = await res.text().catch(() => '')
      throw classifyHttpStatus(res.status, body)
    }
    const json = (await res.json().catch(() => ({}))) as {
      models?: Array<{
        name?: string
        model?: string
        size?: number
        details?: {
          quantization_level?: string
          family?: string
          parameter_size?: string
          context_length?: number
        }
        capabilities?: string[]
      }>
    }
    const models = json.models ?? []
    return models.map((m) => {
      const localName = m.model ?? m.name ?? 'unknown'
      const caps: ProviderCapabilities = {
        tools: Array.isArray(m.capabilities) && m.capabilities.includes('tools'),
        vision: false,
        streaming: true,
      }
      return {
        id: `${this.id}/${localName}`,
        name: localName,
        providerId: this.id,
        providerName: this.name,
        kind: 'local',
        capabilities: caps,
        sizeBytes: m.size,
        quantization: m.details?.quantization_level || undefined,
        contextWindow: m.details?.context_length,
      } satisfies ModelInfo
    })
  }

  async *stream(req: ChatRequest, signal: AbortSignal): AsyncIterable<StreamEvent> {
    const modelLocal = stripProviderPrefix(req.model, this.id)
    const body = buildOllamaChatBody(req, modelLocal)
    let res: Response
    try {
      res = await fetchWithRetry(
        () =>
          fetch(`${this.baseUrl}/api/chat`, {
            method: 'POST',
            headers: { 'content-type': 'application/json' },
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
    try {
      for await (const ev of parseNdjsonStream(res.body)) {
        if (signal.aborted) return
        const obj = ev.obj as {
          message?: {
            role?: string
            content?: string
            thinking?: string
            tool_calls?: Array<{ function?: { name?: string; arguments?: unknown } }>
          }
          done?: boolean
          total_duration?: number
          prompt_eval_count?: number
          eval_count?: number
        }
        // `thinking` and `content` can both be present on one chunk, and
        // `thinking` is the model's scratchpad — emitting it as a delta put the
        // chain-of-thought straight into the answer.
        if (obj.message?.thinking) {
          yield { type: 'reasoning', text: obj.message.thinking }
        }
        if (obj.message?.content) {
          yield { type: 'delta', text: obj.message.content }
        }
        if (obj.message?.tool_calls && Array.isArray(obj.message.tool_calls)) {
          for (const tc of obj.message.tool_calls) {
            const toolCall: ToolCall = {
              id: `call_${Math.random().toString(36).slice(2, 10)}`,
              name: tc.function?.name ?? '',
              arguments:
                typeof tc.function?.arguments === 'string'
                  ? tc.function.arguments
                  : JSON.stringify(tc.function?.arguments ?? {}),
            }
            yield { type: 'tool_call', toolCall }
          }
        }
        if (obj.done) {
          promptTokens = obj.prompt_eval_count ?? promptTokens
          completionTokens = obj.eval_count ?? completionTokens
          if (promptTokens > 0 || completionTokens > 0) {
            yield {
              type: 'usage',
              promptTokens,
              completionTokens,
              totalTokens: promptTokens + completionTokens,
            }
          }
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

function buildOllamaChatBody(req: ChatRequest, model: string): Record<string, unknown> {
  const messages = req.messages.map(toOllamaMessage)
  const body: Record<string, unknown> = {
    model,
    messages,
    stream: true,
  }
  if (typeof req.temperature === 'number') body.options = { temperature: req.temperature }
  if (typeof req.maxTokens === 'number') {
    body.options = { ...(body.options as Record<string, unknown>), num_predict: req.maxTokens }
  }
  if (req.tools && req.tools.length > 0) {
    // Ollama follows the OpenAI tool shape. The previous version synthesized
    // `{name, description: name, parameters: {type:'object'}}` from a bare
    // string — a schema with no properties, which tells the model every tool
    // takes no arguments.
    body.tools = req.tools.map((t) => ({
      type: 'function',
      function: { name: t.name, description: t.description, parameters: t.parameters },
    }))
  }
  return body
}

function toOllamaMessage(m: ChatMessage): Record<string, unknown> {
  const out: Record<string, unknown> = { role: m.role, content: m.content }
  if (m.toolCalls && m.toolCalls.length > 0) {
    out.tool_calls = m.toolCalls.map((tc) => ({
      type: 'function',
      function: { name: tc.name, arguments: safeParseArgs(tc.arguments) },
    }))
  }
  // Ollama identifies a tool result by the name of the tool that produced it,
  // not by a call id.
  if (m.role === 'tool' && m.toolName) out.tool_name = m.toolName
  return out
}

function safeParseArgs(args: string): unknown {
  try {
    return JSON.parse(args)
  } catch {
    return {}
  }
}
