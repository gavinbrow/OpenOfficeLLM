import type {
  ChatMessage,
  ChatRequest,
  ModelInfo,
  ProviderCapabilities,
  StreamEvent,
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
import { parseSseStream, isDoneSentinel } from './sse-parser.js'
import {
  ToolCallAccumulator,
  normalizeToolArguments,
  type ToolCallDelta,
} from './openai-compatible.js'

const CAPS: ProviderCapabilities = { tools: true, vision: true, streaming: true }

export interface AzureConfig {
  baseUrl: string
  apiVersion: string
  deployments?: string[]
}

export class AzureOpenAiAdapter implements ProviderAdapter {
  readonly id = 'azure-openai'
  readonly name = 'Azure OpenAI'
  readonly kind = 'cloud' as const
  readonly capabilities: ProviderCapabilities = CAPS
  private cfg: AzureConfig

  constructor(cfg: AzureConfig) {
    this.cfg = cfg
  }

  isConfigured(): boolean {
    return providerApiKey(this.id) !== null
  }

  private apiKey(): string {
    const k = providerApiKey(this.id)
    if (!k) throw new AuthError('no Azure OpenAI key configured')
    return k
  }

  private deploymentFromModel(model: string): string {
    return stripProviderPrefix(model, this.id)
  }

  async isReachable(): Promise<boolean> {
    if (!this.cfg.deployments || this.cfg.deployments.length === 0) return false
    try {
      const url = `${this.cfg.baseUrl}/openai/deployments?api-version=${this.cfg.apiVersion}`
      const res = await fetchWithRetry(
        () => fetch(url, { method: 'GET', headers: { 'api-key': this.apiKey() } }),
        { maxAttempts: 1, logTag: this.id },
      )
      return res.ok
    } catch {
      return false
    }
  }

  async listModels(): Promise<ModelInfo[]> {
    const deployments = this.cfg.deployments ?? []
    return deployments.map((d) => ({
      id: `${this.id}/${d}`,
      name: d,
      providerId: this.id,
      providerName: this.name,
      kind: 'cloud',
      capabilities: CAPS,
    }))
  }

  async *stream(req: ChatRequest, signal: AbortSignal): AsyncIterable<StreamEvent> {
    const deployment = this.deploymentFromModel(req.model)
    const url = `${this.cfg.baseUrl}/openai/deployments/${encodeURIComponent(deployment)}/chat/completions?api-version=${this.cfg.apiVersion}`
    const body = buildAzureBody(req)
    let res: Response
    try {
      res = await fetchWithRetry(
        () =>
          fetch(url, {
            method: 'POST',
            headers: {
              'content-type': 'application/json',
              'api-key': this.apiKey(),
              accept: 'text/event-stream',
            },
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
    const toolBuffers = new ToolCallAccumulator()
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
          if (choice.delta?.content) yield { type: 'delta', text: choice.delta.content }
          if (choice.delta?.tool_calls) toolBuffers.absorb(choice.delta.tool_calls)
          if (choice.finish_reason) yield* toolBuffers.drain()
        }
      }
      yield* toolBuffers.drain()
      if (promptTokens > 0 || completionTokens > 0) {
        yield {
          type: 'usage',
          promptTokens,
          completionTokens,
          totalTokens: promptTokens + completionTokens,
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

function buildAzureBody(req: ChatRequest): Record<string, unknown> {
  const messages = req.messages.map((m) => ({
    role: m.role,
    content: m.content,
    ...(m.toolCalls
      ? {
          tool_calls: m.toolCalls.map((tc) => ({
            id: tc.id,
            type: 'function',
            function: { name: tc.name, arguments: normalizeToolArguments(tc.arguments) },
          })),
        }
      : {}),
    ...(m.toolCallId ? { tool_call_id: m.toolCallId } : {}),
  }))
  const body: Record<string, unknown> = { messages, stream: true }
  if (typeof req.temperature === 'number') body.temperature = req.temperature
  if (typeof req.maxTokens === 'number') body.max_tokens = req.maxTokens
  if (req.tools && req.tools.length > 0) {
    body.tools = req.tools.map((t) => ({
      type: 'function',
      function: { name: t.name, description: t.description, parameters: t.parameters },
    }))
  }
  return body
}

export { CAPS as AZURE_CAPS }

export type { ChatMessage }
