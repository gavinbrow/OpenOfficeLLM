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

const BASE = 'https://generativelanguage.googleapis.com/v1beta'
const CAPS: ProviderCapabilities = { tools: true, vision: true, streaming: true }

const KNOWN_MODELS = [
  { id: 'gemini-1.5-pro', name: 'Gemini 1.5 Pro', contextWindow: 1000000 },
  { id: 'gemini-1.5-flash', name: 'Gemini 1.5 Flash', contextWindow: 1000000 },
  { id: 'gemini-2.0-flash', name: 'Gemini 2.0 Flash', contextWindow: 1000000 },
]

export class GoogleAdapter implements ProviderAdapter {
  readonly id = 'google'
  readonly name = 'Google'
  readonly kind = 'cloud' as const
  readonly capabilities: ProviderCapabilities = CAPS

  isConfigured(): boolean {
    return providerApiKey(this.id) !== null
  }

  private apiKey(): string {
    const k = providerApiKey(this.id)
    if (!k) throw new AuthError('no Google API key configured')
    return k
  }

  async isReachable(): Promise<boolean> {
    try {
      const url = `${BASE}/models`
      const res = await fetchWithRetry(
        () => fetch(url, { method: 'GET', headers: { 'x-goog-api-key': this.apiKey() } }),
        {
          maxAttempts: 1,
          logTag: this.id,
        },
      )
      return res.ok
    } catch {
      return false
    }
  }

  async listModels(): Promise<ModelInfo[]> {
    try {
      const url = `${BASE}/models`
      const res = await fetchWithRetry(
        () => fetch(url, { headers: { 'x-goog-api-key': this.apiKey() } }),
        { logTag: this.id },
      )
      if (res.ok) {
        const json = (await res.json().catch(() => ({}))) as {
          models?: Array<{
            name?: string
            displayName?: string
            supportedGenerationMethods?: string[]
          }>
        }
        if (json.models && json.models.length > 0) {
          return json.models
            .filter(
              (m) =>
                Array.isArray(m.supportedGenerationMethods) &&
                m.supportedGenerationMethods.includes('generateContent'),
            )
            .map((m) => {
              const localId = (m.name ?? 'unknown').replace(/^models\//, '')
              return {
                id: `${this.id}/${localId}`,
                name: m.displayName ?? localId,
                providerId: this.id,
                providerName: this.name,
                kind: 'cloud',
                capabilities: CAPS,
              } satisfies ModelInfo
            })
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
    const key = this.apiKey()
    const url = `${BASE}/models/${encodeURIComponent(modelLocal)}:streamGenerateContent?alt=sse`
    const body = buildGeminiBody(req)
    let res: Response
    try {
      res = await fetchWithRetry(
        () =>
          fetch(url, {
            method: 'POST',
            headers: {
              'content-type': 'application/json',
              'x-goog-api-key': key,
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
    try {
      for await (const ev of parseSseStream(res.body)) {
        if (signal.aborted) return
        if (!ev.data) continue
        let obj: unknown
        try {
          obj = JSON.parse(ev.data)
        } catch {
          continue
        }
        const chunk = obj as {
          candidates?: Array<{
            content?: {
              parts?: Array<{ text?: string; functionCall?: { name?: string; args?: unknown } }>
            }
            finishReason?: string
          }>
          usageMetadata?: {
            promptTokenCount?: number
            candidatesTokenCount?: number
            totalTokenCount?: number
          }
        }
        if (chunk.usageMetadata) {
          promptTokens = chunk.usageMetadata.promptTokenCount ?? promptTokens
          completionTokens = chunk.usageMetadata.candidatesTokenCount ?? completionTokens
        }
        const cand = chunk.candidates?.[0]
        if (cand?.content?.parts) {
          for (const part of cand.content.parts) {
            if (part.text) yield { type: 'delta', text: part.text }
            if (part.functionCall) {
              const tc: ToolCall = {
                id: `call_${Math.random().toString(36).slice(2, 10)}`,
                name: part.functionCall.name ?? '',
                arguments: JSON.stringify(part.functionCall.args ?? {}),
              }
              yield { type: 'tool_call', toolCall: tc }
            }
          }
        }
      }
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

function stripProviderPrefix(model: string, providerId: string): string {
  if (model.startsWith(`${providerId}/`)) return model.slice(providerId.length + 1)
  return model
}

function buildGeminiBody(req: ChatRequest): Record<string, unknown> {
  const systemText = req.messages
    .filter((m) => m.role === 'system')
    .map((m) => m.content)
    .concat(req.systemPrompt ? [req.systemPrompt] : [])
    .join('\n\n')
  const contents = req.messages.filter((m) => m.role !== 'system').map(toGeminiContent)
  const body: Record<string, unknown> = { contents }
  if (systemText) body.systemInstruction = { parts: [{ text: systemText }] }
  const genConfig: Record<string, unknown> = {}
  if (typeof req.temperature === 'number') genConfig.temperature = req.temperature
  if (typeof req.maxTokens === 'number') genConfig.maxOutputTokens = req.maxTokens
  if (Object.keys(genConfig).length > 0) body.generationConfig = genConfig
  if (req.tools && req.tools.length > 0) {
    body.tools = [
      {
        functionDeclarations: req.tools.map((t) => ({
          name: t.name,
          description: t.description,
          parameters: t.parameters,
        })),
      },
    ]
  }
  return body
}

function toGeminiContent(m: ChatMessage): Record<string, unknown> {
  const role = m.role === 'assistant' ? 'model' : m.role === 'tool' ? 'user' : m.role
  const parts: unknown[] = []
  if (m.content) parts.push({ text: m.content })
  if (m.toolCalls) {
    for (const tc of m.toolCalls) {
      parts.push({ functionCall: { name: tc.name, args: safeParseArgs(tc.arguments) } })
    }
  }
  if (m.role === 'tool' && m.toolCallId) {
    return {
      role: 'user',
      parts: [{ functionResponse: { name: m.toolCallId, response: { result: m.content } } }],
    }
  }
  return { role, parts }
}

function safeParseArgs(args: string): unknown {
  try {
    return JSON.parse(args)
  } catch {
    return {}
  }
}
