import type {
  ChatMessage,
  ChatRequest,
  ModelInfo,
  ProviderCapabilities,
  StreamEvent,
  ToolCall,
} from '@openofficellm/shared'
import { providerApiKey } from './credentials.js'
import { NetworkError, ProviderError, toStreamEvent, type ProviderAdapter } from './types.js'
import { fetchWithRetry } from './retry.js'
import { parseSseStream } from './sse-parser.js'
import { logger } from '../logging.js'

const DEFAULT_BASE = 'http://127.0.0.1:4096'
const CAPS: ProviderCapabilities = { tools: true, vision: false, streaming: true }

export interface OpencodeConfig {
  baseUrl?: string
  password?: string
}

export class OpencodeAdapter implements ProviderAdapter {
  readonly id = 'opencode'
  readonly name = 'opencode'
  readonly kind = 'local' as const
  readonly capabilities: ProviderCapabilities = CAPS
  private baseUrl: string
  private password: string | undefined

  constructor(cfg: OpencodeConfig = {}) {
    this.baseUrl = (cfg.baseUrl ?? DEFAULT_BASE).replace(/\/+$/, '')
    this.password = cfg.password ?? process.env.OPENCODE_SERVER_PASSWORD
  }

  isConfigured(): boolean {
    return true
  }

  private authHeaders(): Record<string, string> {
    const pw = this.password ?? providerApiKey(this.id)
    if (!pw) return {}
    const raw = `:${pw}`
    return { Authorization: `Basic ${Buffer.from(raw, 'utf8').toString('base64')}` }
  }

  async isReachable(): Promise<boolean> {
    try {
      const res = await fetchWithRetry(
        () => fetch(`${this.baseUrl}/health`, { method: 'GET', headers: this.authHeaders() }),
        { maxAttempts: 1, logTag: this.id },
      )
      return res.ok
    } catch {
      return false
    }
  }

  async listModels(): Promise<ModelInfo[]> {
    try {
      const res = await fetchWithRetry(
        () => fetch(`${this.baseUrl}/model`, { headers: this.authHeaders() }),
        { logTag: this.id, isIdempotent: true },
      )
      if (!res.ok) return []
      const json = (await res.json().catch(() => ({}))) as {
        models?: Array<{ id?: string; name?: string }>
      }
      const list = json.models ?? []
      return list.map((m) => ({
        id: `${this.id}/${m.id ?? m.name ?? 'unknown'}`,
        name: m.name ?? m.id ?? 'unknown',
        providerId: this.id,
        providerName: this.name,
        kind: 'local',
        capabilities: CAPS,
      }))
    } catch {
      return []
    }
  }

  async *stream(req: ChatRequest, signal: AbortSignal): AsyncIterable<StreamEvent> {
    const prompt = buildOpencodePrompt(req)
    let sessionId: string
    try {
      const createRes = await fetchWithRetry(
        () =>
          fetch(`${this.baseUrl}/session`, {
            method: 'POST',
            headers: { 'content-type': 'application/json', ...this.authHeaders() },
            body: JSON.stringify({ title: req.conversationId ?? 'openofficellm' }),
            signal,
          }),
        { signal, logTag: this.id, isIdempotent: false },
      )
      if (!createRes.ok) {
        const text = await createRes.text().catch(() => '')
        yield {
          type: 'error',
          code: 'upstream',
          message: `opencode session create failed: ${text.slice(0, 200)}`,
          retryable: true,
        }
        return
      }
      const created = (await createRes.json().catch(() => ({}))) as { id?: string }
      sessionId = created.id ?? ''
      if (!sessionId) {
        yield {
          type: 'error',
          code: 'upstream',
          message: 'opencode returned no session id',
          retryable: false,
        }
        return
      }
    } catch (e) {
      if (signal.aborted || (e as Error).message === 'aborted') return
      yield toStreamEvent(
        e instanceof ProviderError ? e : new NetworkError(String((e as Error).message ?? e)),
      )
      return
    }

    try {
      const promptRes = await fetchWithRetry(
        () =>
          fetch(`${this.baseUrl}/session/${encodeURIComponent(sessionId)}/prompt_async`, {
            method: 'POST',
            headers: { 'content-type': 'application/json', ...this.authHeaders() },
            body: JSON.stringify({ prompt }),
            signal,
          }),
        { signal, logTag: this.id, isIdempotent: false },
      )
      if (!promptRes.ok) {
        const text = await promptRes.text().catch(() => '')
        yield {
          type: 'error',
          code: 'upstream',
          message: `opencode prompt failed: ${text.slice(0, 200)}`,
          retryable: true,
        }
        return
      }
    } catch (e) {
      if (signal.aborted || (e as Error).message === 'aborted') return
      yield toStreamEvent(
        e instanceof ProviderError ? e : new NetworkError(String((e as Error).message ?? e)),
      )
      return
    }

    try {
      const evRes = await fetch(`${this.baseUrl}/global/event`, {
        method: 'GET',
        headers: { accept: 'text/event-stream', ...this.authHeaders() },
        signal,
      })
      if (!evRes.ok || !evRes.body) {
        yield {
          type: 'error',
          code: 'upstream',
          message: 'opencode event stream unavailable',
          retryable: true,
        }
        return
      }
      for await (const ev of parseSseStream(evRes.body)) {
        if (signal.aborted) return
        if (!ev.data) continue
        let obj: unknown
        try {
          obj = JSON.parse(ev.data)
        } catch {
          continue
        }
        const mapped = mapOpencodeEvent(obj, sessionId)
        for (const se of mapped) yield se
      }
    } catch (e) {
      if (signal.aborted) return
      yield toStreamEvent(e)
    }
  }
}

function buildOpencodePrompt(req: ChatRequest): string {
  const lines = req.messages.map((m) => {
    const tag = m.role.toUpperCase()
    return `[${tag}]\n${m.content}`
  })
  if (req.systemPrompt) lines.unshift(`[SYSTEM]\n${req.systemPrompt}`)
  return lines.join('\n\n')
}

function mapOpencodeEvent(obj: unknown, sessionId: string): StreamEvent[] {
  const ev = obj as {
    type?: string
    sessionID?: string
    sessionId?: string
    message?: string
    text?: string
    toolCall?: { id?: string; name?: string; args?: string }
    usage?: { promptTokens?: number; completionTokens?: number }
    done?: boolean
  }
  const belongs = ev.sessionID === sessionId || ev.sessionId === sessionId || !ev.sessionID
  if (!belongs) return []
  switch (ev.type) {
    case 'message':
    case 'text':
      if (ev.text) return [{ type: 'delta', text: ev.text }]
      if (ev.message) return [{ type: 'delta', text: ev.message }]
      return []
    case 'tool_call':
      if (ev.toolCall) {
        const tc: ToolCall = {
          id: ev.toolCall.id ?? `call_${Math.random().toString(36).slice(2, 10)}`,
          name: ev.toolCall.name ?? '',
          arguments: ev.toolCall.args ?? '{}',
        }
        return [{ type: 'tool_call', toolCall: tc }]
      }
      return []
    case 'usage':
      if (ev.usage) {
        return [
          {
            type: 'usage',
            promptTokens: ev.usage.promptTokens ?? 0,
            completionTokens: ev.usage.completionTokens ?? 0,
            totalTokens: (ev.usage.promptTokens ?? 0) + (ev.usage.completionTokens ?? 0),
          },
        ]
      }
      return []
    case 'done':
    case 'end':
      return [{ type: 'done', requestId: sessionId, finishReason: 'stop' }]
    default:
      logger.debug({ msg: 'opencode event unmapped', type: ev.type })
      return []
  }
}

export type { ChatMessage }
