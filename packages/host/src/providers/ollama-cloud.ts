import type { ModelInfo, ProviderCapabilities } from '@openofficellm/shared'
import { providerApiKey } from './credentials.js'
import { OpenAiCompatibleAdapter } from './openai-compatible.js'
import type { ProviderAdapter } from './types.js'

// `api.ollama.com` 301s to the apex domain and the redirect does not land on
// the API, so every request through it failed as an unparseable HTML body.
// The OpenAI-compatible surface is on the apex: https://ollama.com/v1.
const DEFAULT_BASE = 'https://ollama.com/v1'
const CAPS: ProviderCapabilities = { tools: true, vision: false, streaming: true }

export interface OllamaCloudConfig {
  baseUrl?: string
  id?: string
  name?: string
  modelsHint?: string[]
}

/**
 * Ollama Cloud exposes an OpenAI-compatible `/v1/chat/completions` endpoint
 * (SSE streaming), so we delegate to `OpenAiCompatibleAdapter` rather than
 * re-implementing the SSE/tool-call parsing. The only Ollama-Cloud-specific
 * behaviour is the auth header (bearer token from the secret store) and the
 * model hint fallback used when `/v1/models` is unreachable.
 */
export class OllamaCloudAdapter implements ProviderAdapter {
  readonly id: string
  readonly name: string
  readonly kind = 'cloud' as const
  readonly capabilities: ProviderCapabilities = CAPS
  private inner: OpenAiCompatibleAdapter

  constructor(cfg: OllamaCloudConfig = {}) {
    const id = cfg.id ?? 'ollama-cloud'
    this.id = id
    this.name = cfg.name ?? 'Ollama Cloud'
    this.inner = new OpenAiCompatibleAdapter({
      id,
      name: this.name,
      baseUrl: (cfg.baseUrl ?? DEFAULT_BASE).replace(/\/+$/, ''),
      kind: 'cloud',
      capabilities: CAPS,
      modelsHint: cfg.modelsHint,
    })
  }

  isConfigured(): boolean {
    return providerApiKey(this.id) !== null
  }

  async isReachable(): Promise<boolean> {
    return this.inner.isReachable()
  }

  async listModels(): Promise<ModelInfo[]> {
    return this.inner.listModels()
  }

  async *stream(req: Parameters<ProviderAdapter['stream']>[0], signal: AbortSignal) {
    yield* this.inner.stream(req, signal)
  }
}
