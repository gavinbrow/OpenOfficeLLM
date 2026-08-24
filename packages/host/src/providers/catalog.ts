// Provider identity catalog.
//
// Two things live here, both needed by the opencode importer:
//
//  1. Which provider ids this host already ships an adapter for. If opencode
//     has credentials for `anthropic`, the right move is to hand the key to the
//     Anthropic adapter we already register — not to invent a second,
//     OpenAI-shaped `anthropic` provider that would shadow it.
//
//  2. Base URLs for providers opencode knows about but does not write a
//     `baseURL` for. opencode resolves those from models.dev at runtime, so a
//     config entry can be nothing but a model list — which is exactly how
//     `ollama-cloud` appears, and why the importer used to skip it with
//     "has no baseURL".

/** Provider ids for which this host registers its own adapter at startup. */
export const BUILTIN_PROVIDER_IDS: ReadonlySet<string> = new Set([
  'ollama',
  'ollama-cloud',
  'lm-studio',
  'llama-cpp',
  'opencode',
  'openai',
  'anthropic',
  'google',
  'openrouter',
  'groq',
  'deepseek',
  'together',
  'xai',
])

/** opencode provider ids that mean the same thing as one of ours under a
 *  different name. Aliasing lets an imported API key land on the adapter that
 *  can actually use it. */
const PROVIDER_ALIASES: Record<string, string> = {
  togetherai: 'together',
  'together-ai': 'together',
  'google-generative-ai': 'google',
  'google-ai-studio': 'google',
  gemini: 'google',
  'x-ai': 'xai',
  grok: 'xai',
  lmstudio: 'lm-studio',
  lm_studio: 'lm-studio',
  llamacpp: 'llama-cpp',
  'ollama-turbo': 'ollama-cloud',
  ollamacloud: 'ollama-cloud',
}

export function canonicalProviderId(id: string): string {
  return PROVIDER_ALIASES[id] ?? id
}

export interface KnownProvider {
  name: string
  /** OpenAI-compatible base URL, including the version segment. */
  baseUrl: string
}

/** Base URLs for well-known providers, used when the opencode entry omits one.
 *  Only OpenAI-compatible endpoints belong here — anything else has to go
 *  through a dedicated adapter. */
export const KNOWN_PROVIDERS: Record<string, KnownProvider> = {
  // Verified against the live service: api.ollama.com 301s to the apex, and
  // following that redirect lands on the marketing site rather than the API.
  'ollama-cloud': { name: 'Ollama Cloud', baseUrl: 'https://ollama.com/v1' },
  openai: { name: 'OpenAI', baseUrl: 'https://api.openai.com/v1' },
  openrouter: { name: 'OpenRouter', baseUrl: 'https://openrouter.ai/api/v1' },
  groq: { name: 'Groq', baseUrl: 'https://api.groq.com/openai/v1' },
  deepseek: { name: 'DeepSeek', baseUrl: 'https://api.deepseek.com/v1' },
  together: { name: 'Together', baseUrl: 'https://api.together.xyz/v1' },
  xai: { name: 'xAI', baseUrl: 'https://api.x.ai/v1' },
  mistral: { name: 'Mistral', baseUrl: 'https://api.mistral.ai/v1' },
  cerebras: { name: 'Cerebras', baseUrl: 'https://api.cerebras.ai/v1' },
  fireworks: { name: 'Fireworks', baseUrl: 'https://api.fireworks.ai/inference/v1' },
  'fireworks-ai': { name: 'Fireworks', baseUrl: 'https://api.fireworks.ai/inference/v1' },
  perplexity: { name: 'Perplexity', baseUrl: 'https://api.perplexity.ai' },
  nebius: { name: 'Nebius', baseUrl: 'https://api.studio.nebius.ai/v1' },
  moonshotai: { name: 'Moonshot', baseUrl: 'https://api.moonshot.ai/v1' },
  zhipuai: { name: 'Z.ai', baseUrl: 'https://api.z.ai/api/paas/v4' },
  venice: { name: 'Venice', baseUrl: 'https://api.venice.ai/api/v1' },
  ollama: { name: 'Ollama', baseUrl: 'http://127.0.0.1:11434/v1' },
}

/** Providers whose auth we cannot meaningfully reuse: OAuth-only, or a wire
 *  protocol no adapter here speaks. */
export const UNSUPPORTED_PROVIDER_IDS: ReadonlySet<string> = new Set([
  'github-copilot',
  'amazon-bedrock',
  'azure',
  'google-vertex',
  'google-vertex-anthropic',
])
