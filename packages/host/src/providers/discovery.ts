import { OllamaAdapter } from './ollama.js'
import { OpenAiCompatibleAdapter } from './openai-compatible.js'
import { OpencodeAdapter } from './opencode.js'
import type { ProviderAdapter } from './types.js'
import { registerProvider } from './registry.js'
import { logger } from '../logging.js'

interface Probe {
  id: string
  url: string
  factory: () => ProviderAdapter
}

const PROBES: Probe[] = [
  {
    id: 'ollama',
    url: process.env.OLLAMA_HOST
      ? `${process.env.OLLAMA_HOST.replace(/\/+$/, '')}/api/tags`
      : 'http://127.0.0.1:11434/api/tags',
    factory: () => new OllamaAdapter(),
  },
  {
    id: 'lm-studio',
    url: 'http://127.0.0.1:1234/v1/models',
    factory: () =>
      new OpenAiCompatibleAdapter({
        id: 'lm-studio',
        name: 'LM Studio',
        baseUrl: 'http://127.0.0.1:1234/v1',
        kind: 'local',
        authHeaderStyle: 'none',
      }),
  },
  {
    id: 'opencode',
    url: 'http://127.0.0.1:4096/health',
    factory: () => new OpencodeAdapter(),
  },
  {
    id: 'llama-cpp',
    url: 'http://127.0.0.1:8080/v1/models',
    factory: () =>
      new OpenAiCompatibleAdapter({
        id: 'llama-cpp',
        name: 'llama.cpp',
        baseUrl: 'http://127.0.0.1:8080/v1',
        kind: 'local',
        authHeaderStyle: 'none',
      }),
  },
]

const PROBE_TIMEOUT_MS = 500
const DISCOVERY_INTERVAL_MS = 60_000

async function probe(url: string): Promise<boolean> {
  try {
    const ctrl = new AbortController()
    const timer = setTimeout(() => ctrl.abort(), PROBE_TIMEOUT_MS)
    try {
      const res = await fetch(url, { method: 'GET', signal: ctrl.signal })
      return res.ok
    } finally {
      clearTimeout(timer)
    }
  } catch {
    return false
  }
}

export async function discoverProviders(): Promise<ProviderAdapter[]> {
  const found: ProviderAdapter[] = []
  const results = await Promise.all(
    PROBES.map(async (p) => ({ id: p.id, ok: await probe(p.url), factory: p.factory })),
  )
  for (const r of results) {
    if (r.ok) {
      const adapter = r.factory()
      registerProvider(adapter)
      found.push(adapter)
      logger.info({ msg: 'discovered provider', id: r.id })
    }
  }
  return found
}

let loopTimer: NodeJS.Timeout | null = null

export function startDiscoveryLoop(): void {
  if (loopTimer) return
  loopTimer = setInterval(() => {
    void discoverProviders().catch((e) => {
      logger.warn({ msg: 'discovery loop error', error: String(e) })
    })
  }, DISCOVERY_INTERVAL_MS)
  loopTimer.unref()
}

export function stopDiscoveryLoop(): void {
  if (loopTimer) {
    clearInterval(loopTimer)
    loopTimer = null
  }
}

export function probeList(): string[] {
  return PROBES.map((p) => p.id)
}
