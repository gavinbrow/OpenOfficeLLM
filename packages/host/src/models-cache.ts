import type { ModelInfo } from '@openofficellm/shared'
import { allProviders, dedupeModels, findProviderForModel } from './providers/registry.js'
import { logger } from './logging.js'

const CACHE_TTL_MS = 60_000

interface CacheEntry {
  at: number
  models: ModelInfo[]
}

let cache: CacheEntry | null = null
// In-flight refresh promise — coalesces concurrent callers so two requests
// don't both fire Promise.allSettled and race to overwrite the cache.
let refreshPromise: Promise<ModelInfo[]> | null = null

export function bustModelCache(): void {
  cache = null
}

export async function listAllModels(forceRefresh = false): Promise<ModelInfo[]> {
  if (!forceRefresh && cache && Date.now() - cache.at < CACHE_TTL_MS) {
    return cache.models
  }
  // Coalesce concurrent refreshes: if one is already running, wait on it
  // rather than firing a second parallel fetch.
  if (refreshPromise) return refreshPromise
  refreshPromise = doRefresh()
  try {
    return await refreshPromise
  } finally {
    refreshPromise = null
  }
}

async function doRefresh(): Promise<ModelInfo[]> {
  const providers = allProviders()
  const results = await Promise.allSettled(
    providers.map(async (p: ReturnType<typeof allProviders>[number]) => {
      try {
        return await p.listModels()
      } catch (e) {
        logger.warn({ msg: 'model list failed for provider', id: p.id, error: String(e) })
        return [] as ModelInfo[]
      }
    }),
  )
  const all: ModelInfo[] = []
  for (const r of results) {
    if (r.status === 'fulfilled') all.push(...r.value)
  }
  const deduped = dedupeModels(all)
  // Don't clobber a good cache with an empty list from a transient blip —
  // keep the previous models and just refresh the timestamp so the UI
  // doesn't show "no models" for 60s after a momentary network drop.
  if (deduped.length === 0 && cache && cache.models.length > 0) {
    cache = { at: Date.now(), models: cache.models }
    return cache.models
  }
  cache = { at: Date.now(), models: deduped }
  return deduped
}

export function resolveProviderForModel(modelId: string) {
  return findProviderForModel(modelId)
}
