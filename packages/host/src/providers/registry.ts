import type { ModelInfo } from '@openofficellm/shared'
import type { ProviderAdapter } from './types.js'
import { logger } from '../logging.js'

const registry = new Map<string, ProviderAdapter>()

export function registerProvider(adapter: ProviderAdapter): void {
  if (registry.has(adapter.id)) {
    logger.debug({ msg: 'provider already registered, replacing', id: adapter.id })
  }
  registry.set(adapter.id, adapter)
}

export function unregisterProvider(id: string): boolean {
  return registry.delete(id)
}

export function getProvider(id: string): ProviderAdapter | undefined {
  return registry.get(id)
}

export function allProviders(): ProviderAdapter[] {
  return Array.from(registry.values())
}

export function clearProviders(): void {
  registry.clear()
}

export function dedupeModels(models: ModelInfo[]): ModelInfo[] {
  const seen = new Set<string>()
  const out: ModelInfo[] = []
  for (const m of models) {
    if (seen.has(m.id)) continue
    seen.add(m.id)
    out.push(m)
  }
  return out
}

export function findProviderForModel(modelId: string): ProviderAdapter | undefined {
  const providerId = modelId.includes('/') ? modelId.split('/')[0] : modelId
  return registry.get(providerId)
}

export function providerIds(): string[] {
  return Array.from(registry.keys())
}
