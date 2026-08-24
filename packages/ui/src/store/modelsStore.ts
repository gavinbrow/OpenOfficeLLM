// Models store: list of ModelInfo, selected model id, refresh, provider info,
// plus the user's per-model visibility and favourite choices.
//
// Visibility and favourites live here rather than in Settings because they are
// a property of this machine's model list, not of the assistant's behaviour: a
// model id that means something on one host is meaningless on another, and the
// settings document is round-tripped through the host's schema on every save.

import { create } from 'zustand'
import type { ModelInfo, ProviderInfo } from '@openofficellm/shared'
import { getModels, getProviders } from '../api/client'
import { loadPersisted, savePersisted } from './persist'

const SELECTED_KEY = 'models.selected'
const HIDDEN_KEY = 'models.hidden'
const FAVORITES_KEY = 'models.favorites'
/** Favourites were "pinned" before the picker grew provider groups. */
const LEGACY_PINNED_KEY = 'models.pinned'

export interface ModelsState {
  models: ModelInfo[]
  providers: ProviderInfo[]
  selectedModelId: string | null
  /** Model ids the user has switched off; still listed in Settings. */
  hidden: string[]
  favorites: string[]
  loading: boolean
  error: string | null
  lastFetch: number
  selectModel: (id: string) => void
  load: (refresh?: boolean) => Promise<void>
  loadProviders: () => Promise<void>
  /** Show or hide one or many models at once (a whole provider, typically). */
  setHidden: (ids: string[], hidden: boolean) => void
  toggleFavorite: (id: string) => void
  /** Models the picker should offer: usable, minus the ones switched off. */
  visibleModels: () => ModelInfo[]
}

function isStringArray(v: unknown): v is string[] {
  return Array.isArray(v) && v.every((x) => typeof x === 'string')
}

/** Hide cloud-provider models that are not reachable or have no key configured.
 *  Local providers are shown whenever they are reachable. */
function filterUsableModels(models: ModelInfo[], providers: ProviderInfo[]): ModelInfo[] {
  const infoById = new Map(providers.map((p) => [p.id, p]))
  return models.filter((m) => {
    const info = infoById.get(m.providerId)
    if (!info) return false
    if (info.kind === 'local') return info.reachable
    return info.reachable && info.configured
  })
}

/**
 * Choose the model to run with.
 *
 * Keeps the current pick when it is still visible; otherwise prefers a
 * favourite, then anything visible. Returning to a hidden model would leave the
 * composer naming a model absent from its own picker.
 */
function pickSelection(
  models: ModelInfo[],
  hidden: string[],
  favorites: string[],
  current: string | null,
): string | null {
  const hiddenSet = new Set(hidden)
  const visible = models.filter((m) => !hiddenSet.has(m.id))
  if (current && visible.some((m) => m.id === current)) return current
  const favorite = visible.find((m) => favorites.includes(m.id))
  return favorite?.id ?? visible[0]?.id ?? null
}

export const useModelsStore = create<ModelsState>((set, get) => ({
  models: [],
  providers: [],
  selectedModelId: loadPersisted<string | null>(SELECTED_KEY, null),
  hidden: loadPersisted<string[]>(HIDDEN_KEY, [], isStringArray),
  favorites: loadPersisted<string[]>(
    FAVORITES_KEY,
    loadPersisted<string[]>(LEGACY_PINNED_KEY, [], isStringArray),
    isStringArray,
  ),
  loading: false,
  error: null,
  lastFetch: 0,

  selectModel: (id) => {
    savePersisted(SELECTED_KEY, id)
    set({ selectedModelId: id })
  },

  load: async (refresh = false) => {
    set({ loading: true, error: null })
    try {
      // Both calls matter. `filterUsableModels` drops every model whose
      // provider is unknown, so swallowing a provider-list failure turned an
      // auth or network error into a silent, and wrong, "no models found".
      const [modelsRes, providersRes] = await Promise.all([getModels(refresh), getProviders()])
      const providers = providersRes.providers
      const models = filterUsableModels(modelsRes.models, providers)
      const { hidden, favorites, selectedModelId } = get()
      const next = pickSelection(models, hidden, favorites, selectedModelId)
      if (next !== selectedModelId) savePersisted(SELECTED_KEY, next)
      set({
        models,
        providers,
        loading: false,
        lastFetch: Date.now(),
        selectedModelId: next,
      })
    } catch (e) {
      set({ loading: false, error: (e as Error).message })
    }
  },

  loadProviders: async () => {
    try {
      const res = await getProviders()
      set({ providers: res.providers })
    } catch {
      // providers are best-effort; models still work
    }
  },

  setHidden: (ids, hide) => {
    const { hidden, models, favorites, selectedModelId } = get()
    const target = new Set(ids)
    const next = hide
      ? [...hidden.filter((id) => !target.has(id)), ...ids]
      : hidden.filter((id) => !target.has(id))
    savePersisted(HIDDEN_KEY, next)
    const selection = pickSelection(models, next, favorites, selectedModelId)
    if (selection !== selectedModelId) savePersisted(SELECTED_KEY, selection)
    set({ hidden: next, selectedModelId: selection })
  },

  toggleFavorite: (id) => {
    const { favorites } = get()
    const next = favorites.includes(id) ? favorites.filter((f) => f !== id) : [...favorites, id]
    savePersisted(FAVORITES_KEY, next)
    set({ favorites: next })
  },

  visibleModels: () => {
    const { models, hidden } = get()
    if (hidden.length === 0) return models
    const hiddenSet = new Set(hidden)
    return models.filter((m) => !hiddenSet.has(m.id))
  },
}))
