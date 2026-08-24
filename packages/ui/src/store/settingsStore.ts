// Settings store: load/save Settings via the host; defaults per host.

import { create } from 'zustand'
import {
  DEFAULT_SETTINGS,
  migrateStepCap,
  type EditMode,
  type HostKind,
  type Settings,
} from '@openofficellm/shared'
import { getSettings, putSettings } from '../api/client'
import { loadPersisted, savePersisted } from './persist'

export interface SettingsState {
  settings: Settings
  loading: boolean
  saving: boolean
  error: string | null
  load: () => Promise<void>
  save: (next: Settings) => Promise<Settings | null>
  setDefaultModel: (modelId: string) => Promise<void>
  setDefaultMode: (host: HostKind, mode: EditMode) => void
}

export function mergeDefaults(loaded: Partial<Settings>): Settings {
  return {
    ...DEFAULT_SETTINGS,
    ...loaded,
    defaultMode: { ...DEFAULT_SETTINGS.defaultMode, ...loaded.defaultMode },
    defaultContext: { ...DEFAULT_SETTINGS.defaultContext, ...loaded.defaultContext },
    agenticStepCap: migrateStepCap(loaded.agenticStepCap),
    // A localStorage copy written before the v2 config migration still holds
    // the old object shape here; anything not an array is discarded rather than
    // handed to a UI that will call .map on it.
    mcpServers: Array.isArray(loaded.mcpServers) ? loaded.mcpServers : DEFAULT_SETTINGS.mcpServers,
    mcpToolConsent: loaded.mcpToolConsent ?? DEFAULT_SETTINGS.mcpToolConsent,
    agents: Array.isArray(loaded.agents) ? loaded.agents : DEFAULT_SETTINGS.agents,
    importedProviders: Array.isArray(loaded.importedProviders)
      ? loaded.importedProviders
      : DEFAULT_SETTINGS.importedProviders,
    providerOptions:
      loaded.providerOptions && typeof loaded.providerOptions === 'object'
        ? loaded.providerOptions
        : DEFAULT_SETTINGS.providerOptions,
    disabledSkills: Array.isArray(loaded.disabledSkills)
      ? loaded.disabledSkills
      : DEFAULT_SETTINGS.disabledSkills,
    showImportedSkills:
      typeof loaded.showImportedSkills === 'boolean'
        ? loaded.showImportedSkills
        : DEFAULT_SETTINGS.showImportedSkills,
  }
}

export const useSettingsStore = create<SettingsState>((set, get) => ({
  settings: mergeDefaults(loadPersisted<Partial<Settings>>('settings.last', {})),
  loading: false,
  saving: false,
  error: null,
  load: async () => {
    set({ loading: true, error: null })
    try {
      const s = await getSettings()
      const merged = mergeDefaults(s)
      savePersisted('settings.last', merged)
      set({ settings: merged, loading: false })
    } catch (e) {
      set({ loading: false, error: (e as Error).message })
    }
  },
  save: async (next) => {
    set({ saving: true, error: null })
    try {
      const saved = await putSettings(next)
      const merged = mergeDefaults(saved)
      savePersisted('settings.last', merged)
      set({ settings: merged, saving: false })
      return merged
    } catch (e) {
      set({ saving: false, error: (e as Error).message })
      return null
    }
  },
  setDefaultModel: async (modelId) => {
    const next = { ...get().settings, defaultModel: modelId }
    const saved = await get().save(next)
    if (saved) savePersisted('models.selected', modelId)
  },
  setDefaultMode: (host, mode) => {
    const next = {
      ...get().settings,
      defaultMode: { ...get().settings.defaultMode, [host]: mode },
    }
    set({ settings: next })
    savePersisted('settings.last', next)
    void get().save(next)
  },
}))
