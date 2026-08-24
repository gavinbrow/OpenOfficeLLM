import fs from 'node:fs'
import path from 'node:path'
import { z } from 'zod'
import {
  DEFAULT_SETTINGS,
  migrateStepCap,
  type Settings,
  type EditMode,
  type ContextScope,
  type ProviderOptions,
} from '@openofficellm/shared'
import { resolveConfigPath, ensureDirs, DEFAULT_PORT } from './paths.js'
import { logger } from './logging.js'

export const CURRENT_CONFIG_VERSION = 3

const editModeSchema = z.enum(['propose', 'direct', 'agentic'])
const contextScopeSchema = z.enum([
  'none',
  'selection',
  'paragraph',
  'document',
  'sheet',
  'range',
  'page',
])

/** Server ids become part of a tool name (`mcp__<id>__<tool>`), so the
 *  separator must stay unambiguous — no underscores, no whitespace. */
export const mcpServerSchema = z.object({
  id: z
    .string()
    .min(1)
    .max(64)
    .regex(/^[a-z0-9][a-z0-9-]*$/, 'id must be lowercase alphanumeric with dashes'),
  name: z.string().min(1).max(120),
  transport: z.enum(['stdio', 'http']),
  command: z.string().max(1024).optional(),
  args: z.array(z.string().max(1024)).max(64).optional(),
  env: z.record(z.string(), z.string()).optional(),
  url: z.string().url().max(2048).optional(),
  headers: z.record(z.string(), z.string()).optional(),
  enabled: z.boolean(),
  /** Set by the opencode importer. Without it in the schema zod strips the
   *  field on save, and the settings UI loses the ability to say where a
   *  server came from. */
  imported: z.boolean().optional(),
})

export const settingsSchema = z.object({
  defaultModel: z.string().nullable().optional(),
  defaultMode: z.object({
    word: editModeSchema,
    excel: editModeSchema,
    browser: editModeSchema,
  }),
  defaultContext: z.object({
    word: contextScopeSchema,
    excel: contextScopeSchema,
    browser: contextScopeSchema,
  }),
  theme: z.enum(['light', 'dark']).nullable().optional(),
  contextTrimWarningTokens: z.number().int().min(100),
  agenticStepCap: z.number().int().min(1).max(100),
  showReasoning: z.boolean(),
  mcpServers: z.array(mcpServerSchema).max(64),
  mcpToolConsent: z.record(z.string(), z.record(z.string(), z.boolean())),
  mcpConfirmEveryCall: z.boolean(),
  agents: z.array(
    z.object({
      id: z.string().min(1).max(64),
      name: z.string().min(1).max(120),
      description: z.string().max(500).default(''),
      model: z.string().min(1).max(200),
      mode: editModeSchema.optional(),
      tools: z.boolean().optional(),
      permissions: z.record(z.string(), z.string()).optional(),
      source: z.enum(['opencode', 'user']),
      enabled: z.boolean(),
    }),
  ),
  importedProviders: z.array(
    z.object({
      id: z.string().min(1).max(64),
      name: z.string().min(1).max(120),
      baseUrl: z.string().url().max(2048),
      models: z.array(z.string()).optional(),
      enabled: z.boolean(),
    }),
  ),
  providerOptions: z.object({
    ollamaBaseUrl: z.string().url().max(2048).optional(),
  }),
  disabledSkills: z.array(z.string()),
  showImportedSkills: z.boolean(),
  // Capped: this is an allowlist, and an unbounded one is a place for junk to
  // accumulate unnoticed. Nobody legitimately pairs sixteen extensions.
  pairedExtensions: z.array(z.string()).max(16),
})

export const configSchema = z.object({
  version: z.number().int().min(1),
  port: z.number().int().min(1).max(65535),
  settings: settingsSchema,
})

export type Config = z.infer<typeof configSchema>

/** The settings shape as parsed, which differs from the shared `Settings` type
 *  in one place: `defaultModel` is nullable here, because a config file written
 *  by hand may say `null` and rejecting the whole file over it would be
 *  gratuitous. Callers that read config and write it back should use this. */
export type StoredSettings = Config['settings']

export function defaultConfig(port: number): Config {
  return {
    version: CURRENT_CONFIG_VERSION,
    port: port > 0 ? port : DEFAULT_PORT,
    settings: { ...DEFAULT_SETTINGS } as Settings,
  }
}

type Migration = (raw: unknown) => unknown
const migrations: Record<number, Migration> = {
  // v1 → v2: `mcpServers` went from a consent map keyed by id to a list of
  // server definitions, with consent split into its own field. v1 never shipped
  // a UI to add a server, so the old map only ever held the empty object — but
  // migrating rather than discarding keeps any hand-written consent intact.
  2: (raw) => {
    const obj = (raw ?? {}) as Record<string, unknown>
    const settings = (obj.settings ?? {}) as Record<string, unknown>
    const legacy = settings.mcpServers
    const consent: Record<string, Record<string, boolean>> = {}
    if (legacy && typeof legacy === 'object' && !Array.isArray(legacy)) {
      for (const [id, entry] of Object.entries(legacy as Record<string, unknown>)) {
        const e = entry as { tools?: Record<string, boolean> }
        if (e && typeof e === 'object' && e.tools) consent[id] = e.tools
      }
    }
    return {
      ...obj,
      version: 2,
      settings: {
        ...settings,
        mcpServers: [],
        mcpToolConsent: consent,
        showReasoning: settings.showReasoning ?? true,
        mcpConfirmEveryCall: settings.mcpConfirmEveryCall ?? false,
      },
    }
  },
  // v2 → v3: add opencode-import fields (agents, imported providers,
  // provider options, skill toggles).
  3: (raw) => {
    const obj = (raw ?? {}) as Record<string, unknown>
    const settings = (obj.settings ?? {}) as Record<string, unknown>
    return {
      ...obj,
      version: 3,
      settings: {
        ...settings,
        agents: Array.isArray(settings.agents) ? settings.agents : [],
        importedProviders: Array.isArray(settings.importedProviders)
          ? settings.importedProviders
          : [],
        providerOptions:
          settings.providerOptions && typeof settings.providerOptions === 'object'
            ? settings.providerOptions
            : {},
        disabledSkills: Array.isArray(settings.disabledSkills) ? settings.disabledSkills : [],
        showImportedSkills:
          typeof settings.showImportedSkills === 'boolean' ? settings.showImportedSkills : true,
      },
    }
  },
}

function migrate(raw: unknown): unknown {
  if (raw === null || typeof raw !== 'object') return raw
  const obj = raw as Record<string, unknown>
  let version = typeof obj.version === 'number' ? obj.version : 0
  let current: unknown = raw
  while (version < CURRENT_CONFIG_VERSION) {
    const m = migrations[version + 1]
    if (!m) break
    current = m(current)
    version += 1
  }
  return current
}

function normalizeSettings(raw: unknown): Settings {
  if (raw === null || typeof raw !== 'object') return { ...DEFAULT_SETTINGS }
  const obj = raw as Partial<Settings>
  const def = DEFAULT_SETTINGS
  const merge = <T>(base: T, patch: Partial<T> | undefined): T =>
    patch ? { ...base, ...patch } : base
  const settings: Settings = {
    defaultModel: obj.defaultModel ?? def.defaultModel,
    defaultMode: merge(def.defaultMode, obj.defaultMode),
    defaultContext: merge(def.defaultContext, obj.defaultContext),
    theme: obj.theme ?? def.theme,
    contextTrimWarningTokens:
      typeof obj.contextTrimWarningTokens === 'number'
        ? obj.contextTrimWarningTokens
        : def.contextTrimWarningTokens,
    agenticStepCap: migrateStepCap(obj.agenticStepCap),
    showReasoning: typeof obj.showReasoning === 'boolean' ? obj.showReasoning : def.showReasoning,
    // A pre-migration config could still hold the v1 object shape here; coerce
    // rather than let it reach the schema and fail the whole load.
    mcpServers: Array.isArray(obj.mcpServers) ? obj.mcpServers : def.mcpServers,
    mcpToolConsent:
      obj.mcpToolConsent && typeof obj.mcpToolConsent === 'object'
        ? obj.mcpToolConsent
        : def.mcpToolConsent,
    mcpConfirmEveryCall:
      typeof obj.mcpConfirmEveryCall === 'boolean'
        ? obj.mcpConfirmEveryCall
        : def.mcpConfirmEveryCall,
    agents: Array.isArray(obj.agents) ? obj.agents : def.agents,
    importedProviders: Array.isArray(obj.importedProviders)
      ? obj.importedProviders
      : def.importedProviders,
    providerOptions:
      obj.providerOptions && typeof obj.providerOptions === 'object'
        ? (obj.providerOptions as ProviderOptions)
        : def.providerOptions,
    disabledSkills: Array.isArray(obj.disabledSkills) ? obj.disabledSkills : def.disabledSkills,
    showImportedSkills:
      typeof obj.showImportedSkills === 'boolean' ? obj.showImportedSkills : def.showImportedSkills,
    // Anything that is not a list of strings is treated as no pairings at all.
    // Failing closed is the only safe direction for an allowlist.
    pairedExtensions: Array.isArray(obj.pairedExtensions)
      ? obj.pairedExtensions.filter((o): o is string => typeof o === 'string')
      : def.pairedExtensions,
  }
  return settings
}

export function loadConfig(): Config {
  ensureDirs()
  let raw: unknown
  try {
    const text = fs.readFileSync(resolveConfigPath(), 'utf8')
    raw = JSON.parse(text)
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code === 'ENOENT') {
      const fresh = defaultConfig(0)
      saveConfig(fresh)
      return fresh
    }
    logger.warn({ msg: 'config read failed, using defaults', error: String(e) })
    return defaultConfig(0)
  }
  const migrated = migrate(raw)
  const parsed = configSchema.safeParse({
    ...(migrated as Record<string, unknown>),
    settings: normalizeSettings((migrated as { settings?: unknown }).settings),
  })
  if (!parsed.success) {
    logger.warn({ msg: 'config validation failed, using defaults', error: parsed.error.message })
    return defaultConfig(0)
  }
  return parsed.data
}

export function saveConfig(cfg: Config): void {
  ensureDirs()
  const dir = path.dirname(resolveConfigPath())
  const tmp = path.join(dir, `.config.${process.pid}.tmp`)
  const data = JSON.stringify(cfg, null, 2)
  fs.writeFileSync(tmp, data, { encoding: 'utf8' })
  try {
    fs.renameSync(tmp, resolveConfigPath())
  } catch (e) {
    try {
      fs.copyFileSync(tmp, resolveConfigPath())
      fs.unlinkSync(tmp)
    } catch {
      try {
        fs.unlinkSync(tmp)
      } catch {
        // ignore
      }
      throw e
    }
  }
}

export function updateSettings(settings: StoredSettings): Config {
  const cfg = loadConfig()
  const next: Config = { ...cfg, settings }
  const parsed = configSchema.parse(next)
  saveConfig(parsed)
  return parsed
}

export function setConfigPort(port: number): Config {
  const cfg = loadConfig()
  const next: Config = { ...cfg, port }
  saveConfig(next)
  return next
}

export type { Settings, EditMode, ContextScope }
