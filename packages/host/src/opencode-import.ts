// Import providers, MCP servers and agents from a local opencode install.
//
// The first version of this looked in exactly one place for exactly one thing:
// `~/.config/opencode/opencode.json(c)`, and inside it, providers that spelled
// out a `baseURL`. That is not how opencode configs are actually shaped. Two
// gaps made the import come back near-empty on a real install:
//
//   * Credentials do not live in the config at all. `opencode auth login`
//     writes them to `auth.json` under the *data* directory, so a provider the
//     user is signed into looks, from the config's point of view, like nothing.
//
//   * A provider entry frequently has no `baseURL`. opencode resolves those
//     from models.dev at runtime. `ollama-cloud` is the common case: the config
//     carries a model list and a price override, and nothing else.
//
// So: search every plausible config location, union the config's providers with
// whatever is in auth.json, and fill missing base URLs from a local catalog.
// Providers this host already has a first-class adapter for get the key handed
// to that adapter rather than being re-created as generic OpenAI-compatible.

import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import type {
  AgentConfig,
  EditMode,
  ImportedProviderConfig,
  McpServerConfig,
} from '@openofficellm/shared'
import { setSecret, listConfigured } from './secrets.js'
import { logger } from './logging.js'
import {
  BUILTIN_PROVIDER_IDS,
  KNOWN_PROVIDERS,
  UNSUPPORTED_PROVIDER_IDS,
  canonicalProviderId,
} from './providers/catalog.js'

export interface OpencodeImportResult {
  ok: boolean
  providerCount: number
  mcpCount: number
  agentCount: number
  providers: ImportedProviderConfig[]
  mcpServers: McpServerConfig[]
  agents: AgentConfig[]
  /** Keys stored against an adapter this host already ships (Anthropic, Google,
   *  OpenAI…). Not "imported providers", but very much part of the import. */
  linkedProviders: string[]
  /** Files actually read, so the UI can say where the data came from. */
  sources: string[]
  /** Directories searched, for the "found nothing" message. */
  searched: string[]
  errors: string[]
}

interface OpencodeProviderEntry {
  name?: string
  npm?: string
  options?: {
    baseURL?: string
    baseUrl?: string
    apiKey?: string
    apikey?: string
    api_key?: string
  }
  models?: Record<string, { name?: string } | null>
  enabled?: boolean
}

interface OpencodeAgentEntry {
  description?: string
  /** opencode's own notion of mode: primary | subagent | all. Unrelated to our
   *  EditMode, which is why it is not read as one. */
  mode?: string
  model?: string
  disable?: boolean
  temperature?: number
  permission?: Record<string, unknown>
  tools?: Record<string, boolean>
}

interface OpencodeMcpEntry {
  type?: 'remote' | 'local' | 'stdio'
  url?: string
  command?: string | string[]
  args?: string[]
  enabled?: boolean
  environment?: Record<string, string>
  env?: Record<string, string>
  headers?: Record<string, string>
}

interface OpencodeConfig {
  provider?: Record<string, OpencodeProviderEntry>
  agent?: Record<string, OpencodeAgentEntry>
  mcp?: Record<string, OpencodeMcpEntry>
  disabled_providers?: string[]
}

/** `{ "<provider>": { "type": "api", "key": "…" } }`, written by
 *  `opencode auth login`. OAuth entries carry tokens instead of a key. */
type OpencodeAuth = Record<string, { type?: string; key?: string } | undefined>

const CONFIG_FILENAMES = ['opencode.jsonc', 'opencode.json', 'config.json']

/** Every directory opencode might keep its config in, most specific first.
 *
 *  `OPENCODE_CONFIG_DIR` is an override, not an addition: when it is set it is
 *  the only directory searched, matching opencode's own reading of the
 *  variable. `XDG_CONFIG_HOME` likewise *replaces* `~/.config` rather than
 *  sitting alongside it. */
export function configDirCandidates(): string[] {
  const explicit = process.env.OPENCODE_CONFIG_DIR
  if (explicit) return [explicit]

  const home = os.homedir()
  const dirs: string[] = []
  const push = (d: string | undefined) => {
    if (d && !dirs.includes(d)) dirs.push(d)
  }
  push(
    process.env.XDG_CONFIG_HOME
      ? path.join(process.env.XDG_CONFIG_HOME, 'opencode')
      : path.join(home, '.config', 'opencode'),
  )
  if (process.env.APPDATA) push(path.join(process.env.APPDATA, 'opencode'))
  if (process.env.LOCALAPPDATA) push(path.join(process.env.LOCALAPPDATA, 'opencode'))
  push(path.join(home, '.opencode'))
  return dirs
}

/** Where `auth.json` lives — the *data* directory, which is not the config
 *  directory on any platform. Same override rules as above. */
export function dataDirCandidates(): string[] {
  const explicit = process.env.OPENCODE_DATA_DIR
  if (explicit) return [explicit]

  const home = os.homedir()
  const dirs: string[] = []
  const push = (d: string | undefined) => {
    if (d && !dirs.includes(d)) dirs.push(d)
  }
  push(
    process.env.XDG_DATA_HOME
      ? path.join(process.env.XDG_DATA_HOME, 'opencode')
      : path.join(home, '.local', 'share', 'opencode'),
  )
  if (process.env.LOCALAPPDATA) push(path.join(process.env.LOCALAPPDATA, 'opencode'))
  if (process.env.APPDATA) push(path.join(process.env.APPDATA, 'opencode'))
  push(path.join(home, '.opencode'))
  return dirs
}

interface ReadResult {
  config: OpencodeConfig
  auth: OpencodeAuth
  sources: string[]
  searched: string[]
  errors: string[]
  found: boolean
}

function readAll(): ReadResult {
  const sources: string[] = []
  const searched: string[] = []
  const errors: string[] = []
  let config: OpencodeConfig = {}
  let found = false

  // Merge across config locations rather than stopping at the first hit: a
  // machine can legitimately have a global config and a second one under
  // OPENCODE_CONFIG_DIR, and taking only one silently drops half the providers.
  for (const dir of configDirCandidates()) {
    searched.push(dir)
    for (const name of CONFIG_FILENAMES) {
      const p = path.join(dir, name)
      if (!fs.existsSync(p)) continue
      try {
        const text = fs.readFileSync(p, 'utf8')
        const parsed = JSON.parse(stripJsonComments(text)) as OpencodeConfig
        config = mergeConfigs(config, parsed)
        sources.push(p)
        found = true
      } catch (e) {
        logger.warn({ msg: 'failed to parse opencode config', path: p, error: String(e) })
        errors.push(`Could not parse ${name} in ${dir}.`)
      }
      // One config file per directory; the filenames are alternatives.
      break
    }
  }

  let auth: OpencodeAuth = {}
  for (const dir of dataDirCandidates()) {
    const p = path.join(dir, 'auth.json')
    if (!fs.existsSync(p)) continue
    try {
      const parsed = JSON.parse(fs.readFileSync(p, 'utf8')) as OpencodeAuth
      auth = { ...parsed, ...auth }
      sources.push(p)
      found = true
    } catch (e) {
      logger.warn({ msg: 'failed to parse opencode auth', path: p, error: String(e) })
      errors.push(`Could not parse auth.json in ${dir}.`)
    }
  }

  return { config, auth, sources, searched, errors, found }
}

/** Shallow per-section merge; earlier configs win so the most specific
 *  directory (OPENCODE_CONFIG_DIR) keeps precedence. */
function mergeConfigs(base: OpencodeConfig, next: OpencodeConfig): OpencodeConfig {
  return {
    provider: { ...next.provider, ...base.provider },
    agent: { ...next.agent, ...base.agent },
    mcp: { ...next.mcp, ...base.mcp },
    disabled_providers: [...(next.disabled_providers ?? []), ...(base.disabled_providers ?? [])],
  }
}

function toId(raw: string): string {
  return raw
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
}

export function stripJsonComments(text: string): string {
  // Strips `//` line comments, `/* */` block comments and trailing commas —
  // all three are legal in a .jsonc and all three make JSON.parse throw. String
  // literals are tracked so a URL like "https://x//y" survives intact.
  let out = ''
  let inString = false
  let escape = false
  let i = 0
  while (i < text.length) {
    const ch = text[i]
    if (inString) {
      out += ch
      if (escape) escape = false
      else if (ch === '\\') escape = true
      else if (ch === '"') inString = false
      i++
      continue
    }
    if (ch === '"') {
      inString = true
      out += ch
      i++
      continue
    }
    if (ch === '/' && text[i + 1] === '/') {
      while (i < text.length && text[i] !== '\n') i++
      continue
    }
    if (ch === '/' && text[i + 1] === '*') {
      i += 2
      while (i < text.length && !(text[i] === '*' && text[i + 1] === '/')) i++
      i += 2
      continue
    }
    out += ch
    i++
  }
  // Trailing commas, now that comments are gone and cannot hide one.
  return out.replace(/,(\s*[}\]])/g, '$1')
}

function providerKey(entry: OpencodeProviderEntry | undefined): string | undefined {
  const k = entry?.options?.apiKey ?? entry?.options?.apikey ?? entry?.options?.api_key
  return typeof k === 'string' && k.trim() ? k.trim() : undefined
}

function providerBaseUrl(entry: OpencodeProviderEntry | undefined): string | undefined {
  const u = entry?.options?.baseURL ?? entry?.options?.baseUrl
  return typeof u === 'string' && u.trim() ? u.trim().replace(/\/+$/, '') : undefined
}

/** Model *ids*, taken from the record keys. The `name` field beside them is a
 *  display label — sending it upstream as the model id produces a 404, which is
 *  what the previous version did. */
function modelIds(entry: OpencodeProviderEntry | undefined): string[] {
  if (!entry?.models) return []
  return Object.keys(entry.models).filter((k) => k.trim().length > 0)
}

interface ImportedProviders {
  providers: ImportedProviderConfig[]
  linked: string[]
  errors: string[]
}

function importProviders(
  cfg: OpencodeConfig,
  auth: OpencodeAuth,
  storeKeys: boolean,
): ImportedProviders {
  const providers: ImportedProviderConfig[] = []
  const linked: string[] = []
  const errors: string[] = []
  const disabled = new Set((cfg.disabled_providers ?? []).map((d) => canonicalProviderId(toId(d))))

  // The union is the point: a provider can appear only in the config (a custom
  // endpoint with an inline key), only in auth.json (signed in via the CLI, no
  // config entry at all), or in both.
  const rawIds = new Set<string>([...Object.keys(cfg.provider ?? {}), ...Object.keys(auth ?? {})])

  const seen = new Set<string>()
  for (const rawId of rawIds) {
    const id = canonicalProviderId(toId(rawId))
    if (!id || seen.has(id)) continue
    seen.add(id)

    if (disabled.has(id)) continue
    if (UNSUPPORTED_PROVIDER_IDS.has(id)) {
      errors.push(`Skipped "${rawId}" — its sign-in cannot be reused outside opencode.`)
      continue
    }

    const entry = cfg.provider?.[rawId]
    if (entry?.enabled === false) continue

    const authEntry = auth[rawId] ?? auth[id]
    const key =
      providerKey(entry) ??
      (authEntry && (authEntry.type === undefined || authEntry.type === 'api')
        ? authEntry.key
        : undefined)

    if (key && storeKeys) {
      try {
        setSecret(id, key)
      } catch (e) {
        logger.warn({ msg: 'failed to store imported provider key', id, error: String(e) })
        errors.push(`Could not store the API key for "${rawId}".`)
      }
    } else if (!key && authEntry && authEntry.type && authEntry.type !== 'api') {
      errors.push(`"${rawId}" is signed in with ${authEntry.type}, which cannot be re-used here.`)
    }

    // Providers this host already ships an adapter for: the key is the whole
    // import. Registering a second, generic provider under the same id would
    // shadow the real one and lose its native protocol (Anthropic and Google
    // are not OpenAI-shaped).
    const configuredBaseUrl = providerBaseUrl(entry)
    if (BUILTIN_PROVIDER_IDS.has(id) && !configuredBaseUrl) {
      if (key) linked.push(id)
      continue
    }

    const known = KNOWN_PROVIDERS[id]
    const baseUrl = configuredBaseUrl ?? known?.baseUrl
    if (!baseUrl) {
      errors.push(
        `Provider "${rawId}" has no base URL in the opencode config and is not a provider this host knows; skipped.`,
      )
      continue
    }

    providers.push({
      id,
      name: entry?.name ?? known?.name ?? rawId,
      baseUrl,
      models: modelIds(entry),
      enabled: true,
    })
  }

  providers.sort((a, b) => a.id.localeCompare(b.id))
  return { providers, linked: [...new Set(linked)].sort(), errors }
}

function importMcp(
  cfg: OpencodeConfig,
  maxServers: number,
): { servers: McpServerConfig[]; errors: string[] } {
  const servers: McpServerConfig[] = []
  const errors: string[] = []
  if (!cfg.mcp) return { servers, errors }
  const seen = new Set<string>()
  for (const [rawId, entry] of Object.entries(cfg.mcp)) {
    if (servers.length >= maxServers) {
      errors.push(`MCP server limit (${maxServers}) reached; remaining servers skipped.`)
      break
    }
    const id = toId(rawId)
    if (!id || seen.has(id)) {
      errors.push(`MCP server "${rawId}" has no usable id; skipped.`)
      continue
    }
    seen.add(id)
    const transport = entry.type === 'remote' ? 'http' : 'stdio'
    if (transport === 'http') {
      if (!entry.url) {
        errors.push(`MCP server "${rawId}" has no URL; skipped.`)
        continue
      }
      servers.push({
        id,
        name: rawId,
        transport: 'http',
        url: entry.url,
        headers: entry.headers,
        enabled: false, // user opts in per-server, and per-tool, in Settings
        imported: true,
      })
      continue
    }
    // opencode allows `command` to be the whole argv array.
    const argv = Array.isArray(entry.command)
      ? entry.command
      : entry.command
        ? [entry.command, ...(entry.args ?? [])]
        : []
    if (argv.length === 0) {
      errors.push(`MCP server "${rawId}" has no command; skipped.`)
      continue
    }
    servers.push({
      id,
      name: rawId,
      transport: 'stdio',
      command: argv[0],
      args: argv.slice(1),
      env: entry.environment ?? entry.env,
      enabled: false,
      imported: true,
    })
  }
  return { servers, errors }
}

/** opencode's `mode` on an agent is primary/subagent/all — a routing hint, not
 *  one of our edit modes. Treating it as one produced a spurious
 *  "invalid mode subagent" warning on every import. */
const OPENCODE_AGENT_MODES = new Set(['primary', 'subagent', 'all'])
const EDIT_MODES: EditMode[] = ['propose', 'direct', 'agentic']

function importAgents(cfg: OpencodeConfig): { agents: AgentConfig[]; errors: string[] } {
  const agents: AgentConfig[] = []
  const errors: string[] = []
  if (!cfg.agent) return { agents, errors }
  const seen = new Set<string>()
  for (const [rawId, entry] of Object.entries(cfg.agent)) {
    const id = toId(rawId)
    if (!id || seen.has(id)) continue
    seen.add(id)
    if (entry.disable) continue
    if (!entry.model) {
      errors.push(`Agent "${rawId}" has no model; skipped.`)
      continue
    }
    const mode = entry.mode
    const editMode = mode && EDIT_MODES.includes(mode as EditMode) ? (mode as EditMode) : undefined
    if (mode && !editMode && !OPENCODE_AGENT_MODES.has(mode)) {
      errors.push(`Agent "${rawId}" has an unrecognised mode "${mode}"; using the default.`)
    }
    // opencode writes `provider/model`, which is already our wire format — but
    // the provider half may be an alias of one of ours.
    const slash = entry.model.indexOf('/')
    const model =
      slash > 0
        ? `${canonicalProviderId(toId(entry.model.slice(0, slash)))}/${entry.model.slice(slash + 1)}`
        : entry.model

    agents.push({
      id,
      name: rawId,
      description: (entry.description ?? '').slice(0, 500),
      model,
      mode: editMode,
      tools: true,
      permissions: normalizePermissions(entry.permission),
      source: 'opencode',
      enabled: false, // user opts in in Settings
    })
  }
  return { agents, errors }
}

/** The config schema stores permissions as string→string. opencode allows
 *  nested objects (e.g. `bash: { "git *": "allow" }`); flatten anything that is
 *  not already a plain string so the whole import does not fail validation. */
function normalizePermissions(
  raw: Record<string, unknown> | undefined,
): Record<string, string> | undefined {
  if (!raw || typeof raw !== 'object') return undefined
  const out: Record<string, string> = {}
  for (const [k, v] of Object.entries(raw)) {
    if (typeof v === 'string') out[k] = v
    else if (v && typeof v === 'object') out[k] = 'conditional'
  }
  return Object.keys(out).length > 0 ? out : undefined
}

const MAX_MCP_SERVERS = 64

export function importFromOpencode(commit = true): OpencodeImportResult {
  const { config, auth, sources, searched, errors: readErrors, found } = readAll()

  if (!found) {
    return {
      ok: false,
      providerCount: 0,
      mcpCount: 0,
      agentCount: 0,
      providers: [],
      mcpServers: [],
      agents: [],
      linkedProviders: [],
      sources: [],
      searched,
      errors: [
        `No opencode config or credentials found. Looked in: ${searched.join(', ')}.`,
        ...readErrors,
      ],
    }
  }

  const { providers, linked, errors: providerErrors } = importProviders(config, auth, commit)
  const { servers, errors: mcpErrors } = importMcp(config, MAX_MCP_SERVERS)
  const { agents, errors: agentErrors } = importAgents(config)
  const errors = [...readErrors, ...providerErrors, ...mcpErrors, ...agentErrors]

  logger.info({
    msg: commit ? 'opencode import complete' : 'opencode preview',
    sources: sources.length,
    providers: providers.length,
    linked: linked.length,
    mcpServers: servers.length,
    agents: agents.length,
    errors: errors.length,
  })

  return {
    // A config that parsed is a successful read even if every entry in it was
    // unusable — the errors list is how that gets reported, not a 404.
    ok: true,
    providerCount: providers.length,
    mcpCount: servers.length,
    agentCount: agents.length,
    providers,
    mcpServers: servers,
    agents,
    linkedProviders: linked,
    sources,
    searched,
    errors,
  }
}

/**
 * Re-import API keys from opencode's auth.json for any provider that is no
 * longer in the secrets store.
 *
 * The full import (`importFromOpencode`) is a user-initiated action — it
 * merges providers, MCP servers, and agents into config, which is a
 * destructive operation. But keys disappearing after an update or a
 * secrets.dat corruption is a common silent failure: the host starts, the
 * provider shows "no key", and the user has to re-import or re-enter keys
 * by hand. This function does the one safe thing that fixes that: read
 * auth.json, and for any provider that has a key there but not in the
 * secrets store, store it. No config changes, no provider registration —
 * just keys.
 *
 * Returns the list of provider ids whose keys were restored, for logging.
 */
export function syncKeysFromOpencode(): string[] {
  const existing = new Set(listConfigured())
  const restored: string[] = []

  // Read auth.json directly — we don't need the full config, just the keys.
  let auth: OpencodeAuth = {}
  for (const dir of dataDirCandidates()) {
    const p = path.join(dir, 'auth.json')
    if (!fs.existsSync(p)) continue
    try {
      const parsed = JSON.parse(fs.readFileSync(p, 'utf8')) as OpencodeAuth
      auth = { ...auth, ...parsed }
    } catch {
      // ignore — a corrupt auth.json is not something we can fix
    }
  }

  for (const [rawId, entry] of Object.entries(auth)) {
    if (!entry || entry.type !== undefined && entry.type !== 'api') continue
    if (!entry.key) continue
    const id = canonicalProviderId(toId(rawId))
    if (!id) continue
    if (existing.has(id)) continue

    try {
      setSecret(id, entry.key)
      restored.push(id)
    } catch (e) {
      logger.warn({ msg: 'failed to restore provider key from opencode', id, error: String(e) })
    }
  }

  if (restored.length > 0) {
    logger.info({ msg: 'restored provider keys from opencode auth.json', providers: restored })
  }

  return restored
}
