// Typed HTTP client for the host service.
//
// The Office pane is served *by* the host, so it talks to it same-origin and
// gets its bearer token from a <meta> tag the host injects into the HTML. A
// browser extension is served from its own `chrome-extension://` origin and so
// has neither: it must be told where the host is and what token to present.
// Hence `configureApi` — call it before the first request, or leave it alone
// and the same-origin defaults apply.

import type {
  ApiError,
  DetectedHost,
  HealthResponse,
  McpCallResponse,
  McpServerConfig,
  McpServerInfo,
  ModelInfo,
  ProviderInfo,
  Settings,
  Skill,
  UpdateApplyResponse,
} from '@openofficellm/shared'

export interface ModelsResponse {
  models: ModelInfo[]
}

export interface ProvidersResponse {
  providers: ProviderInfo[]
}

export interface ProviderKeyResponse {
  ok: boolean
  configured: boolean
}

export interface ProviderTestResponse {
  ok: boolean
  reachable: boolean
  modelCount?: number
  error?: string
}

export interface CancelResponse {
  cancelled: boolean
}

/** Origin of the host service. Empty means same-origin, which is what the
 *  Office pane wants: relative paths follow whatever port the host bound. */
let _base = ''

/** Extra headers attached to every request. The extension uses this to name
 *  itself to the host, which is how the host recognises it when Chrome elects
 *  not to send an Origin on a privileged fetch. */
let _extraHeaders: Record<string, string> = {}

/** Set by a shell that cannot use the same-origin defaults. */
export function configureApi(opts: {
  baseUrl?: string
  token?: string
  headers?: Record<string, string>
}): void {
  if (opts.baseUrl !== undefined) _base = opts.baseUrl.replace(/\/+$/, '')
  if (opts.token !== undefined) _token = opts.token
  if (opts.headers !== undefined) _extraHeaders = { ...opts.headers }
}

/** Headers every request carries, whether or not it also carries a body. */
export function baseHeaders(): Record<string, string> {
  return { ..._extraHeaders }
}

/** Absolute URL for an API path, or the path itself when running same-origin. */
export function apiUrl(path: string): string {
  return _base ? `${_base}${path}` : path
}

/**
 * `import.meta.env`, without requiring the bundler's ambient types.
 *
 * This package is consumed as source by two shells with different tsconfigs,
 * and a `vite/client` reference that one of them fails to pick up is a type
 * error in a file neither of them owns. Reading it through a cast keeps the
 * package self-contained; Vite still statically replaces the expression at
 * build time, and outside Vite the whole thing is simply absent.
 */
function viteEnv(): Record<string, unknown> {
  const meta = import.meta as unknown as { env?: Record<string, unknown> }
  return meta.env ?? {}
}

let _token: string | null | undefined = undefined

function readToken(): string | null {
  if (_token !== undefined) return _token
  if (typeof document !== 'undefined') {
    const meta = document.querySelector('meta[name="auth-token"]')
    const content = meta?.getAttribute('content')
    if (content && content.length > 0) {
      _token = content
      return content
    }
  }
  // The query-string fallback is a dev convenience only. The production pane is
  // served same-origin by the host, which injects the token into the meta tag.
  // Shipping the query fallback would persist the token in browser history,
  // the referrer header, and WebView2's navigation log. Gate it behind DEV and
  // strip it from the URL immediately after read.
  if (viteEnv().DEV === true && typeof window !== 'undefined') {
    const params = new URLSearchParams(window.location.search)
    const q = params.get('token')
    if (q) {
      _token = q
      try {
        const url = new URL(window.location.href)
        url.searchParams.delete('token')
        window.history.replaceState(null, '', url.toString())
      } catch {
        // ignore — best-effort scrub
      }
      return q
    }
  }
  const env = viteEnv().VITE_AUTH_TOKEN as string | undefined
  if (env) {
    _token = env
    return env
  }
  _token = null
  return null
}

export function getAuthToken(): string | null {
  return readToken()
}

function authHeaders(): Record<string, string> {
  const token = readToken()
  const h: Record<string, string> = {
    ..._extraHeaders,
    'Content-Type': 'application/json',
  }
  if (token) h['Authorization'] = `Bearer ${token}`
  return h
}

async function parseError(res: Response): Promise<ApiError> {
  let body: unknown
  try {
    body = await res.json()
  } catch {
    return { code: `http_${res.status}`, message: res.statusText || 'network error' }
  }
  const e = body as Partial<ApiError>
  return {
    code: e.code ?? `http_${res.status}`,
    message: e.message ?? res.statusText ?? 'unknown error',
    retryable: e.retryable,
  }
}

async function jsonRequest<T>(
  method: string,
  path: string,
  body?: unknown,
  init?: RequestInit,
): Promise<T> {
  let res: Response
  try {
    res = await fetch(apiUrl(path), {
      method,
      headers: authHeaders(),
      body: body !== undefined ? JSON.stringify(body) : undefined,
      ...init,
    })
  } catch (e) {
    const err: ApiError = {
      code: 'network',
      message: (e as Error).message ?? 'network error',
      retryable: true,
    }
    throw err
  }
  if (res.status === 401 || res.status === 403) {
    const err = await parseError(res)
    err.code = err.code === `http_${res.status}` ? 'forbidden' : err.code
    err.message = 'Authentication failed. Relaunch the add-in from the ribbon.'
    throw err
  }
  if (!res.ok) throw await parseError(res)
  return (await res.json()) as T
}

export async function getHealth(): Promise<HealthResponse> {
  const res = await fetch(apiUrl('/api/health'), {
    method: 'GET',
    headers: { ..._extraHeaders, 'Cache-Control': 'no-cache' },
  })
  if (!res.ok) throw await parseError(res)
  return (await res.json()) as HealthResponse
}

export async function applyUpdate(): Promise<UpdateApplyResponse> {
  return jsonRequest<UpdateApplyResponse>('POST', '/api/update/apply')
}

export async function skipUpdate(version: string): Promise<void> {
  await jsonRequest<void>('POST', '/api/update/skip', { version })
}

export function getProviders(): Promise<ProvidersResponse> {
  return jsonRequest<ProvidersResponse>('GET', '/api/providers')
}

export function getModels(refresh = false): Promise<ModelsResponse> {
  const path = refresh ? '/api/models?refresh=1' : '/api/models'
  return jsonRequest<ModelsResponse>('GET', path)
}

export function getSettings(): Promise<Settings> {
  return jsonRequest<Settings>('GET', '/api/settings')
}

export function putSettings(settings: Settings): Promise<Settings> {
  return jsonRequest<Settings>('PUT', '/api/settings', settings)
}

export function putProviderKey(id: string, key: string): Promise<ProviderKeyResponse> {
  return jsonRequest<ProviderKeyResponse>('PUT', `/api/providers/${encodeURIComponent(id)}/key`, {
    key,
  })
}

export function deleteProviderKey(id: string): Promise<ProviderKeyResponse> {
  return jsonRequest<ProviderKeyResponse>('DELETE', `/api/providers/${encodeURIComponent(id)}/key`)
}

export function testProvider(id: string): Promise<ProviderTestResponse> {
  return jsonRequest<ProviderTestResponse>('POST', `/api/providers/${encodeURIComponent(id)}/test`)
}

export function cancelChat(requestId: string): Promise<CancelResponse> {
  return jsonRequest<CancelResponse>('POST', '/api/chat/cancel', { requestId })
}

// ─── Skills ──────────────────────────────────────────────────────────────

export interface SkillsResponse {
  skills: Skill[]
  /** Where user skills live, shown in settings so the user can drop a file in. */
  userDir: string
}

export function getSkills(host?: DetectedHost): Promise<SkillsResponse> {
  const q = host && host !== 'none' ? `?host=${host}` : ''
  return jsonRequest<SkillsResponse>('GET', `/api/skills${q}`)
}

export type SkillDraft = Omit<Skill, 'builtIn' | 'source' | 'path'>

export function putSkill(skill: SkillDraft): Promise<{ ok: boolean }> {
  const { id, ...body } = skill
  return jsonRequest<{ ok: boolean }>('PUT', `/api/skills/${encodeURIComponent(id)}`, body)
}

export function deleteSkill(id: string): Promise<{ ok: boolean }> {
  return jsonRequest<{ ok: boolean }>('DELETE', `/api/skills/${encodeURIComponent(id)}`)
}

export interface OpencodeImportResult {
  ok: boolean
  providerCount: number
  mcpCount: number
  agentCount: number
  providers: { id: string; name: string; baseUrl: string; modelCount: number }[]
  mcpServers: { id: string; name: string; transport: 'stdio' | 'http' }[]
  agents: { id: string; name: string; model: string }[]
  /** Providers this host already had an adapter for, which gained a key. */
  linkedProviders: string[]
  /** Config and credential files actually read. */
  sources: string[]
  /** Directories searched, shown when nothing was found. */
  searched: string[]
  errors: string[]
}

const EMPTY_IMPORT: Omit<OpencodeImportResult, 'ok' | 'errors'> = {
  providerCount: 0,
  mcpCount: 0,
  agentCount: 0,
  providers: [],
  mcpServers: [],
  agents: [],
  linkedProviders: [],
  sources: [],
  searched: [],
}

export function emptyImportResult(message: string): OpencodeImportResult {
  return { ...EMPTY_IMPORT, ok: false, errors: [message] }
}

export function previewOpencodeImport(): Promise<OpencodeImportResult> {
  return jsonRequest<OpencodeImportResult>('GET', '/api/opencode/import')
}

export function runOpencodeImport(): Promise<OpencodeImportResult> {
  return jsonRequest<OpencodeImportResult>('POST', '/api/opencode/import')
}

// ─── MCP ─────────────────────────────────────────────────────────────────

export interface McpServersResponse {
  servers: McpServerInfo[]
}

export function getMcpServers(refresh = false): Promise<McpServersResponse> {
  return jsonRequest<McpServersResponse>('GET', `/api/mcp/servers${refresh ? '?refresh=1' : ''}`)
}

export function putMcpServer(server: McpServerConfig): Promise<McpServersResponse> {
  return jsonRequest<McpServersResponse>('POST', '/api/mcp/servers', server)
}

export function deleteMcpServer(id: string): Promise<McpServersResponse> {
  return jsonRequest<McpServersResponse>('DELETE', `/api/mcp/servers/${encodeURIComponent(id)}`)
}

export function setMcpConsent(
  serverId: string,
  tool: string,
  enabled: boolean,
): Promise<McpServersResponse> {
  return jsonRequest<McpServersResponse>('POST', '/api/mcp/consent', { serverId, tool, enabled })
}

export function callMcpTool(tool: string, args: Record<string, unknown>): Promise<McpCallResponse> {
  return jsonRequest<McpCallResponse>('POST', '/api/mcp/call', { tool, arguments: args })
}

export class ClientError extends Error {
  code: string
  retryable?: boolean
  constructor(api: ApiError) {
    super(api.message)
    this.name = 'ClientError'
    this.code = api.code
    this.retryable = api.retryable
  }
}
