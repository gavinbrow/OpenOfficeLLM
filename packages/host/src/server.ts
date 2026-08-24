import path from 'node:path'
import http from 'node:http'
import https from 'node:https'
import { Hono } from 'hono'
import { streamSSE } from 'hono/streaming'
import type { Context } from 'hono'
import { serve } from '@hono/node-server'
import { DEFAULT_PORT, HOST_INTERFACE, HOST_VERSION, ensureDirs } from './paths.js'
import { z } from 'zod'
import {
  loadConfig,
  updateSettings,
  setConfigPort,
  settingsSchema,
  mcpServerSchema,
} from './config.js'
import { selectPort } from './port.js'
import { createAuthToken, checkToken } from './auth.js'
import { isPairedExtension } from './pairing.js'
import { loadCertMaterial } from './tls.js'
import { staticDirExists, serveStatic, warnIfMissing } from './static.js'
import { ensureOfficeJsCache, getCachedOfficeJs } from './officejs.js'
import { writeManifest } from './manifest.js'
import { logger } from './logging.js'
import type { CliOptions } from './index.js'
import { listConfigured, setSecret, deleteSecret, isFallbackMode } from './secrets.js'
import { registerProvider, allProviders, getProvider } from './providers/registry.js'
import { OllamaAdapter } from './providers/ollama.js'
import { OllamaCloudAdapter } from './providers/ollama-cloud.js'
import { presetAdapters, createImportedAdapter } from './providers/openai-compatible.js'
import { AnthropicAdapter } from './providers/anthropic.js'
import { GoogleAdapter } from './providers/google.js'
import { discoverProviders, startDiscoveryLoop, stopDiscoveryLoop } from './providers/discovery.js'
import { listAllModels, bustModelCache, resolveProviderForModel } from './models-cache.js'
import { newRequest, cancelRequest, completeRequest, wrapStream } from './chat-sessions.js'
import { buildSystemPrompt, withSystemPrompt } from './prompt.js'
import { splitReasoning } from './providers/reasoning.js'
import {
  getSkills,
  skillsForHost,
  findSkill,
  saveUserSkill,
  deleteUserSkill,
} from './skills/loader.js'
import {
  syncServers,
  stopAllServers,
  listServers,
  enabledToolDefinitions,
  callTool as callMcpTool,
} from './mcp/registry.js'
import { importFromOpencode } from './opencode-import.js'
import {
  type ChatRequest,
  type HealthResponse,
  type HostKind,
  type ProviderInfo,
  type Skill,
  type StreamEvent,
  type ToolDefinition,
} from '@openofficellm/shared'
import { ProviderError } from './providers/types.js'

export interface StartServerResult {
  port: number
  close: () => Promise<void>
}

const startTime = Date.now()

/** Body accepted by PUT /api/skills/:id. The id is taken from the path so a
 *  body cannot name a different file than the URL it was sent to. */
const skillWriteSchema = z.object({
  name: z.string().min(1).max(120),
  description: z.string().max(500).default(''),
  hosts: z
    .array(z.enum(['word', 'excel', 'browser']))
    .max(3)
    .default([]),
  mode: z.enum(['propose', 'direct', 'agentic']).optional(),
  model: z.string().max(200).optional(),
  icon: z.string().max(8).optional(),
  contextScope: z
    .enum(['none', 'selection', 'paragraph', 'document', 'sheet', 'range', 'page'])
    .optional(),
  prompt: z.string().min(1).max(20_000),
})

function findAddinDist(): string {
  // process.execPath is the SEA binary when packaged, or node.exe in dev. In
  // both cases the directory it lives in is the reliable anchor for finding
  // the web bundle — the installer places it at <install-dir>/web/, and in
  // dev the monorepo layout puts it near packages/addin/dist.
  //
  // fileURLToPath(import.meta.url) is NOT used because the CJS bundle (for
  // the SEA binary) has no import.meta — tsup shims it to an empty object.
  const exeDir = path.dirname(process.execPath)
  const cwd = process.cwd()
  const candidates = [
    // SEA / installer layout: <install-dir>/web/ sits next to host.exe.
    path.resolve(exeDir, 'web'),
    // Dev monorepo: run from the repo root, add-in is at packages/addin/dist.
    path.resolve(cwd, 'packages/addin/dist'),
    // Dev: run from packages/host, add-in is two levels up.
    path.resolve(cwd, '../addin/dist'),
  ]
  for (const c of candidates) {
    if (staticDirExists(c)) return c
  }
  // Fall back to the installer layout rather than the dev path — in a packaged
  // install this is where the files genuinely are, and a clear error path beats
  // a wrong directory that happens to exist.
  return path.resolve(exeDir, 'web')
}

/** Narrow a `?host=` query value to a real host. Anything else means "do not
 *  filter", which is the right answer for a caller that did not say. */
function isHostKind(v: string | undefined): v is HostKind {
  return v === 'word' || v === 'excel' || v === 'browser'
}

function ownOrigin(port: number): string {
  return `https://${HOST_INTERFACE}:${port}`
}

/** Extension origins the user has paired, read fresh per request so pairing
 *  takes effect without a restart. */
function pairedOrigins(): string[] {
  try {
    return loadConfig().settings.pairedExtensions
  } catch {
    // A config that will not load is not a reason to trust everyone.
    return []
  }
}

/** Whether this Origin may talk to the API at all, and on what basis.
 *
 *  Two callers are legitimate: the task pane, which the host serves itself and
 *  which is therefore same-origin, and a paired browser extension, which is
 *  cross-origin by construction and has to be named explicitly. Everything
 *  else — a web page that discovered the port, another extension — is refused
 *  here regardless of what token it presents. */
function classifyOrigin(c: Context, port: number): 'same-origin' | 'paired-extension' | 'rejected' {
  const origin = c.req.header('origin') ?? ''
  const method = c.req.method.toUpperCase()
  const paired = pairedOrigins()

  if (origin === ownOrigin(port)) return 'same-origin'
  if (isPairedExtension(origin, paired)) return 'paired-extension'

  // Chrome does not always attach an Origin to a fetch made from an extension
  // page that holds host permission for the target — it treats the request as
  // privileged rather than cross-origin. The extension therefore also names
  // itself in a custom header, which is a sound CSRF signal in its own right:
  // a web page cannot set a custom header on a cross-origin request without a
  // preflight, and the OPTIONS handler below refuses every origin that is not
  // already paired.
  const declared = c.req.header('x-openofficellm-extension') ?? ''
  if (declared !== '' && isPairedExtension(declared, paired)) {
    // If an Origin *is* present it has to be the same extension. Otherwise a
    // paired extension's id could be borrowed by anything that learned it.
    if (origin === '' || origin === declared) return 'paired-extension'
    return 'rejected'
  }

  // Same-origin GETs sometimes omit Origin; same-origin POST/PUT/DELETE always
  // include it per the fetch spec. Allow an empty Origin only for GETs. A
  // cross-origin request from a web page can never suppress the header, so
  // this does not widen anything: it only tolerates the pane's own reads.
  if (origin === '' && method === 'GET') return 'same-origin'
  return 'rejected'
}

/** CORS headers for a paired extension. Only ever sent to an origin that has
 *  already been matched against the allowlist — never echoed blindly. */
function applyCors(c: Context, origin: string): void {
  c.header('Access-Control-Allow-Origin', origin)
  c.header(
    'Access-Control-Allow-Headers',
    'Authorization, Content-Type, Cache-Control, X-OpenOfficeLLM-Extension',
  )
  c.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS')
  // The token travels in an Authorization header, not a cookie, so credentials
  // stay off: allowing them would let any page that can reach this port ride
  // along on ambient state instead of proving it holds the token.
  c.header('Vary', 'Origin')
}

function authMiddleware(token: string, port: number) {
  return async (c: Context, next: () => Promise<void>) => {
    // The origin to echo back in CORS headers. When Chrome omitted Origin and
    // the extension named itself in the custom header instead, that header is
    // the origin — echoing an empty string would produce a useless
    // `Access-Control-Allow-Origin:` and break the very request it allowed.
    const origin = c.req.header('origin') || (c.req.header('x-openofficellm-extension') ?? '')
    const method = c.req.method.toUpperCase()
    const kind = classifyOrigin(c, port)

    // Preflight carries no Authorization header by definition, so it is
    // answered on the origin check alone and never reaches a route.
    if (method === 'OPTIONS') {
      if (kind !== 'paired-extension') {
        return c.json({ code: 'forbidden', message: 'forbidden' }, 403)
      }
      applyCors(c, origin)
      return c.body(null, 204)
    }

    if (kind === 'paired-extension') applyCors(c, origin)

    if (c.req.path === '/api/health') {
      // Health is deliberately unauthenticated — it is how the pane discovers
      // whether the service is up before it has a token — but it still answers
      // only to origins that are allowed to be here at all.
      if (kind === 'rejected') {
        return c.json({ code: 'forbidden', message: 'forbidden' }, 403)
      }
      await next()
      return
    }

    const authOk = checkToken(c.req.header('authorization'), token)
    if (kind === 'rejected' || !authOk) {
      return c.json({ code: 'forbidden', message: 'forbidden' }, 403)
    }
    await next()
  }
}

export async function startServer(
  opts: CliOptions,
  hooks: { onShutdown?: () => void } = {},
): Promise<() => Promise<void>> {
  ensureDirs()
  const cfg = loadConfig()
  const portSelection = await selectPort(opts.port || cfg.port || DEFAULT_PORT)
  if (portSelection.port !== cfg.port) {
    setConfigPort(portSelection.port)
  }
  const port = portSelection.port

  const auth = createAuthToken()
  const addinDist = findAddinDist()
  warnIfMissing(addinDist)

  logger.info({ msg: 'ensuring office.js cache' })
  const officeJs = await ensureOfficeJsCache()
  logger.info({ msg: 'office.js ready', source: officeJs.meta.source })

  const cert = loadCertMaterial()
  if (cert.renewed) {
    logger.info({ msg: 'leaf cert was renewed on startup' })
  }

  const ollamaAdapter = new OllamaAdapter()
  if (cfg.settings.providerOptions.ollamaBaseUrl) {
    ollamaAdapter.setBaseUrl(cfg.settings.providerOptions.ollamaBaseUrl)
  }
  registerProvider(ollamaAdapter)
  for (const adapter of presetAdapters()) {
    registerProvider(adapter)
  }
  registerProvider(new AnthropicAdapter())
  registerProvider(new GoogleAdapter())
  // Ollama Cloud registers unconditionally, like every other cloud adapter. It
  // used to register only when it appeared in `importedProviders`, which meant
  // that importing an Ollama Cloud *key* from opencode — the normal case, since
  // opencode keeps credentials in auth.json and the config entry carries no
  // base URL — left the provider invisible. An adapter with no key simply shows
  // as "no key" in Settings; that is the honest state, and it gives the user
  // somewhere to paste one.
  if (process.env.OLLAMA_CLOUD_API_KEY) {
    setSecret('ollama-cloud', process.env.OLLAMA_CLOUD_API_KEY)
  }
  const ollamaCloudImported = cfg.settings.importedProviders.find((p) => p.id === 'ollama-cloud')
  registerProvider(
    new OllamaCloudAdapter({
      baseUrl: ollamaCloudImported?.baseUrl,
      modelsHint: ollamaCloudImported?.models,
    }),
  )
  // Register imported OpenAI-compatible providers.
  for (const p of cfg.settings.importedProviders) {
    if (p.id === 'ollama-cloud') continue
    registerProvider(createImportedAdapter(p))
  }
  await discoverProviders()
  startDiscoveryLoop()

  // Best-effort: a broken MCP server must not stop the host from serving the
  // pane, so failures are logged and the server is marked errored in the UI.
  void syncServers().catch((e) =>
    logger.warn({ msg: 'initial mcp sync failed', error: String((e as Error).message ?? e) }),
  )

  const app = new Hono()
  app.use('/api/*', authMiddleware(auth.token, port))

  app.onError((err, c) => {
    logger.error({ msg: 'unhandled error', error: String(err), stack: (err as Error).stack })
    return c.json({ code: 'error', message: 'internal error' }, 500)
  })

  app.get('/api/health', (c) => {
    const resp: HealthResponse = {
      status: 'ok',
      version: HOST_VERSION,
      port,
      uptimeSeconds: Math.floor((Date.now() - startTime) / 1000),
    }
    return c.json(resp)
  })

  app.get('/api/settings', (c) => {
    const current = loadConfig()
    return c.json(current.settings)
  })

  app.put('/api/settings', async (c) => {
    let body: unknown
    try {
      body = await c.req.json()
    } catch {
      return c.json({ code: 'bad_request', message: 'invalid JSON' }, 400)
    }
    const parsed = settingsSchema.safeParse(body)
    if (!parsed.success) {
      return c.json(
        { code: 'bad_request', message: 'validation failed', errors: parsed.error.message },
        400,
      )
    }
    const next = updateSettings(parsed.data)
    // Settings carry the MCP server list, so a save can add, remove, enable, or
    // disable a server. Reconcile rather than waiting for the next request.
    void syncServers().catch((e) =>
      logger.warn({ msg: 'mcp sync failed after settings save', error: String(e) }),
    )
    return c.json(next.settings)
  })

  app.get('/api/providers', async (c) => {
    const providers = allProviders()
    const configured = new Set(listConfigured())
    const infos: ProviderInfo[] = await Promise.all(
      providers.map(async (p) => ({
        id: p.id,
        name: p.name,
        kind: p.kind,
        reachable: await p.isReachable().catch(() => false),
        configured: p.kind === 'local' ? true : configured.has(p.id),
        capabilities: p.capabilities,
      })),
    )
    infos.sort((a, b) => a.id.localeCompare(b.id))
    return c.json({ providers: infos })
  })

  app.get('/api/models', async (c) => {
    const refresh = c.req.query('refresh') === '1'
    if (refresh) bustModelCache()
    const models = await listAllModels(refresh)
    return c.json({ models })
  })

  app.post('/api/chat', (c) => {
    return streamSSE(c, async (stream) => {
      let reqBody: ChatRequest
      try {
        reqBody = (await c.req.json()) as ChatRequest
      } catch {
        const ev: StreamEvent = { type: 'error', code: 'bad_request', message: 'invalid JSON' }
        await stream.writeSSE({ event: 'error', data: JSON.stringify(ev) })
        return
      }
      const provider = resolveProviderForModel(reqBody.model)
      if (!provider) {
        const ev: StreamEvent = {
          type: 'error',
          code: 'model_not_found',
          message: `no provider for model ${reqBody.model}`,
        }
        await stream.writeSSE({ event: 'error', data: JSON.stringify(ev) })
        return
      }
      const { requestId, controller } = newRequest()
      reqBody.requestId = requestId

      // Resolve the skill, build the system prompt, and fold the document
      // context in. Until this existed the `context` field crossed the wire and
      // was then dropped by every adapter, so the model was never told anything
      // about the document it was being asked about.
      let skill: Skill | undefined
      if (reqBody.skillId) {
        skill = findSkill(getSkills().skills, reqBody.skillId)
        if (!skill) {
          logger.warn({ msg: 'unknown skill requested', skillId: reqBody.skillId })
        }
      }

      // The pane sends the document tool catalog; the host unions in whichever
      // MCP tools the user has enabled. Doing the union here means the pane
      // cannot ask for an MCP tool that consent has not granted.
      const paneTools: ToolDefinition[] = Array.isArray(reqBody.tools) ? reqBody.tools : []
      const mcpTools = enabledToolDefinitions()
      const tools = [...paneTools, ...mcpTools]
      reqBody.tools = tools.length > 0 ? tools : undefined

      const systemPrompt = buildSystemPrompt({
        req: reqBody,
        skillPrompt: skill?.prompt,
        hasTools: tools.length > 0,
      })
      reqBody.messages = withSystemPrompt(reqBody, systemPrompt)

      // Forward client-disconnect to the upstream provider fetch. Hono's
      // streamSSE does not wire c.req.raw.signal to the AbortController on
      // Node, so without this the pane closing mid-stream leaves the upstream
      // HTTP connection open until the provider finishes — wasting tokens and
      // leaking an InFlight entry.
      const onDisconnect = () => controller.abort()
      const reqSignal = c.req.raw.signal
      if (reqSignal) {
        if (reqSignal.aborted) onDisconnect()
        else reqSignal.addEventListener('abort', onDisconnect, { once: true })
      }

      const startEv: StreamEvent = { type: 'start', requestId, model: reqBody.model }
      try {
        await stream.writeSSE({ event: 'start', data: JSON.stringify(startEv) })
      } catch {
        // Client already gone; abort upstream and bail.
        onDisconnect()
        return
      }

      let sawTerminal = false
      try {
        const source = provider.stream(reqBody, controller.signal)
        // splitReasoning sits inside wrapStream so inline <think> blocks are
        // separated before the pane ever sees them. Providers that expose
        // reasoning as its own field already emit `reasoning` events, which
        // pass through untouched.
        const wrapped = wrapStream(splitReasoning(source), requestId, controller)
        const dropReasoning = loadConfig().settings.showReasoning === false
        for await (const ev of wrapped) {
          if (controller.signal.aborted) return
          if (ev.type === 'reasoning' && dropReasoning) continue
          await stream.writeSSE({ event: ev.type, data: JSON.stringify(ev) })
          if (ev.type === 'done' || ev.type === 'error') {
            sawTerminal = true
            return
          }
        }
        // The provider stream ended without emitting a terminal event. This is
        // the normal case for Ollama/OpenAI/Anthropic (only opencode emits
        // `done`). Synthesize a `done` so the pane doesn't treat a clean close
        // as a stream interruption.
        if (!sawTerminal && !controller.signal.aborted) {
          const doneEv: StreamEvent = { type: 'done', requestId, finishReason: 'stop' }
          await stream.writeSSE({ event: 'done', data: JSON.stringify(doneEv) })
        }
      } catch (e) {
        // Never leak internal error messages to the webview — they may contain
        // file paths or stack fragments. Log full detail server-side, send a
        // generic message to the client.
        if (e instanceof ProviderError) {
          logger.warn({ msg: 'provider error', code: e.code, message: e.message })
          const ev: StreamEvent = {
            type: 'error',
            code: e.code,
            message: e.message,
            retryable: e.retryable,
          }
          await stream.writeSSE({ event: 'error', data: JSON.stringify(ev) })
        } else {
          logger.error({
            msg: 'unhandled error in chat stream',
            error: String(e),
            stack: (e as Error).stack,
          })
          const ev: StreamEvent = {
            type: 'error',
            code: 'internal_error',
            message: 'internal error',
          }
          await stream.writeSSE({ event: 'error', data: JSON.stringify(ev) })
        }
      } finally {
        if (reqSignal) reqSignal.removeEventListener('abort', onDisconnect)
        completeRequest(requestId)
      }
    })
  })

  app.post('/api/chat/cancel', async (c) => {
    let body: { requestId?: string }
    try {
      body = (await c.req.json()) as { requestId?: string }
    } catch {
      return c.json({ code: 'bad_request', message: 'invalid JSON' }, 400)
    }
    if (!body.requestId) {
      return c.json({ code: 'bad_request', message: 'requestId required' }, 400)
    }
    const ok = cancelRequest(body.requestId)
    return c.json({ cancelled: ok })
  })

  app.put('/api/providers/:id/key', async (c) => {
    const id = c.req.param('id')
    let body: { key?: string }
    try {
      body = (await c.req.json()) as { key?: string }
    } catch {
      return c.json({ code: 'bad_request', message: 'invalid JSON' }, 400)
    }
    if (!body.key) {
      return c.json({ code: 'bad_request', message: 'key required' }, 400)
    }
    try {
      setSecret(id, body.key)
      return c.json({ ok: true, configured: true })
    } catch (e) {
      logger.error({
        msg: 'failed to store provider key',
        providerId: id,
        error: String((e as Error).message ?? e),
      })
      return c.json({ code: 'error', message: 'failed to store key' }, 500)
    }
  })

  app.delete('/api/providers/:id/key', (c) => {
    const id = c.req.param('id')
    const ok = deleteSecret(id)
    return c.json({ ok, configured: ok })
  })

  app.post('/api/providers/:id/test', async (c) => {
    const id = c.req.param('id')
    const provider = getProvider(id)
    if (!provider) {
      return c.json({ code: 'not_found', message: `unknown provider ${id}` }, 404)
    }
    try {
      const reachable = await provider.isReachable()
      const models = reachable ? await provider.listModels().catch(() => []) : []
      return c.json({ ok: reachable, reachable, modelCount: models.length })
    } catch (e) {
      // Log full detail server-side; return only a generic flag to the client.
      // Upstream error bodies can echo headers (including API keys) if the
      // user misconfigured baseUrl toward a logging proxy.
      logger.warn({
        msg: 'provider test failed',
        providerId: id,
        error: String((e as Error).message ?? e),
      })
      return c.json({ ok: false, reachable: false })
    }
  })

  // ─── Skills ──────────────────────────────────────────────────────────

  app.get('/api/skills', (c) => {
    const host = c.req.query('host')
    const { skills, userDir } = getSkills(c.req.query('refresh') === '1')
    const filtered = isHostKind(host) ? skillsForHost(skills, host) : skills
    return c.json({ skills: filtered, userDir })
  })

  app.put('/api/skills/:id', async (c) => {
    let body: unknown
    try {
      body = await c.req.json()
    } catch {
      return c.json({ code: 'bad_request', message: 'invalid JSON' }, 400)
    }
    const parsed = skillWriteSchema.safeParse(body)
    if (!parsed.success) {
      return c.json({ code: 'bad_request', message: 'validation failed' }, 400)
    }
    try {
      // The id comes from the path, not the body, and saveUserSkill sanitizes
      // it into a filename — a body claiming `../../evil` cannot escape the
      // skills directory.
      saveUserSkill({
        ...parsed.data,
        id: c.req.param('id'),
        builtIn: false,
        source: 'user',
      })
      return c.json({ ok: true })
    } catch (e) {
      logger.warn({ msg: 'failed to save skill', error: String((e as Error).message ?? e) })
      return c.json({ code: 'error', message: 'could not save skill' }, 400)
    }
  })

  app.delete('/api/skills/:id', (c) => {
    const ok = deleteUserSkill(c.req.param('id'))
    return c.json({ ok })
  })

  // ─── Opencode import ─────────────────────────────────────────────────

  /** Shared response shape for preview and commit. Base URLs are included so
   *  the user can see *where* an imported provider points before enabling it;
   *  API keys never appear here. */
  const importSummary = (result: ReturnType<typeof importFromOpencode>) => ({
    ok: result.ok,
    providerCount: result.providerCount,
    mcpCount: result.mcpCount,
    agentCount: result.agentCount,
    providers: result.providers.map((p) => ({
      id: p.id,
      name: p.name,
      baseUrl: p.baseUrl,
      modelCount: p.models?.length ?? 0,
    })),
    mcpServers: result.mcpServers.map((s) => ({ id: s.id, name: s.name, transport: s.transport })),
    agents: result.agents.map((a) => ({ id: a.id, name: a.name, model: a.model })),
    linkedProviders: result.linkedProviders,
    sources: result.sources,
    searched: result.searched,
    errors: result.errors,
  })

  app.post('/api/opencode/import', async (c) => {
    try {
      const result = importFromOpencode()
      if (!result.ok) {
        return c.json({ code: 'not_found', message: result.errors[0] ?? 'nothing to import' }, 404)
      }
      const cfg = loadConfig()
      // Merge imports into settings: replace by id.
      const providersById = new Map(cfg.settings.importedProviders.map((p) => [p.id, p]))
      for (const p of result.providers) providersById.set(p.id, p)
      const mcpById = new Map(cfg.settings.mcpServers.map((s) => [s.id, s]))
      for (const s of result.mcpServers) {
        const existing = mcpById.get(s.id)
        // Preserve any existing enabled state; imported servers default off.
        mcpById.set(s.id, { ...s, enabled: existing?.enabled ?? false })
      }
      const agentsById = new Map(cfg.settings.agents.map((a) => [a.id, a]))
      for (const a of result.agents) {
        const existing = agentsById.get(a.id)
        agentsById.set(a.id, { ...a, enabled: existing?.enabled ?? false })
      }
      const next: typeof cfg.settings = {
        ...cfg.settings,
        importedProviders: [...providersById.values()],
        mcpServers: [...mcpById.values()],
        agents: [...agentsById.values()],
      }
      updateSettings(next)
      // Register newly imported providers into the live registry so their
      // models appear without a host restart. Ollama Cloud is registered as
      // a dedicated adapter; other imported providers are OpenAI-compatible.
      for (const p of result.providers) {
        if (p.id === 'ollama-cloud') {
          registerProvider(new OllamaCloudAdapter({ baseUrl: p.baseUrl, modelsHint: p.models }))
        } else {
          registerProvider(createImportedAdapter(p))
        }
      }
      // Providers that already had an adapter only gained a key — but a key is
      // exactly what gates them appearing in the model list, so the cache still
      // has to be dropped and their models re-listed.
      bustModelCache()
      // Reconcile MCP servers with newly imported ones.
      void syncServers().catch((e) =>
        logger.warn({ msg: 'mcp sync failed after opencode import', error: String(e) }),
      )
      return c.json(importSummary(result))
    } catch (e) {
      logger.warn({ msg: 'opencode import failed', error: String((e as Error).message ?? e) })
      return c.json({ code: 'error', message: 'import failed' }, 500)
    }
  })

  app.get('/api/opencode/import', (c) => {
    // Preview what would be imported without saving or storing keys.
    return c.json(importSummary(importFromOpencode(false)))
  })

  // ─── MCP ─────────────────────────────────────────────────────────────

  app.get('/api/mcp/servers', async (c) => {
    if (c.req.query('refresh') === '1') await syncServers()
    return c.json({ servers: listServers() })
  })

  app.post('/api/mcp/servers', async (c) => {
    let body: unknown
    try {
      body = await c.req.json()
    } catch {
      return c.json({ code: 'bad_request', message: 'invalid JSON' }, 400)
    }
    const parsed = mcpServerSchema.safeParse(body)
    if (!parsed.success) {
      return c.json(
        { code: 'bad_request', message: 'validation failed', errors: parsed.error.message },
        400,
      )
    }
    const cfg = loadConfig()
    const next = cfg.settings.mcpServers.filter((s) => s.id !== parsed.data.id)
    next.push(parsed.data)
    updateSettings({ ...cfg.settings, mcpServers: next })
    await syncServers()
    return c.json({ servers: listServers() })
  })

  app.delete('/api/mcp/servers/:id', async (c) => {
    const id = c.req.param('id')
    const cfg = loadConfig()
    const { [id]: _removed, ...consent } = cfg.settings.mcpToolConsent
    updateSettings({
      ...cfg.settings,
      mcpServers: cfg.settings.mcpServers.filter((s) => s.id !== id),
      mcpToolConsent: consent,
    })
    await syncServers()
    return c.json({ servers: listServers() })
  })

  app.post('/api/mcp/consent', async (c) => {
    let body: { serverId?: string; tool?: string; enabled?: boolean }
    try {
      body = (await c.req.json()) as typeof body
    } catch {
      return c.json({ code: 'bad_request', message: 'invalid JSON' }, 400)
    }
    if (!body.serverId || !body.tool || typeof body.enabled !== 'boolean') {
      return c.json({ code: 'bad_request', message: 'serverId, tool, enabled required' }, 400)
    }
    const cfg = loadConfig()
    const forServer = { ...(cfg.settings.mcpToolConsent[body.serverId] ?? {}) }
    if (body.enabled) forServer[body.tool] = true
    else delete forServer[body.tool]
    updateSettings({
      ...cfg.settings,
      mcpToolConsent: { ...cfg.settings.mcpToolConsent, [body.serverId]: forServer },
    })
    return c.json({ servers: listServers() })
  })

  app.post('/api/mcp/call', async (c) => {
    let body: { tool?: string; arguments?: Record<string, unknown> }
    try {
      body = (await c.req.json()) as typeof body
    } catch {
      return c.json({ code: 'bad_request', message: 'invalid JSON' }, 400)
    }
    if (!body.tool) {
      return c.json({ code: 'bad_request', message: 'tool required' }, 400)
    }
    // callTool enforces consent itself and returns an error *result* rather
    // than throwing, so a model asking for a disabled tool gets a message it
    // can act on instead of killing the turn.
    const result = await callMcpTool(body.tool, body.arguments ?? {})
    return c.json(result)
  })

  // ─── Extension pairing ─────────────────────────────────────────────────
  //
  // Deliberately outside `/api/*`, because it is the one request an extension
  // can make before it holds a token: it is how it gets one. The allowlist is
  // what gates it, and nothing else — so the check is done inline here rather
  // than inherited from middleware that also knows about bearer tokens.
  //
  // This does not widen the blast radius of the token. Anything that can run
  // as a paired extension origin could equally read the token out of the pane
  // HTML the host already serves unauthenticated on this port. What pairing
  // buys is that *only* a named extension can, and that the user chose it.
  app.options('/pair', (c) => {
    const origin = c.req.header('origin') ?? ''
    if (!isPairedExtension(origin, pairedOrigins())) {
      return c.json({ code: 'forbidden', message: 'forbidden' }, 403)
    }
    applyCors(c, origin)
    return c.body(null, 204)
  })

  app.get('/pair', (c) => {
    const origin = c.req.header('origin') || (c.req.header('x-openofficellm-extension') ?? '')
    if (!isPairedExtension(origin, pairedOrigins())) {
      logger.warn({ msg: 'pairing refused', origin: origin || '(none)' })
      return c.json(
        {
          code: 'not_paired',
          message:
            'This extension is not paired with the host. Run the host with --pair <extension-id> to allow it.',
        },
        403,
      )
    }
    applyCors(c, origin)
    // No-store: the token must not sit in any HTTP cache, least of all one
    // shared across extension profiles.
    c.header('Cache-Control', 'no-store')
    logger.info({ msg: 'pairing token issued', origin })
    return c.json({ token: auth.token, version: HOST_VERSION, port })
  })

  app.get('/office/office.js', () => {
    const content = getCachedOfficeJs()
    if (!content) {
      return new Response('// office.js not cached', {
        status: 503,
        headers: { 'content-type': 'text/javascript' },
      })
    }
    return new Response(content, {
      status: 200,
      headers: {
        'content-type': 'text/javascript; charset=utf-8',
        'cache-control': 'public, max-age=3600',
      },
    })
  })

  app.all('/api/*', (c) => {
    return c.json({ code: 'not_found', message: 'not found' }, 404)
  })

  app.get('*', (c) => {
    const resp = serveStatic(c, {
      rootDir: addinDist,
      token: auth.token,
      officeJsLocalAvailable: officeJs.meta.source !== 'cdn',
    })
    if (resp) return resp
    return c.notFound()
  })

  const server = serve({
    port,
    hostname: HOST_INTERFACE,
    fetch: app.fetch,
    createServer: ((
      serverOpts: Record<string, unknown>,
      listener: (req: http.IncomingMessage, res: http.ServerResponse) => void,
    ) =>
      https.createServer(
        { key: cert.leafKeyPem, cert: cert.leafCertPem, ...serverOpts },
        listener,
      )) as never,
  })

  await new Promise<void>((resolve, reject) => {
    server.once('listening', () => resolve())
    server.once('error', reject)
  })

  logger.info({ msg: `listening on https://${HOST_INTERFACE}:${port}` })

  // Rewrite the manifest from the port we actually bound. If the preferred
  // port was taken and selectPort scanned forward, every absolute URL in the
  // manifest is now wrong, and the only symptom the user would see is a task
  // pane that refuses to load. Office picks the new file up on its next
  // launch, so this self-heals across a restart of Word rather than needing
  // --repair.
  try {
    const m = writeManifest({ port })
    if (m.changed) {
      logger.info({
        msg: 'manifest regenerated; restart Word or Excel to pick it up',
        path: m.path,
      })
    }
  } catch (e) {
    logger.warn({
      msg: 'could not write add-in manifest',
      error: String((e as Error).message ?? e),
    })
  }

  if (isFallbackMode()) {
    logger.warn({ msg: 'secrets running in fallback mode — keys not DPAPI-protected' })
  }

  const close = async (): Promise<void> => {
    stopDiscoveryLoop()
    // Before closing the socket: stdio MCP servers are child processes, and
    // leaving them behind orphans one per host restart.
    await stopAllServers().catch(() => undefined)
    await new Promise<void>((resolve) => {
      server.close(() => resolve())
    })
    hooks.onShutdown?.()
  }

  return close
}
