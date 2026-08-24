// MCP server lifecycle, tool catalog, and consent enforcement (P5.10, P5.11).
//
// The consent model is the point of this file. Every tool is disabled until the
// user turns it on, per server AND per tool, and the check happens here — at
// the call site — not in the UI. A tool that is not explicitly enabled in
// config cannot be invoked no matter what the model asks for, so a prompt
// injection that talks the model into calling `delete_everything` gets an
// error string back instead of a side effect.

import type {
  McpServerConfig,
  McpServerInfo,
  McpServerStatus,
  McpTool,
  ToolDefinition,
} from '@openofficellm/shared'
import { loadConfig } from '../config.js'
import { logger } from '../logging.js'
import { McpClient, type McpToolResult } from './client.js'

/** Namespace prefix keeping MCP tools from colliding with document tools. */
export const MCP_PREFIX = 'mcp__'

interface ServerState {
  cfg: McpServerConfig
  client: McpClient | null
  status: McpServerStatus
  error?: string
  tools: McpTool[]
  /** Consecutive start failures, for backoff. */
  failures: number
  nextRetryAt: number
}

const servers = new Map<string, ServerState>()

/** `mcp__<server>__<tool>`. Server ids are sanitized on save so the separator
 *  stays unambiguous. */
export function qualifyToolName(serverId: string, tool: string): string {
  return `${MCP_PREFIX}${serverId}__${tool}`
}

export function parseQualifiedToolName(
  qualified: string,
): { serverId: string; tool: string } | null {
  if (!qualified.startsWith(MCP_PREFIX)) return null
  const rest = qualified.slice(MCP_PREFIX.length)
  const sep = rest.indexOf('__')
  if (sep <= 0) return null
  return { serverId: rest.slice(0, sep), tool: rest.slice(sep + 2) }
}

export function isMcpToolName(name: string): boolean {
  return name.startsWith(MCP_PREFIX)
}

/** True if the user has explicitly enabled this tool on this server. Absent
 *  means disabled — never default a newly appeared tool to on. */
export function isToolEnabled(serverId: string, tool: string): boolean {
  const settings = loadConfig().settings
  const server = settings.mcpServers.find((s) => s.id === serverId)
  if (!server || !server.enabled) return false
  return settings.mcpToolConsent[serverId]?.[tool] === true
}

function stateFor(cfg: McpServerConfig): ServerState {
  const existing = servers.get(cfg.id)
  if (existing) {
    existing.cfg = cfg
    return existing
  }
  const fresh: ServerState = {
    cfg,
    client: null,
    status: 'stopped',
    tools: [],
    failures: 0,
    nextRetryAt: 0,
  }
  servers.set(cfg.id, fresh)
  return fresh
}

async function startServer(state: ServerState): Promise<void> {
  if (state.status === 'ready' || state.status === 'starting') return
  if (Date.now() < state.nextRetryAt) return

  state.status = 'starting'
  state.error = undefined
  const client = new McpClient(state.cfg)
  try {
    await client.start()
    state.tools = await client.listTools()
    state.client = client
    state.status = 'ready'
    state.failures = 0
    logger.info({ msg: 'mcp server ready', server: state.cfg.id, tools: state.tools.length })
  } catch (e) {
    await client.close().catch(() => undefined)
    state.client = null
    state.tools = []
    state.status = 'error'
    // The message can contain a spawn command line, which for an HTTP server
    // may carry a token in the URL. Keep it short and log the detail instead.
    state.error = String((e as Error).message ?? e).slice(0, 200)
    state.failures += 1
    // Capped exponential backoff: a server that is broken should not be
    // respawned on every settings fetch.
    const backoff = Math.min(60_000, 2_000 * 2 ** Math.min(state.failures, 5))
    state.nextRetryAt = Date.now() + backoff
    logger.warn({
      msg: 'mcp server failed to start',
      server: state.cfg.id,
      error: state.error,
      retryInMs: backoff,
    })
  }
}

async function stopServer(state: ServerState): Promise<void> {
  if (state.client) await state.client.close().catch(() => undefined)
  state.client = null
  state.tools = []
  state.status = 'stopped'
  state.error = undefined
}

/**
 * Reconcile running servers against config: start newly enabled ones, stop
 * disabled or deleted ones. Safe to call repeatedly.
 */
export async function syncServers(): Promise<void> {
  const configured = loadConfig().settings.mcpServers
  const wanted = new Map(configured.map((c) => [c.id, c]))

  for (const [id, state] of [...servers]) {
    const cfg = wanted.get(id)
    if (!cfg || !cfg.enabled) {
      await stopServer(state)
      if (!cfg) servers.delete(id)
      continue
    }
    // A changed command or url means the running process is the wrong one.
    if (JSON.stringify(state.cfg) !== JSON.stringify(cfg)) {
      await stopServer(state)
      state.cfg = cfg
      state.failures = 0
      state.nextRetryAt = 0
    }
  }

  await Promise.all(configured.filter((c) => c.enabled).map((cfg) => startServer(stateFor(cfg))))
}

export async function stopAllServers(): Promise<void> {
  await Promise.all([...servers.values()].map((s) => stopServer(s)))
  servers.clear()
}

/** Everything the settings UI needs: definitions, live status, discovered
 *  tools, and each tool's consent state. */
export function listServers(): McpServerInfo[] {
  const settings = loadConfig().settings
  return settings.mcpServers.map((cfg) => {
    const state = servers.get(cfg.id)
    const consent = settings.mcpToolConsent[cfg.id] ?? {}
    return {
      ...cfg,
      status: state?.status ?? 'stopped',
      error: state?.error,
      tools: (state?.tools ?? []).map((t) => ({ ...t, enabled: consent[t.name] === true })),
    }
  })
}

/**
 * Tool definitions for every MCP tool the user has enabled, ready to append to
 * the document tool catalog for a chat turn.
 */
export function enabledToolDefinitions(): ToolDefinition[] {
  const out: ToolDefinition[] = []
  for (const state of servers.values()) {
    if (state.status !== 'ready') continue
    for (const tool of state.tools) {
      if (!isToolEnabled(state.cfg.id, tool.name)) continue
      const schema = tool.inputSchema as ToolDefinition['parameters'] | undefined
      out.push({
        name: qualifyToolName(state.cfg.id, tool.name),
        description: `[${state.cfg.name}] ${tool.description ?? tool.name}`,
        parameters:
          schema && schema.type === 'object'
            ? schema
            : { type: 'object', properties: {}, required: [] },
      })
    }
  }
  return out
}

/**
 * Invoke an MCP tool by its qualified name.
 *
 * Returns an error *result* rather than throwing for anything the model could
 * have caused — an unknown tool, a disabled tool, a server that is down. The
 * model can read that and adapt; an exception would abort the whole turn.
 */
export async function callTool(
  qualified: string,
  args: Record<string, unknown>,
): Promise<McpToolResult> {
  const parsed = parseQualifiedToolName(qualified)
  if (!parsed) {
    return { content: `Unknown tool: ${qualified}`, isError: true }
  }
  const { serverId, tool } = parsed

  if (!isToolEnabled(serverId, tool)) {
    return {
      content: `Tool "${tool}" on server "${serverId}" is not enabled. The user must enable it in Settings → MCP servers before it can be used.`,
      isError: true,
    }
  }

  const state = servers.get(serverId)
  if (!state || state.status !== 'ready' || !state.client) {
    // A disabled-then-enabled server may not have been started yet.
    await syncServers()
  }
  const ready = servers.get(serverId)
  if (!ready || !ready.client || ready.status !== 'ready') {
    return {
      content: `MCP server "${serverId}" is not running${ready?.error ? `: ${ready.error}` : ''}.`,
      isError: true,
    }
  }

  try {
    const result = await ready.client.callTool(tool, args)
    logger.info({
      msg: 'mcp tool called',
      server: serverId,
      tool,
      isError: result.isError,
      // Arguments and results are deliberately not logged: they routinely
      // contain document text and whatever the tool reached out and fetched.
    })
    return result
  } catch (e) {
    const message = String((e as Error).message ?? e).slice(0, 200)
    logger.warn({ msg: 'mcp tool call failed', server: serverId, tool, error: message })
    return { content: `Tool call failed: ${message}`, isError: true }
  }
}
