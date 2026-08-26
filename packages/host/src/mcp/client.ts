// MCP client core (P5.7–P5.9).
//
// Implements the slice of the Model Context Protocol we actually need:
// initialize, tools/list, tools/call. Both transports speak JSON-RPC 2.0 —
// stdio frames messages as newline-delimited JSON, streamable HTTP POSTs them
// and accepts either a JSON body or an SSE stream in reply.
//
// No SDK dependency. The MCP TypeScript SDK pulls a large tree into a process
// that already holds the user's API keys, for three method calls we can write
// in a page.

import { spawn, type ChildProcessWithoutNullStreams } from 'node:child_process'
import type { McpServerConfig, McpTool } from '@openofficellm/shared'
import { logger } from '../logging.js'
import { HOST_VERSION } from '../paths.js'

const PROTOCOL_VERSION = '2024-11-05'
const CLIENT_INFO = { name: 'OpenOfficeLLM', version: HOST_VERSION }

/** Wall-clock cap on any single request. A hung server must not wedge a chat
 *  turn forever — the model gets an error result and can carry on. */
const REQUEST_TIMEOUT_MS = 30_000
const INIT_TIMEOUT_MS = 15_000

interface JsonRpcResponse {
  jsonrpc: '2.0'
  id?: number | string
  result?: unknown
  error?: { code: number; message: string; data?: unknown }
}

export interface McpToolResult {
  content: string
  isError: boolean
}

export interface Transport {
  start(): Promise<void>
  request(method: string, params?: unknown, timeoutMs?: number): Promise<unknown>
  notify(method: string, params?: unknown): Promise<void>
  close(): Promise<void>
}

// ─── stdio ───────────────────────────────────────────────────────────────

class StdioTransport implements Transport {
  private child: ChildProcessWithoutNullStreams | null = null
  private nextId = 1
  private pending = new Map<
    number | string,
    { resolve: (v: unknown) => void; reject: (e: Error) => void; timer: NodeJS.Timeout }
  >()
  private buffer = ''

  constructor(private cfg: McpServerConfig) {}

  async start(): Promise<void> {
    if (!this.cfg.command) throw new Error('stdio server has no command')
    const child = spawn(this.cfg.command, this.cfg.args ?? [], {
      // Inherit the host's environment so servers find node/python on PATH,
      // then layer the server's own vars on top.
      env: { ...process.env, ...(this.cfg.env ?? {}) },
      stdio: ['pipe', 'pipe', 'pipe'],
      // Without this, a server whose command is a .cmd shim (every npx-based
      // server on Windows) fails to spawn with ENOENT.
      shell: process.platform === 'win32',
    }) as ChildProcessWithoutNullStreams
    this.child = child

    child.stdout.setEncoding('utf8')
    child.stdout.on('data', (chunk: string) => this.onData(chunk))

    child.stderr.setEncoding('utf8')
    child.stderr.on('data', (chunk: string) => {
      // Servers use stderr for ordinary logging, so this is not an error path.
      // Truncated because a chatty server should not be able to fill the log.
      logger.debug({ msg: 'mcp stderr', server: this.cfg.id, text: String(chunk).slice(0, 500) })
    })

    child.on('exit', (code, signal) => {
      const err = new Error(
        `MCP server exited (code ${code ?? 'null'}, signal ${signal ?? 'none'})`,
      )
      this.failAll(err)
      this.child = null
    })
    child.on('error', (err) => {
      this.failAll(err instanceof Error ? err : new Error(String(err)))
      this.child = null
    })

    await this.request(
      'initialize',
      {
        protocolVersion: PROTOCOL_VERSION,
        capabilities: {},
        clientInfo: CLIENT_INFO,
      },
      INIT_TIMEOUT_MS,
    )
    await this.notify('notifications/initialized')
  }

  private onData(chunk: string): void {
    this.buffer += chunk
    for (;;) {
      const nl = this.buffer.indexOf('\n')
      if (nl === -1) break
      const line = this.buffer.slice(0, nl).trim()
      this.buffer = this.buffer.slice(nl + 1)
      if (!line) continue
      let msg: JsonRpcResponse
      try {
        msg = JSON.parse(line) as JsonRpcResponse
      } catch {
        // Servers occasionally print a banner to stdout before speaking
        // JSON-RPC. Ignoring unparseable lines is the difference between a
        // usable server and a dead one.
        continue
      }
      if (msg.id === undefined) continue // a notification from the server
      const entry = this.pending.get(msg.id)
      if (!entry) continue
      clearTimeout(entry.timer)
      this.pending.delete(msg.id)
      if (msg.error) entry.reject(new Error(msg.error.message || 'MCP error'))
      else entry.resolve(msg.result)
    }
  }

  private failAll(err: Error): void {
    for (const [, entry] of this.pending) {
      clearTimeout(entry.timer)
      entry.reject(err)
    }
    this.pending.clear()
  }

  request(method: string, params?: unknown, timeoutMs = REQUEST_TIMEOUT_MS): Promise<unknown> {
    const child = this.child
    if (!child) return Promise.reject(new Error('MCP server not running'))
    const id = this.nextId++
    const payload = `${JSON.stringify({ jsonrpc: '2.0', id, method, params })}\n`
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id)
        reject(new Error(`MCP request timed out: ${method}`))
      }, timeoutMs)
      this.pending.set(id, { resolve, reject, timer })
      child.stdin.write(payload, (err) => {
        if (err) {
          clearTimeout(timer)
          this.pending.delete(id)
          reject(err)
        }
      })
    })
  }

  async notify(method: string, params?: unknown): Promise<void> {
    const child = this.child
    if (!child) return
    child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method, params })}\n`)
  }

  async close(): Promise<void> {
    this.failAll(new Error('MCP server stopped'))
    const child = this.child
    this.child = null
    if (!child) return
    child.kill()
    // SIGKILL after a grace period; a server that ignores SIGTERM would
    // otherwise outlive the host and hold its port or its handles.
    await new Promise<void>((resolve) => {
      const t = setTimeout(() => {
        try {
          child.kill('SIGKILL')
        } catch {
          // already gone
        }
        resolve()
      }, 2000)
      child.once('exit', () => {
        clearTimeout(t)
        resolve()
      })
    })
  }
}

// ─── streamable HTTP ─────────────────────────────────────────────────────

class HttpTransport implements Transport {
  private nextId = 1
  private sessionId: string | null = null

  constructor(private cfg: McpServerConfig) {}

  private headers(): Record<string, string> {
    const h: Record<string, string> = {
      'content-type': 'application/json',
      accept: 'application/json, text/event-stream',
      ...(this.cfg.headers ?? {}),
    }
    if (this.sessionId) h['mcp-session-id'] = this.sessionId
    return h
  }

  async start(): Promise<void> {
    await this.request(
      'initialize',
      { protocolVersion: PROTOCOL_VERSION, capabilities: {}, clientInfo: CLIENT_INFO },
      INIT_TIMEOUT_MS,
    )
    await this.notify('notifications/initialized')
  }

  async request(
    method: string,
    params?: unknown,
    timeoutMs = REQUEST_TIMEOUT_MS,
  ): Promise<unknown> {
    if (!this.cfg.url) throw new Error('http server has no url')
    const id = this.nextId++
    const ctrl = new AbortController()
    const timer = setTimeout(() => ctrl.abort(), timeoutMs)
    let res: Response
    try {
      res = await fetch(this.cfg.url, {
        method: 'POST',
        headers: this.headers(),
        body: JSON.stringify({ jsonrpc: '2.0', id, method, params }),
        signal: ctrl.signal,
      })
    } finally {
      clearTimeout(timer)
    }
    // The server assigns a session on initialize and expects it echoed back.
    const session = res.headers.get('mcp-session-id')
    if (session) this.sessionId = session

    if (!res.ok) {
      throw new Error(`MCP HTTP ${res.status}`)
    }

    const contentType = res.headers.get('content-type') ?? ''
    const text = await res.text()
    const msg = contentType.includes('text/event-stream')
      ? parseSseResponse(text, id)
      : (JSON.parse(text) as JsonRpcResponse)

    if (!msg) throw new Error('no JSON-RPC response in stream')
    if (msg.error) throw new Error(msg.error.message || 'MCP error')
    return msg.result
  }

  async notify(method: string, params?: unknown): Promise<void> {
    if (!this.cfg.url) return
    try {
      await fetch(this.cfg.url, {
        method: 'POST',
        headers: this.headers(),
        body: JSON.stringify({ jsonrpc: '2.0', method, params }),
      })
    } catch {
      // Notifications are fire-and-forget by definition.
    }
  }

  async close(): Promise<void> {
    this.sessionId = null
  }
}

/** Pull the response matching `id` out of an SSE body. */
function parseSseResponse(text: string, id: number): JsonRpcResponse | null {
  let fallback: JsonRpcResponse | null = null
  for (const block of text.split(/\n\n/)) {
    const dataLines = block
      .split('\n')
      .filter((l) => l.startsWith('data:'))
      .map((l) => l.slice(5).trim())
    if (dataLines.length === 0) continue
    try {
      const parsed = JSON.parse(dataLines.join('\n')) as JsonRpcResponse
      if (parsed.id === id) return parsed
      if (parsed.id !== undefined && fallback === null) fallback = parsed
    } catch {
      continue
    }
  }
  return fallback
}

// ─── Client ──────────────────────────────────────────────────────────────

export class McpClient {
  private transport: Transport

  constructor(readonly cfg: McpServerConfig) {
    this.transport = cfg.transport === 'stdio' ? new StdioTransport(cfg) : new HttpTransport(cfg)
  }

  async start(): Promise<void> {
    await this.transport.start()
  }

  async listTools(): Promise<McpTool[]> {
    const result = (await this.transport.request('tools/list')) as {
      tools?: Array<{ name?: string; description?: string; inputSchema?: Record<string, unknown> }>
    }
    return (result?.tools ?? [])
      .filter((t) => typeof t.name === 'string' && t.name.length > 0)
      .map((t) => ({
        serverId: this.cfg.id,
        name: t.name!,
        description: t.description,
        inputSchema: t.inputSchema,
        // Consent lives in config, not in the server's own advertisement. A
        // server does not get to declare itself trusted.
        enabled: false,
      }))
  }

  async callTool(name: string, args: Record<string, unknown>): Promise<McpToolResult> {
    const result = (await this.transport.request('tools/call', {
      name,
      arguments: args,
    })) as {
      content?: Array<{ type?: string; text?: string }>
      isError?: boolean
    }
    const parts = (result?.content ?? [])
      .map((c) => (c.type === 'text' && typeof c.text === 'string' ? c.text : ''))
      .filter(Boolean)
    return {
      content: parts.join('\n') || '(tool returned no textual content)',
      isError: result?.isError === true,
    }
  }

  async close(): Promise<void> {
    await this.transport.close()
  }
}
