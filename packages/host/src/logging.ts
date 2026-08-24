import fs from 'node:fs'
import path from 'node:path'
import { resolveLogDir, resolveLogPath } from './paths.js'

export type LogLevel = 'debug' | 'info' | 'warn' | 'error'

interface LoggerOptions {
  level?: LogLevel
  verboseBodies?: boolean
}

interface LogRecord {
  ts: string
  level: LogLevel
  msg: string
  [k: string]: unknown
}

const LEVEL_ORDER: Record<LogLevel, number> = { debug: 10, info: 20, warn: 30, error: 40 }

let currentLevel: LogLevel = 'info'
let verboseBodies = false
let initialized = false
let stream: fs.WriteStream | null = null
let currentLogDate = ''

const KEY_PATTERNS = [
  /key/i,
  /token/i,
  /secret/i,
  /authorization/i,
  /password/i,
  /api[-_]?key/i,
  /bearer/i,
  /anthropic/i,
  /openai/i,
  /sk-/i,
]

function redact(value: unknown, seen = new Set<unknown>()): unknown {
  if (value === null || value === undefined) return value
  if (typeof value === 'string') {
    if (!verboseBodies && value.length > 200) return value.slice(0, 80) + '…<redacted>'
    return value
  }
  if (typeof value !== 'object') return value
  if (seen.has(value)) return '[Circular]'
  seen.add(value)
  if (Array.isArray(value)) {
    return value.map((v) => redact(v, seen))
  }
  const obj = value as Record<string, unknown>
  const out: Record<string, unknown> = {}
  for (const [k, v] of Object.entries(obj)) {
    if (KEY_PATTERNS.some((p) => p.test(k))) {
      out[k] = '[redacted]'
    } else if (k === 'content' && !verboseBodies && typeof v === 'string') {
      out[k] = '[body redacted]'
    } else if (k === 'messages' && !verboseBodies) {
      out[k] = `[${Array.isArray(v) ? v.length : 0} messages redacted]`
    } else {
      out[k] = redact(v, seen)
    }
  }
  return out
}

function openStream(): void {
  const dir = resolveLogDir()
  try {
    fs.mkdirSync(dir, { recursive: true })
  } catch {
    // ignore
  }
  stream = fs.createWriteStream(resolveLogPath(), { flags: 'a' })
  stream.on('error', (e) => {
    console.error('log stream error:', e)
  })
}

function rotateIfNeeded(): void {
  const today = new Date().toISOString().slice(0, 10)
  if (currentLogDate === '') {
    currentLogDate = today
  }
  if (currentLogDate === today && stream) return
  if (stream) {
    try {
      stream.end()
    } catch {
      // ignore
    }
    stream = null
  }
  if (currentLogDate !== today) {
    rotateOldLogs()
    currentLogDate = today
  }
  openStream()
}

function rotateOldLogs(): void {
  const logDir = resolveLogDir()
  try {
    const files = fs.readdirSync(logDir)
    const cutoff = Date.now() - 7 * 24 * 60 * 60 * 1000
    for (const f of files) {
      if (!f.startsWith('host.') || !f.endsWith('.log')) continue
      if (f === 'host.log') continue
      const datePart = f.slice('host.'.length, -'.log'.length)
      const d = Date.parse(datePart)
      if (Number.isNaN(d)) continue
      if (d < cutoff) {
        try {
          fs.unlinkSync(path.join(logDir, f))
        } catch {
          // ignore
        }
      }
    }
  } catch {
    // ignore — directory may not exist yet
  }
}

function writeLine(line: string): void {
  rotateIfNeeded()
  if (stream) {
    try {
      stream.write(line + '\n')
    } catch {
      // fall back to stderr
      process.stderr.write(line + '\n')
    }
  } else {
    process.stderr.write(line + '\n')
  }
}

function log(level: LogLevel, msg: string, extra?: Record<string, unknown>): void {
  if (LEVEL_ORDER[level] < LEVEL_ORDER[currentLevel]) return
  const record: LogRecord = {
    ts: new Date().toISOString(),
    level,
    msg,
    ...(extra ?? {}),
  }
  const safe = redact(record) as LogRecord
  const line = JSON.stringify(safe)
  writeLine(line)
  if (level === 'error') {
    process.stderr.write(line + '\n')
  }
}

export function initLogger(opts: LoggerOptions = {}): void {
  if (opts.level) currentLevel = opts.level
  if (opts.verboseBodies !== undefined) verboseBodies = opts.verboseBodies
  if (!initialized) {
    initialized = true
    currentLogDate = new Date().toISOString().slice(0, 10)
    openStream()
    rotateOldLogs()
  }
}

export const logger = {
  debug: (extra: Record<string, unknown> | string) => {
    if (typeof extra === 'string') log('debug', extra)
    else log('debug', (extra.msg as string) ?? '', extra)
  },
  info: (extra: Record<string, unknown> | string) => {
    if (typeof extra === 'string') log('info', extra)
    else log('info', (extra.msg as string) ?? '', extra)
  },
  warn: (extra: Record<string, unknown> | string) => {
    if (typeof extra === 'string') log('warn', extra)
    else log('warn', (extra.msg as string) ?? '', extra)
  },
  error: (extra: Record<string, unknown> | string) => {
    if (typeof extra === 'string') log('error', extra)
    else log('error', (extra.msg as string) ?? '', extra)
  },
  isVerbose: () => verboseBodies,
  level: () => currentLevel,
}

export function flushLogs(): Promise<void> {
  return new Promise((resolve) => {
    if (stream) {
      stream.end(() => resolve())
    } else {
      resolve()
    }
  })
}
