import { logger } from '../logging.js'
import { RateLimitError, ProviderError } from './types.js'

const MAX_ATTEMPTS = 3
const BASE_DELAY_MS = 500
const MAX_DELAY_MS = 8000

export interface FetchWithRetryOptions {
  maxAttempts?: number
  baseDelayMs?: number
  maxDelayMs?: number
  isIdempotent?: boolean
  signal?: AbortSignal
  logTag?: string
}

function sleep(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) {
      reject(new Error('aborted'))
      return
    }
    const timer = setTimeout(() => {
      signal?.removeEventListener('abort', onAbort)
      resolve()
    }, ms)
    const onAbort = () => {
      clearTimeout(timer)
      reject(new Error('aborted'))
    }
    signal?.addEventListener('abort', onAbort, { once: true })
  })
}

function backoffMs(attempt: number, base: number, cap: number): number {
  const exp = base * Math.pow(2, attempt - 1)
  const jitter = Math.random() * base
  return Math.min(cap, exp + jitter)
}

export async function fetchWithRetry(
  fn: (attempt: number) => Promise<Response>,
  opts: FetchWithRetryOptions = {},
): Promise<Response> {
  const maxAttempts = opts.maxAttempts ?? MAX_ATTEMPTS
  const base = opts.baseDelayMs ?? BASE_DELAY_MS
  const cap = opts.maxDelayMs ?? MAX_DELAY_MS
  const idempotent = opts.isIdempotent ?? true
  let lastErr: unknown
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    if (opts.signal?.aborted) throw new Error('aborted')
    try {
      const res = await fn(attempt)
      if (!shouldRetry(res.status)) return res
      // Non-idempotent requests (e.g. a chat that may invoke tools) must not
      // be retried once the server saw them — a 5xx after the model already
      // emitted a tool_call would, on retry, re-send the prompt and risk a
      // second tool invocation / double document edit. Only retry network
      // errors (caught below), where no response was received.
      if (!idempotent) {
        return res
      }
      if (attempt >= maxAttempts) return res
      const delay = backoffMs(attempt, base, cap)
      logger.warn({
        msg: 'fetch retryable status, backing off',
        tag: opts.logTag,
        status: res.status,
        attempt,
        delayMs: delay,
      })
      try {
        await sleep(delay, opts.signal)
      } catch {
        throw new Error('aborted')
      }
      lastErr = new ProviderError('upstream', `status ${res.status}`, {
        retryable: true,
        statusCode: res.status,
      })
    } catch (e) {
      lastErr = e
      if ((e as Error).message === 'aborted' || opts.signal?.aborted) {
        throw e
      }
      if (e instanceof ProviderError && !e.retryable) throw e
      if (attempt >= maxAttempts) throw e
      // Network errors (no response received) are safe to retry for both
      // idempotent and non-idempotent requests — the server never saw it.
      const delay = backoffMs(attempt, base, cap)
      logger.warn({
        msg: 'fetch error, backing off',
        tag: opts.logTag,
        attempt,
        delayMs: delay,
        error: String((e as Error).message ?? e),
      })
      try {
        await sleep(delay, opts.signal)
      } catch {
        throw new Error('aborted')
      }
    }
  }
  throw lastErr ?? new ProviderError('unknown', 'retry loop exhausted')
}

function shouldRetry(status: number): boolean {
  return status === 429 || (status >= 500 && status <= 599)
}

export { RateLimitError }
