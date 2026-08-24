import type {
  ChatRequest,
  ModelInfo,
  ProviderCapabilities,
  StreamEvent,
} from '@openofficellm/shared'

export type ProviderKind = 'local' | 'cloud'

export interface ProviderAdapter {
  readonly id: string
  readonly name: string
  readonly kind: ProviderKind
  readonly capabilities: ProviderCapabilities
  isReachable(): Promise<boolean>
  isConfigured(): boolean
  listModels(): Promise<ModelInfo[]>
  stream(req: ChatRequest, signal: AbortSignal): AsyncIterable<StreamEvent>
}

export type ProviderErrorCode =
  | 'auth'
  | 'rate_limit'
  | 'context_length'
  | 'model_not_found'
  | 'network'
  | 'cancelled'
  | 'bad_request'
  | 'upstream'
  | 'unknown'

export class ProviderError extends Error {
  readonly code: ProviderErrorCode
  readonly retryable: boolean
  readonly statusCode?: number
  constructor(
    code: ProviderErrorCode,
    message: string,
    opts: { retryable?: boolean; statusCode?: number; cause?: unknown } = {},
  ) {
    super(message, { cause: opts.cause })
    this.name = 'ProviderError'
    this.code = code
    this.retryable = opts.retryable ?? false
    this.statusCode = opts.statusCode
  }
}

export class AuthError extends ProviderError {
  constructor(message: string, opts: { statusCode?: number; cause?: unknown } = {}) {
    super('auth', message, { statusCode: 401, cause: opts.cause })
    this.name = 'AuthError'
  }
}

export class RateLimitError extends ProviderError {
  constructor(message: string, opts: { statusCode?: number; cause?: unknown } = {}) {
    super('rate_limit', message, { retryable: true, statusCode: 429, cause: opts.cause })
    this.name = 'RateLimitError'
  }
}

export class ContextLengthError extends ProviderError {
  constructor(message: string, opts: { cause?: unknown } = {}) {
    super('context_length', message, { cause: opts.cause })
    this.name = 'ContextLengthError'
  }
}

export class ModelNotFoundError extends ProviderError {
  constructor(message: string, opts: { cause?: unknown } = {}) {
    super('model_not_found', message, { cause: opts.cause })
    this.name = 'ModelNotFoundError'
  }
}

export class NetworkError extends ProviderError {
  constructor(message: string, opts: { cause?: unknown } = {}) {
    super('network', message, { retryable: true, cause: opts.cause })
    this.name = 'NetworkError'
  }
}

export class CancelledError extends ProviderError {
  constructor(message = 'request cancelled') {
    super('cancelled', message, { retryable: false })
    this.name = 'CancelledError'
  }
}

export function classifyHttpStatus(status: number, bodyText: string): ProviderError {
  if (status === 401 || status === 403) return new AuthError(`upstream auth failed (${status})`)
  if (status === 404)
    return new ModelNotFoundError(`upstream model not found (${status}): ${bodyText.slice(0, 200)}`)
  if (status === 429) return new RateLimitError(`upstream rate limit (${status})`)
  if (status >= 400 && status < 500) {
    const msg = bodyText.match(/context.*(length|window|long)/i)
      ? `context length exceeded (${status})`
      : `bad request (${status}): ${bodyText.slice(0, 200)}`
    if (/context/i.test(bodyText)) return new ContextLengthError(msg)
    return new ProviderError('bad_request', msg, { statusCode: status })
  }
  if (status >= 500) {
    return new ProviderError('upstream', `upstream error (${status}): ${bodyText.slice(0, 200)}`, {
      retryable: true,
      statusCode: status,
    })
  }
  return new ProviderError('unknown', `unexpected status ${status}`)
}

export function toStreamEvent(err: unknown): StreamEvent {
  if (err instanceof ProviderError) {
    return {
      type: 'error',
      code: err.code,
      message: err.message,
      retryable: err.retryable,
    }
  }
  if (err instanceof Error) {
    const name = err.name.toLowerCase()
    if (name.includes('abort')) {
      return { type: 'error', code: 'cancelled', message: 'cancelled', retryable: false }
    }
    return { type: 'error', code: 'network', message: err.message, retryable: true }
  }
  return { type: 'error', code: 'unknown', message: String(err) }
}

export type { ChatRequest, ModelInfo, ProviderCapabilities, StreamEvent }
