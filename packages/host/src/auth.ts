import crypto from 'node:crypto'

const TOKEN_BYTES = 32

export interface AuthState {
  token: string
  issuedAt: number
}

export function createAuthToken(): AuthState {
  const bytes = crypto.randomBytes(TOKEN_BYTES)
  const token = bytes.toString('base64url')
  return { token, issuedAt: Date.now() }
}

export function constantTimeEqual(a: string, b: string): boolean {
  const ab = Buffer.from(a, 'utf8')
  const bb = Buffer.from(b, 'utf8')
  if (ab.length !== bb.length) return false
  return crypto.timingSafeEqual(ab, bb)
}

export function extractBearer(authHeader: string | undefined | null): string | null {
  if (!authHeader) return null
  const trimmed = authHeader.trim()
  const m = /^Bearer\s+(.+)$/i.exec(trimmed)
  if (!m) return null
  return m[1].trim()
}

export function checkToken(authHeader: string | undefined | null, expected: string): boolean {
  const presented = extractBearer(authHeader)
  if (!presented) return false
  return constantTimeEqual(presented, expected)
}
