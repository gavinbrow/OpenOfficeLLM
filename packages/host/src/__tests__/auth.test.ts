import { describe, it, expect } from 'vitest'
import { createAuthToken, checkToken, extractBearer, constantTimeEqual } from '../auth.js'

describe('auth', () => {
  it('createAuthToken produces a base64url string of ~43 chars', () => {
    const { token, issuedAt } = createAuthToken()
    expect(typeof token).toBe('string')
    expect(token.length).toBeGreaterThanOrEqual(40)
    expect(token).toMatch(/^[A-Za-z0-9_-]+$/)
    expect(issuedAt).toBeGreaterThan(0)
  })

  it('tokens are unique across calls', () => {
    const a = createAuthToken().token
    const b = createAuthToken().token
    expect(a).not.toBe(b)
  })

  it('extractBearer parses a Bearer header', () => {
    expect(extractBearer('Bearer abc123')).toBe('abc123')
    expect(extractBearer('bearer abc123')).toBe('abc123')
    expect(extractBearer('Bearer  abc123 ')).toBe('abc123')
  })

  it('extractBearer rejects non-Bearer schemes', () => {
    expect(extractBearer('Basic abc123')).toBeNull()
    expect(extractBearer('')).toBeNull()
    expect(extractBearer(undefined)).toBeNull()
    expect(extractBearer(null)).toBeNull()
  })

  it('checkToken accepts a matching token', () => {
    const { token } = createAuthToken()
    expect(checkToken(`Bearer ${token}`, token)).toBe(true)
  })

  it('checkToken rejects a wrong token', () => {
    const { token } = createAuthToken()
    expect(checkToken('Bearer wrongtoken', token)).toBe(false)
  })

  it('checkToken rejects missing header', () => {
    const { token } = createAuthToken()
    expect(checkToken(undefined, token)).toBe(false)
    expect(checkToken(null, token)).toBe(false)
    expect(checkToken('', token)).toBe(false)
  })

  it('constantTimeEqual handles equal strings', () => {
    expect(constantTimeEqual('abc', 'abc')).toBe(true)
  })

  it('constantTimeEqual handles unequal strings', () => {
    expect(constantTimeEqual('abc', 'abd')).toBe(false)
    expect(constantTimeEqual('abc', 'abcd')).toBe(false)
    expect(constantTimeEqual('abcd', 'abc')).toBe(false)
  })
})
