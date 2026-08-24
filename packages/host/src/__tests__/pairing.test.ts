// The extension allowlist is the whole trust boundary for the browser shell:
// an origin on it can read the auth token, which is a key to every configured
// provider. So the tests here are mostly about what must NOT be admitted.

import { describe, it, expect } from 'vitest'
import {
  addPairing,
  extensionOrigin,
  isExtensionId,
  isPairedExtension,
  originToExtensionId,
  removePairing,
} from '../pairing.js'

const ID = 'abcdefghijklmnopabcdefghijklmnop'
const ORIGIN = `chrome-extension://${ID}`

describe('extension ids', () => {
  it('accepts a well-formed id', () => {
    expect(isExtensionId(ID)).toBe(true)
    expect(extensionOrigin(ID)).toBe(ORIGIN)
  })

  it('rejects ids that are the wrong length or alphabet', () => {
    expect(isExtensionId('abc')).toBe(false)
    expect(isExtensionId(`${ID}a`)).toBe(false)
    // Chrome ids use a–p only; q–z and digits are not extension ids.
    expect(isExtensionId('zbcdefghijklmnopabcdefghijklmnop')).toBe(false)
    expect(isExtensionId('1bcdefghijklmnopabcdefghijklmnop')).toBe(false)
    expect(extensionOrigin('nope')).toBeNull()
  })

  it('round-trips an origin back to its id', () => {
    expect(originToExtensionId(ORIGIN)).toBe(ID)
    expect(originToExtensionId('https://evil.example')).toBeNull()
    // A path makes it a URL, not an origin.
    expect(originToExtensionId(`${ORIGIN}/sidepanel.html`)).toBeNull()
  })
})

describe('isPairedExtension', () => {
  it('admits an exactly-matching paired origin', () => {
    expect(isPairedExtension(ORIGIN, [ORIGIN])).toBe(true)
  })

  it('refuses everything that is not on the list', () => {
    expect(isPairedExtension(ORIGIN, [])).toBe(false)
    expect(isPairedExtension(undefined, [ORIGIN])).toBe(false)
    expect(isPairedExtension('', [ORIGIN])).toBe(false)
  })

  it('refuses a different extension, including one with a similar id', () => {
    const other = `chrome-extension://${'p'.repeat(32)}`
    expect(isPairedExtension(other, [ORIGIN])).toBe(false)
  })

  it('refuses a web origin even when it is somehow on the list', () => {
    // A malformed config entry must not become a way in for a web page.
    expect(isPairedExtension('https://evil.example', ['https://evil.example'])).toBe(false)
  })

  it('refuses a prefix or suffix of a paired origin', () => {
    expect(isPairedExtension(`${ORIGIN}.evil.example`, [ORIGIN])).toBe(false)
    expect(isPairedExtension(`${ORIGIN}/`, [ORIGIN])).toBe(false)
  })

  it('trims surrounding whitespace but does not otherwise normalise', () => {
    expect(isPairedExtension(`  ${ORIGIN}  `, [ORIGIN])).toBe(true)
    expect(isPairedExtension(ORIGIN.toUpperCase(), [ORIGIN])).toBe(false)
  })
})

describe('addPairing / removePairing', () => {
  it('adds a valid id once and reports the duplicate', () => {
    const first = addPairing([], ID)
    expect(first).toMatchObject({ origins: [ORIGIN], added: true })
    expect(addPairing(first.origins, ID)).toMatchObject({ origins: [ORIGIN], added: false })
  })

  it('refuses a malformed id and leaves the list untouched', () => {
    const r = addPairing([ORIGIN], 'not-an-id')
    expect(r.added).toBe(false)
    expect(r.error).toMatch(/not a Chrome extension id/)
    expect(r.origins).toEqual([ORIGIN])
  })

  it('removes by id or by full origin', () => {
    expect(removePairing([ORIGIN], ID)).toMatchObject({ origins: [], removed: true })
    expect(removePairing([ORIGIN], ORIGIN)).toMatchObject({ origins: [], removed: true })
    expect(removePairing([ORIGIN], 'other')).toMatchObject({ origins: [ORIGIN], removed: false })
  })
})
