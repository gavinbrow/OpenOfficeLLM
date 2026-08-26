import { describe, it, expect } from 'vitest'
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { HOST_VERSION } from '../paths.js'
import { MANIFEST_VERSION, manifestVersionFor } from '../manifest.js'

const here = path.dirname(fileURLToPath(import.meta.url))
const rootPkg = path.resolve(here, '..', '..', '..', '..', 'package.json')
const pkg = JSON.parse(fs.readFileSync(rootPkg, 'utf8')) as { version: string }

describe('HOST_VERSION drift guard', () => {
  it('matches the root package.json version', () => {
    expect(HOST_VERSION).toBe(pkg.version)
  })
})

describe('manifest <Version>', () => {
  it('tracks the app version instead of a fixed constant', () => {
    expect(MANIFEST_VERSION).toBe(manifestVersionFor(HOST_VERSION))
  })

  // Office refuses a manifest below 1.0 outright ("Manifest Version Too Low"),
  // with no error anywhere in the UI — see the note in manifest.ts.
  it('never emits a version below 1.0', () => {
    for (const v of ['0.0.1', '0.1.2', '1.0.0', '2.10.3']) {
      expect(Number.parseInt(manifestVersionFor(v).split('.')[0], 10)).toBeGreaterThanOrEqual(1)
    }
  })

  it('orders monotonically across the 0.x to 1.0 boundary', () => {
    const rank = (v: string) =>
      manifestVersionFor(v)
        .split('.')
        .map((n) => Number.parseInt(n, 10))
    const ascending = ['0.1.2', '0.2.0', '0.9.9', '1.0.0', '1.0.1', '2.0.0']
    for (let i = 1; i < ascending.length; i++) {
      const prev = rank(ascending[i - 1])
      const next = rank(ascending[i])
      const cmp = prev.findIndex((n, j) => n !== next[j])
      expect(cmp).toBeGreaterThanOrEqual(0)
      expect(next[cmp]).toBeGreaterThan(prev[cmp])
    }
  })
})
