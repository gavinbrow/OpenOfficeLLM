import { describe, it, expect } from 'vitest'
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { HOST_VERSION } from '../paths.js'

const here = path.dirname(fileURLToPath(import.meta.url))
const rootPkg = path.resolve(here, '..', '..', '..', '..', 'package.json')
const pkg = JSON.parse(fs.readFileSync(rootPkg, 'utf8')) as { version: string }

describe('HOST_VERSION drift guard', () => {
  it('matches the root package.json version', () => {
    expect(HOST_VERSION).toBe(pkg.version)
  })
})