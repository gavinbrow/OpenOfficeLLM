import { describe, it, expect } from 'vitest'
import { DEFAULT_SETTINGS, migrateStepCap } from '../index'

describe('migrateStepCap', () => {
  it('gives the new default when nothing was stored', () => {
    expect(migrateStepCap(undefined)).toBe(DEFAULT_SETTINGS.agenticStepCap)
    expect(DEFAULT_SETTINGS.agenticStepCap).toBe(30)
  })

  it('raises the superseded default rather than leaving the old ceiling in place', () => {
    expect(migrateStepCap(10)).toBe(30)
  })

  it('leaves a cap the user actually chose alone, above or below the default', () => {
    expect(migrateStepCap(5)).toBe(5)
    expect(migrateStepCap(12)).toBe(12)
    expect(migrateStepCap(60)).toBe(60)
  })

  it('falls back to the default for values that are not usable numbers', () => {
    expect(migrateStepCap('30')).toBe(30)
    expect(migrateStepCap(null)).toBe(30)
    expect(migrateStepCap(Number.NaN)).toBe(30)
  })
})
