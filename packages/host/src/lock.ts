import fs from 'node:fs'
import process from 'node:process'
import { resolveLockPath, ensureDirs } from './paths.js'

interface LockState {
  pid: number
  startedAt: number
}

function isPidAlive(pid: number): boolean {
  try {
    process.kill(pid, 0)
    return true
  } catch (e) {
    const code = (e as NodeJS.ErrnoException).code
    if (code === 'ESRCH') return false
    if (code === 'EPERM') return true
    return false
  }
}

function readLock(): LockState | null {
  try {
    const raw = fs.readFileSync(resolveLockPath(), 'utf8').trim()
    const parsed = JSON.parse(raw) as LockState
    if (typeof parsed.pid !== 'number') return null
    return parsed
  } catch {
    return null
  }
}

export function acquireLock(): boolean {
  ensureDirs()
  const lockPath = resolveLockPath()
  const existing = readLock()
  if (existing) {
    if (isPidAlive(existing.pid)) {
      return false
    }
    try {
      fs.unlinkSync(lockPath)
    } catch {
      // ignore — stale lock we couldn't delete; we'll overwrite below
    }
  }
  const state: LockState = { pid: process.pid, startedAt: Date.now() }
  try {
    fs.writeFileSync(lockPath, JSON.stringify(state), { flag: 'wx' })
    return true
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code === 'EEXIST') {
      const again = readLock()
      if (again && isPidAlive(again.pid)) return false
      try {
        fs.writeFileSync(lockPath, JSON.stringify(state))
        return true
      } catch {
        return false
      }
    }
    return false
  }
}

export function releaseLock(): void {
  try {
    fs.unlinkSync(resolveLockPath())
  } catch {
    // already gone
  }
}

export function getLockHolder(): LockState | null {
  return readLock()
}

export { isPidAlive }
