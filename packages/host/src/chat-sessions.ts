import crypto from 'node:crypto'
import type { StreamEvent } from '@openofficellm/shared'

interface InFlight {
  controller: AbortController
  requestId: string
  startedAt: number
}

const inflight = new Map<string, InFlight>()

export function newRequest(): { requestId: string; controller: AbortController } {
  const requestId = crypto.randomUUID()
  const controller = new AbortController()
  inflight.set(requestId, { controller, requestId, startedAt: Date.now() })
  return { requestId, controller }
}

export function cancelRequest(requestId: string): boolean {
  const entry = inflight.get(requestId)
  if (!entry) return false
  try {
    entry.controller.abort()
  } catch {
    // ignore
  }
  return true
}

export function completeRequest(requestId: string): void {
  inflight.delete(requestId)
}

export function activeRequestCount(): number {
  return inflight.size
}

export function getRequest(requestId: string): InFlight | undefined {
  return inflight.get(requestId)
}

export async function* wrapStream(
  source: AsyncIterable<StreamEvent>,
  requestId: string,
  controller: AbortController,
): AsyncIterable<StreamEvent> {
  try {
    for await (const ev of source) {
      if (controller.signal.aborted) return
      yield ev
      if (ev.type === 'done' || ev.type === 'error') return
    }
  } finally {
    // The InFlight entry is owned by wrapStream — complete it on every exit
    // path: normal completion, error, abort, or client disconnect. The
    // server.ts handler also calls completeRequest in its finally as a
    // belt-and-braces idempotent guard.
    completeRequest(requestId)
  }
}
