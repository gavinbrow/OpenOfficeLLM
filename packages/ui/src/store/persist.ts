// Small localStorage helpers used by the stores. Never stores secrets.
//
// `loadPersisted` validates array/object shapes so a corrupted or adversarial
// localStorage entry (e.g. `conversations: [{messages: "not-an-array"}]`)
// falls back to the default instead of crashing the pane. Without this guard,
// a malformed entry would throw inside a `.map` call deep in a component,
// surface only via the ErrorBoundary, and lose the user's session.

const PREFIX = 'openofficellm:'

export function loadPersisted<T>(key: string, fallback: T, validate?: (v: unknown) => v is T): T {
  if (typeof localStorage === 'undefined') return fallback
  try {
    const raw = localStorage.getItem(PREFIX + key)
    if (!raw) return fallback
    const parsed = JSON.parse(raw) as unknown
    if (validate) {
      return validate(parsed) ? parsed : fallback
    }
    return parsed as T
  } catch {
    return fallback
  }
}

export function savePersisted<T>(key: string, value: T): void {
  if (typeof localStorage === 'undefined') return
  try {
    localStorage.setItem(PREFIX + key, JSON.stringify(value))
  } catch {
    // quota or disabled storage — ignore
  }
}

export function removePersisted(key: string): void {
  if (typeof localStorage === 'undefined') return
  try {
    localStorage.removeItem(PREFIX + key)
  } catch {
    // ignore
  }
}

/** Type guard for arrays — used by stores loading conversation/message lists. */
export function isArray(v: unknown): v is unknown[] {
  return Array.isArray(v)
}
