// Letting a browser extension talk to the host service.
//
// The task pane needs none of this: it is *served by* the host, so it is
// same-origin and the host injects the bearer token straight into the HTML it
// hands over. An extension is served from its own `chrome-extension://<id>`
// origin, so it is cross-origin by construction and there is no HTML for the
// host to stamp a token into.
//
// What replaces those two mechanisms is an explicit allowlist. The user pairs
// an extension id once; from then on that origin — and no other — may read the
// token and call the API. The allowlist is the entire security boundary here,
// so the matching is exact and deliberately unclever: no wildcards, no
// prefixes, no case folding beyond the id itself.

/** A Chrome/Edge extension id: 32 lowercase letters a–p. Firefox uses a UUID
 *  in braces, which this deliberately does not accept — the pairing flow is
 *  only implemented for Chromium browsers so far, and silently admitting an
 *  origin shape we have not thought about is how allowlists rot. */
const EXTENSION_ID_RE = /^[a-p]{32}$/

const EXTENSION_ORIGIN_RE = /^chrome-extension:\/\/([a-p]{32})$/

export function isExtensionId(value: string): boolean {
  return EXTENSION_ID_RE.test(value.trim())
}

/** The canonical origin for an extension id, or null if the id is malformed. */
export function extensionOrigin(id: string): string | null {
  const trimmed = id.trim()
  return isExtensionId(trimmed) ? `chrome-extension://${trimmed}` : null
}

/** Parse an origin back to an id, for display and for de-duplication. */
export function originToExtensionId(origin: string): string | null {
  const m = EXTENSION_ORIGIN_RE.exec(origin.trim())
  return m ? m[1] : null
}

/**
 * Whether a request's Origin header names a paired extension.
 *
 * Exact string equality against the stored list. An extension origin has no
 * path, port or subdomain component, so there is nothing to normalise and
 * nothing a lenient comparison would buy except a way to be wrong.
 */
export function isPairedExtension(
  origin: string | undefined | null,
  paired: readonly string[],
): boolean {
  if (!origin) return false
  const trimmed = origin.trim()
  if (originToExtensionId(trimmed) === null) return false
  return paired.includes(trimmed)
}

/** Add an id to a list, returning the new list and whether anything changed. */
export function addPairing(
  paired: readonly string[],
  id: string,
): { origins: string[]; added: boolean; error?: string } {
  const origin = extensionOrigin(id)
  if (!origin) {
    return {
      origins: [...paired],
      added: false,
      error: `"${id}" is not a Chrome extension id (expected 32 letters a–p).`,
    }
  }
  if (paired.includes(origin)) return { origins: [...paired], added: false }
  return { origins: [...paired, origin], added: true }
}

export function removePairing(
  paired: readonly string[],
  id: string,
): { origins: string[]; removed: boolean } {
  const origin = extensionOrigin(id) ?? id.trim()
  const origins = paired.filter((o) => o !== origin)
  return { origins, removed: origins.length !== paired.length }
}
