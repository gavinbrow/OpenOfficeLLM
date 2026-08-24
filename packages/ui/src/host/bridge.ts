// The seam between the chat UI and whatever surface it is running on.
//
// `packages/ui` renders the assistant; it must not know whether the thing on
// the other side is a Word task pane or a Chrome side panel. Each shell builds
// a Shell object at startup and registers it here before the first render, and
// every host-dependent thing the UI needs — the document adapter, the tool
// catalog, the per-document identity used to keep chats apart — comes through
// this one object.
//
// Registration is a function call rather than React context on purpose: the
// Zustand stores need the shell too, and they are plain modules with no access
// to the component tree.
//
// What deliberately does NOT go through here is per-host wording — suggestion
// chips, banner copy, the name of the thing being edited. Those switch on the
// host union in the components. The union lives in `shared`, so the UI knowing
// all three members costs nothing, whereas routing every string through the
// shell would mean two shells to edit for one label change.

import type {
  DetectedHost,
  Edit,
  EditMode,
  HostKind,
  TextFormatting,
  ToolDefinition,
} from '@openofficellm/shared'
import type { HostAdapter } from './adapter'

export interface ToolOutcome {
  content: string
  isError: boolean
}

export interface ToolExecContext {
  adapter: HostAdapter
  mode: EditMode
  /** Stage an edit for review instead of applying it. Returns the description
   *  shown in the proposal card. */
  propose: (edit: Edit, description: string) => void
}

export interface Shell {
  /** Which family of shell this is. Both can report a `'none'` host — the pane
   *  opened in a plain browser, a side panel with no eligible tab — and the UI
   *  needs to tell those two situations apart to explain them. */
  readonly surface: 'office' | 'browser'

  /** Which surface this is. `'none'` means the UI is running detached and
   *  every document tool is withheld. */
  getHost(): DetectedHost

  /** The document adapter, or null when there is nothing to act on. */
  getAdapter(): HostAdapter | null

  /** Stable identity for the thing being worked on, used to keep each
   *  document's or tab's chat separate. Storage is scoped to the shell's
   *  origin, not to the document, so without this every surface reopens
   *  whichever chat was used last. */
  getDocumentKey(): string

  /** Dark/light as the surrounding application reports it. */
  isDark(): boolean

  /** Tools the model may call on this host. `allowWrites` is false when the
   *  shell must not offer to change anything this turn. */
  toolCatalog(host: DetectedHost, allowWrites: boolean): ToolDefinition[]

  /** True if calling `name` would change the document. Drives the propose /
   *  direct split, so a shell that gets this wrong applies edits the user
   *  never approved. */
  isWriteTool(name: string): boolean

  /** `argsJson` is the raw string the model emitted, not a parsed object —
   *  parsing is the shell's job because only it knows which coercions its own
   *  tools tolerate. */
  executeDocumentTool(name: string, argsJson: string, ctx: ToolExecContext): Promise<ToolOutcome>

  /** Render a formatting payload in the user's terms for the proposal card.
   *  Optional — a shell with no rich formatting vocabulary gets the generic
   *  key/value rendering, which is plainer but not wrong. */
  describeFormatting?(formatting: TextFormatting): string
}

let _shell: Shell | null = null

export function registerShell(shell: Shell): void {
  _shell = shell
}

/**
 * The registered shell.
 *
 * Throwing when nothing is registered is deliberate. The alternative — a
 * silent no-op shell — turns "the entrypoint forgot to call registerShell"
 * into a UI that renders fine and quietly refuses to touch the document,
 * which is far harder to diagnose than a stack trace on the first render.
 */
export function shell(): Shell {
  if (_shell === null) {
    throw new Error(
      'No shell registered. The entrypoint must call registerShell() before rendering.',
    )
  }
  return _shell
}

/** Test seam, and for shells that tear down between documents. */
export function __resetShellForTest(): void {
  _shell = null
}

// ─── Convenience accessors ───────────────────────────────────────────────
// Thin wrappers so call sites read as `getHost()` rather than
// `shell().getHost()`. These three are used almost everywhere, and unlike
// `shell()` they tolerate an unregistered shell: components can render during
// a test that never bootstrapped one, they just render the detached state.

export function getSurface(): 'office' | 'browser' {
  return _shell?.surface ?? 'office'
}

export function getHost(): DetectedHost {
  return _shell?.getHost() ?? 'none'
}

export function getAdapter(): HostAdapter | null {
  return _shell?.getAdapter() ?? null
}

export function getDocumentKey(): string {
  return _shell?.getDocumentKey() ?? 'detached'
}

export function isDark(): boolean {
  return _shell?.isDark() ?? false
}

/**
 * The host key to read per-host settings under.
 *
 * `Settings.defaultMode` and `defaultContext` are keyed by real host, but the
 * UI still has to render a mode picker and a scope picker when nothing is
 * attached. Falling back to the shell's primary host keeps those controls
 * meaningful — and keeps whatever the user picks while detached from being
 * written into the wrong host's slot.
 */
export function settingsHost(host: DetectedHost = getHost()): HostKind {
  if (host !== 'none') return host
  return getSurface() === 'browser' ? 'browser' : 'word'
}
