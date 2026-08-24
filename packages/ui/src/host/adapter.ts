// The HostAdapter contract (P4.1).
//
// Everything that touches a concrete document API lives behind this interface,
// so the chat loop, the tool dispatcher, and the edit-mode machinery are all
// host-agnostic. The Word, Excel and browser-page implementations are the only
// files that know the difference between a paragraph, a cell and a DOM node.

import type {
  ContextScope,
  HostKind,
  DocumentContext,
  Edit,
  Snapshot,
  TextTarget,
} from '@openofficellm/shared'

export interface SnapshotPayload extends Snapshot {
  /** Host-specific restore data. Word: OOXML. Excel: values + formulas per
   *  touched sheet. Never leaves the pane — the host service only ever sees the
   *  id and the size. */
  data: unknown
}

export interface ApplyResult {
  /** Human-readable confirmation, handed back to the model as the tool result. */
  summary: string
  /** False when the edit was rejected (e.g. a range that no longer exists). */
  ok: boolean
}

export interface HostAdapter {
  readonly host: HostKind

  /** Read document context for a scope. Never throws for an empty document —
   *  returns a context with empty text so the caller can distinguish "empty"
   *  from "failed". */
  getContext(scope: ContextScope): Promise<DocumentContext>

  /** Apply edits as a single batch so one Ctrl+Z reverts the whole change.
   *
   *  Optional: a host that owns no document cannot honour this. The browser
   *  adapter omits it — a web page belongs to whoever served it, and an
   *  assistant that rewrites it has changed nothing the user can save and
   *  everything they might screenshot. Read tools only there. */
  applyEdits?(edits: Edit[]): Promise<ApplyResult>

  /** Capture enough state to undo a run. Omitted by hosts that cannot edit. */
  snapshot?(): Promise<SnapshotPayload>

  /** Restore a snapshot captured by this adapter. */
  restore?(snapshot: SnapshotPayload): Promise<void>

  /** Free-text search, used by the search_document tool. */
  search(query: string, limit?: number): Promise<SearchHit[]>

  /** Describe the formatting currently in force on a target. Word only —
   *  the tool that calls it is withheld from hosts that do not implement it. */
  readFormatting?(target: TextTarget): Promise<string>
}

export interface SearchHit {
  /** Where the hit is, in host terms: a paragraph index or a cell address. */
  location: string
  text: string
}

/** Scopes that make sense for a host, in the order the UI offers them. */
export const SCOPES_FOR_HOST: Record<HostKind, ContextScope[]> = {
  word: ['none', 'selection', 'paragraph', 'document'],
  excel: ['none', 'range', 'sheet'],
  browser: ['none', 'selection', 'page'],
}

/** Rough token estimate for context sizing. Deliberately conservative: an
 *  underestimate means we send a payload the model rejects, which costs the
 *  user a round trip; an overestimate only costs a warning they can ignore. */
export function estimateTokens(text: string): number {
  return Math.ceil(text.length / 3.5)
}
