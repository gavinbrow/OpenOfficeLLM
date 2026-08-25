// OpenOfficeLLM shared protocol types.
//
// Zero runtime dependencies. Consumed by `packages/host` (Node) and
// `packages/addin` (browser). Keep everything wire-format-stable: a change
// here is a change to the host service's API contract.

// ─── Chat ────────────────────────────────────────────────────────────────

export type UserRole = 'user' | 'assistant' | 'system' | 'tool'

export interface ChatMessage {
  role: UserRole
  content: string
  /** Chain-of-thought the model emitted alongside `content`. Kept separate so
   *  it never renders as the answer, and never goes back upstream on the next
   *  turn — reasoning is a per-turn artifact, not conversation history. */
  reasoning?: string
  /** Present on assistant messages that produced a tool call, or on tool
   *  messages that carry the result. */
  toolCalls?: ToolCall[]
  /** Only on role === 'tool' messages. */
  toolCallId?: string
  /** Only on role === 'tool' messages. Ollama keys tool results by tool name
   *  rather than by call id, so both have to survive the round trip. */
  toolName?: string
  /** Optional per-message id, stable across streaming so the UI can update
   *  a partial assistant message in place. */
  id?: string
  /** Model that produced this assistant message (host echoes it back). */
  model?: string
  /** Unix ms timestamp. */
  createdAt?: number
}

/** A tool the model may call, in the provider-neutral shape every adapter
 *  translates from. `parameters` is a JSON Schema object. */
export interface ToolDefinition {
  name: string
  description: string
  parameters: {
    type: 'object'
    properties: Record<string, unknown>
    required?: string[]
  }
}

export interface ToolCall {
  id: string
  /** Namespaced tool name. Document tools are bare (`replace_text`); MCP tools
   *  are namespaced `mcp__<server>__<tool>`. */
  name: string
  /** JSON-encoded arguments. */
  arguments: string
}

/** A surface the assistant can read and act on.
 *
 *  Word and Excel are the Office task pane; `browser` is the Chrome extension
 *  acting on the active tab. They differ enormously in what they can write —
 *  a browser tab is somebody else's document — but they answer the same
 *  questions about scope, context and tool dispatch, which is what this union
 *  is for. */
export type HostKind = 'word' | 'excel' | 'browser'

/** A host that may not have been detected yet, or may not exist at all — the
 *  pane opened in a plain browser during development, or an extension side
 *  panel with no eligible tab in front of it. */
export type DetectedHost = HostKind | 'none'

export type ContextScope =
  | 'none'
  | 'selection'
  | 'paragraph'
  | 'document'
  | 'sheet'
  | 'range'
  /** Browser: the readable content of the active tab. */
  | 'page'

export interface ChatRequest {
  /** Stable conversation id, scoped per document where practical. */
  conversationId?: string
  messages: ChatMessage[]
  /** `<providerId>/<modelId>` — the normalized model id from /api/models. */
  model: string
  mode: EditMode
  /** What the user has chosen to include as document context for this turn. */
  context?: DocumentContext
  /** Tools the model is allowed to call this turn. The pane sends the document
   *  tool catalog for the active host; the host service unions in the MCP tools
   *  the user has enabled. */
  tools?: ToolDefinition[]
  /** If set, invokes a skill's system-prompt template. */
  skillId?: string
  /** Optional system prompt appended after the skill template. */
  systemPrompt?: string
  /** Sampling parameters; providers map to their own shapes. */
  temperature?: number
  maxTokens?: number
  /** Host assigns and echoes back on the first event, so the pane can cancel. */
  requestId?: string
}

// ─── Streaming ───────────────────────────────────────────────────────────

export type StreamEvent =
  | { type: 'start'; requestId: string; model: string }
  | { type: 'delta'; text: string }
  /** Chain-of-thought. A separate event so the pane can render it collapsed
   *  instead of splicing it into the answer. */
  | { type: 'reasoning'; text: string }
  | { type: 'tool_call'; toolCall: ToolCall }
  | { type: 'tool_result'; toolCallId: string; result: string; isError?: boolean }
  | { type: 'usage'; promptTokens: number; completionTokens: number; totalTokens: number }
  | { type: 'cost'; estimatedCostUsd: number }
  | { type: 'done'; requestId: string; finishReason?: string }
  | { type: 'error'; code: string; message: string; retryable?: boolean }

export const isStreamError = (e: StreamEvent): e is Extract<StreamEvent, { type: 'error' }> =>
  e.type === 'error'

// ─── Models and providers ────────────────────────────────────────────────

export interface ProviderCapabilities {
  tools: boolean
  vision: boolean
  streaming: boolean
}

export interface ProviderInfo {
  id: string
  name: string
  kind: 'local' | 'cloud'
  /** True if the provider is reachable on the network right now. */
  reachable: boolean
  /** True if any secret/credential is configured for this provider. */
  configured: boolean
  capabilities: ProviderCapabilities
}

export interface ModelInfo {
  /** Fully-qualified id: `<providerId>/<modelId>`. */
  id: string
  /** Just the model part, for display. */
  name: string
  providerId: string
  /** Display label for the provider (e.g. "Ollama", "Anthropic"). */
  providerName: string
  kind: 'local' | 'cloud'
  capabilities: ProviderCapabilities
  /** Approximate context window in tokens, if known. */
  contextWindow?: number
  /** For local models, size on disk in bytes. */
  sizeBytes?: number
  /** For local models, quantization label e.g. "Q4_K_M". */
  quantization?: string
  /** Static price per 1M tokens, if known. */
  inputPricePer1k?: number
  outputPricePer1k?: number
}

// ─── Document context and edits ──────────────────────────────────────────

export interface DocumentContext {
  host: DetectedHost
  scope: ContextScope
  /** Word: selection/paragraph/body text; Excel: serialized range with
   *  sampling applied. */
  text: string
  /** Word: heading outline; Excel: sheet names + used-range schema. */
  outline?: string
  /** Excel-only: column types inferred from sampling. */
  schema?: ColumnSchema[]
  /** Excel-only: row count of the full used range (sampling never sends all rows). */
  totalRows?: number
  /** Approximate token count the host computed for this context, for the
   *  ContextChips UI. */
  tokenEstimate?: number
}

export interface ColumnSchema {
  name: string
  /** Inferred JS type. `null` = mixed/unknown. */
  type: 'string' | 'number' | 'boolean' | 'date' | 'null'
  sampleValues: string[]
}

export type EditMode = 'propose' | 'direct' | 'agentic'

export type Edit =
  | { kind: 'replaceSelection'; text: string }
  | { kind: 'insertAfter'; text: string }
  | { kind: 'insertBefore'; text: string }
  | { kind: 'replaceRange'; rangeId: string; text: string }
  | { kind: 'setCellValues'; sheet: string; cells: CellValue[] }
  | { kind: 'setCellFormulas'; sheet: string; cells: CellFormula[] }
  | { kind: 'applyFormatting'; rangeId: string; formatting: Record<string, unknown> }
  | { kind: 'addComment'; rangeId: string; text: string }
  // ─ Word formatting and structure ─
  | { kind: 'formatText'; target: TextTarget; formatting: TextFormatting }
  | { kind: 'insertParagraph'; at: DocumentAnchor; text: string; style?: string }
  // Paragraph lists rather than ranges: "bullet points 2 and 5" is a normal
  // request, and a from/to pair cannot express it.
  | { kind: 'deleteParagraphs'; paragraphs: number[] }
  | { kind: 'setList'; paragraphs: number[]; listType: ListType; level?: number }
  | {
      kind: 'insertTable'
      at: DocumentAnchor
      rows: string[][]
      headerRow?: boolean
      style?: string
    }
  | { kind: 'insertBreak'; at: DocumentAnchor; breakType: BreakKind }
  | { kind: 'insertHyperlink'; target: TextTarget; url: string; text?: string }
  | { kind: 'replaceAll'; find: string; replace: string; matchCase?: boolean; wholeWord?: boolean }
  | {
      kind: 'setHeaderFooter'
      part: 'header' | 'footer'
      text: string
      alignment?: TextAlignment
      pageNumber?: boolean
    }
  | { kind: 'setPageSetup'; setup: PageSetup }

/**
 * What a Word text operation applies to.
 *
 * `selection` is the default because it matches what the user is looking at.
 * `paragraphs` carries indices the model got from read_document or
 * search_document; `search` lets it format a phrase without knowing where the
 * phrase lives, which is the only way to reach a run of text inside a paragraph.
 */
export interface TextTarget {
  kind: 'selection' | 'document' | 'paragraphs' | 'search'
  /** `paragraphs` only: zero-based indices, already expanded from any range. */
  paragraphs?: number[]
  /** `search` only: the literal text to match. */
  search?: string
  matchCase?: boolean
  wholeWord?: boolean
  /** `search` only: apply to the first hit rather than every hit. */
  firstOnly?: boolean
}

/** Where an insertion goes. A number is a zero-based paragraph index, and the
 *  new content lands after that paragraph. */
export type DocumentAnchor = number | 'start' | 'end' | 'selection'

export type TextAlignment = 'left' | 'center' | 'right' | 'justify'
export type ListType = 'bullet' | 'number' | 'none'
export type BreakKind = 'page' | 'line' | 'section' | 'sectionContinuous'

/**
 * Character- and paragraph-level formatting.
 *
 * Every field is optional and an omitted field is left alone — there is no way
 * to express "reset to default" other than naming the value you want, which is
 * deliberate: a formatter that clears unmentioned properties would wipe a
 * document's styling on the first bold request.
 */
export interface TextFormatting {
  // Character
  bold?: boolean
  italic?: boolean
  /** `true`/`false` for a single underline, or a named style such as "double". */
  underline?: boolean | string
  strikeThrough?: boolean
  doubleStrikeThrough?: boolean
  superscript?: boolean
  subscript?: boolean
  smallCaps?: boolean
  allCaps?: boolean
  /** `#RRGGBB` or a colour name. */
  color?: string
  /** `#RRGGBB`, a colour name, or `null` to clear the highlight. */
  highlightColor?: string | null
  /** Points. */
  size?: number
  /** Font family name, e.g. "Calibri". */
  font?: string

  // Paragraph
  /** Built-in or custom style name, e.g. "Heading 1". */
  style?: string
  alignment?: TextAlignment
  /** Points between lines. */
  lineSpacing?: number
  /** Points of space before/after the paragraph. */
  spaceBefore?: number
  spaceAfter?: number
  /** Points of indent. */
  leftIndent?: number
  rightIndent?: number
  firstLineIndent?: number
  /** 1–9 for outline levels, 10 for body text. */
  outlineLevel?: number
  // Pagination controls (keep-with-next, widow/orphan) are deliberately absent:
  // Office.js exposes them on styles only, never on an individual paragraph.
}

export interface PageSetup {
  orientation?: 'portrait' | 'landscape'
  /** Inches. Omitted edges keep their current margin. */
  margins?: { top?: number; bottom?: number; left?: number; right?: number }
  pageSize?: 'letter' | 'legal' | 'tabloid' | 'a3' | 'a4' | 'a5'
}

export interface CellValue {
  /** e.g. "A1" or "B12". */
  cell: string
  value: string | number | boolean | null
}

export interface CellFormula {
  cell: string
  formula: string
}

/** Opaque snapshot the host/pane can pass around for revert. The actual
 *  payload is host-specific (OOXML for Word, values+formulas for Excel) and
 *  lives in the pane; the host only sees an id + size. */
export interface Snapshot {
  id: string
  host: HostKind
  /** ISO timestamp. */
  createdAt: string
  /** Approximate byte size, for warnings on large documents. */
  sizeBytes: number
}

// ─── Skills, MCP ─────────────────────────────────────────────────────────

export interface Skill {
  id: string
  name: string
  description: string
  /** Which hosts this skill applies to. Empty = every host. */
  hosts: HostKind[]
  mode?: EditMode
  /** Override the selected model when the skill is invoked. */
  model?: string
  icon?: string
  /** Optional context scope the skill implies. */
  contextScope?: ContextScope
  /** The system-prompt template, with `{{selection}}`, `{{document}}`, etc. */
  prompt?: string
  /** True if the skill ships with the app, false if user-authored. */
  builtIn: boolean
  /** Where the skill came from, for the settings list. */
  source: 'built-in' | 'user' | 'opencode'
  /** Absolute path for user skills, so the editor can write back. Never set
   *  for built-ins or opencode imports (which are read-only). */
  path?: string
}

/** Persisted MCP server definition. Lives in config.json, not in the pane. */
export interface McpServerConfig {
  id: string
  name: string
  transport: 'stdio' | 'http'
  /** stdio only. */
  command?: string
  args?: string[]
  /** stdio only. Merged over the host's own environment. */
  env?: Record<string, string>
  /** http only. */
  url?: string
  /** http only. Sent on every request; values may name a secret to resolve
   *  rather than carrying it inline — see `McpHeaderRef`. */
  headers?: Record<string, string>
  /** False keeps the server defined but never started. */
  enabled: boolean
  /** Whether this server came from an opencode import. */
  imported?: boolean
}

/** Agent/subagent definition imported from opencode or created by the user. */
export interface AgentConfig {
  id: string
  name: string
  description: string
  model: string
  /** Default mode for this agent's runs. */
  mode?: EditMode
  /** Allow the agent to use tools. */
  tools?: boolean
  /** Permission profile, copied from opencode. */
  permissions?: Record<string, string>
  /** Where this agent came from. */
  source: 'opencode' | 'user'
  /** Whether the user has enabled this agent. */
  enabled: boolean
}

/** Provider definition imported from opencode. We keep a lightweight record
 *  so the settings UI can show imported providers and the host can recreate
 *  them on startup. */
export interface ImportedProviderConfig {
  id: string
  name: string
  /** OpenAI-compatible base URL. */
  baseUrl: string
  /** Model aliases / hints. */
  models?: string[]
  /** Whether the user has enabled this provider. */
  enabled: boolean
}

/** Provider-specific runtime options persisted in settings. */
export interface ProviderOptions {
  /** Custom base URL for the local Ollama adapter. */
  ollamaBaseUrl?: string
}

export type McpServerStatus = 'stopped' | 'starting' | 'ready' | 'error'

export interface McpServerInfo extends McpServerConfig {
  status: McpServerStatus
  /** Populated when status === 'error'. Already redacted. */
  error?: string
  tools: McpTool[]
}

export interface McpTool {
  serverId: string
  name: string
  description?: string
  /** JSON-schema of the tool's input. */
  inputSchema?: Record<string, unknown>
  /** Whether the user has opted this specific tool in. */
  enabled: boolean
}

export interface McpCallRequest {
  serverId: string
  tool: string
  arguments: Record<string, unknown>
}

export interface McpCallResponse {
  /** Text rendering of the tool result, ready to hand back to the model. */
  content: string
  isError: boolean
}

// ─── Settings ────────────────────────────────────────────────────────────

export interface Settings {
  /** Default `<providerId>/<modelId>`. */
  defaultModel?: string
  /** Default edit mode, per host. */
  defaultMode: Record<HostKind, EditMode>
  /** Context defaults per host. */
  defaultContext: Record<HostKind, ContextScope>
  /** Theme override; undefined = follow Office. */
  theme?: 'light' | 'dark'
  /** Token threshold for trimming warnings. */
  contextTrimWarningTokens: number
  /** Tool-calling rounds allowed in one turn. Applies to both modes: a
   *  formatting request routinely needs a read, several edits and a re-read,
   *  and a cap that stops short of that just abandons the task half-done. */
  agenticStepCap: number
  /** Show the model's chain-of-thought in a collapsed block. Off means it is
   *  discarded entirely. Either way it never renders as the answer. */
  showReasoning: boolean
  /** Defined MCP servers. */
  mcpServers: McpServerConfig[]
  /** Per-server, per-tool opt-in. A tool absent from this map is DISABLED —
   *  the default for every newly discovered tool is off, so a server that
   *  gains tools on upgrade cannot silently gain reach. */
  mcpToolConsent: Record<string, Record<string, boolean>>
  /** Prompt before every MCP tool call, even for tools already enabled. */
  mcpConfirmEveryCall: boolean
  /** Agents/subagents the user has imported or created. */
  agents: AgentConfig[]
  /** Providers imported from opencode (e.g. LM-Studio, Ollama Cloud). */
  importedProviders: ImportedProviderConfig[]
  /** Global per-provider runtime options. */
  providerOptions: ProviderOptions
  /** Skill ids the user has explicitly disabled. */
  disabledSkills: string[]
  /** Whether to show imported opencode skills alongside built-ins. */
  showImportedSkills: boolean
  /** Origins of browser extensions the user has paired with this host, as
   *  exact `chrome-extension://<id>` strings.
   *
   *  An origin listed here may read the auth token from `/pair` and call the
   *  API cross-origin. That is the entire trust decision, which is why it is a
   *  deliberate list the user adds to rather than anything inferred. */
  pairedExtensions: string[]
}

export const DEFAULT_SETTINGS: Settings = {
  defaultModel: undefined,
  // The browser cannot be edited the way a document can, so its mode is fixed
  // at 'propose' in the UI; the entry exists so the record is total.
  defaultMode: { word: 'propose', excel: 'propose', browser: 'propose' },
  defaultContext: { word: 'selection', excel: 'range', browser: 'page' },
  theme: undefined,
  contextTrimWarningTokens: 6000,
  agenticStepCap: 30,
  showReasoning: true,
  mcpServers: [],
  mcpToolConsent: {},
  mcpConfirmEveryCall: false,
  agents: [],
  importedProviders: [],
  providerOptions: {},
  disabledSkills: [],
  showImportedSkills: true,
  pairedExtensions: [],
}

/** The step cap shipped as 10, which cut agentic Word work off partway through
 *  routine multi-edit tasks. A config written before the raise still holds that
 *  exact value, so treat it as "never chosen" and adopt the new default. Any
 *  other number is a cap the user actually picked, and is left alone. */
const SUPERSEDED_STEP_CAP = 10

export function migrateStepCap(stored: unknown): number {
  if (typeof stored !== 'number' || !Number.isFinite(stored)) {
    return DEFAULT_SETTINGS.agenticStepCap
  }
  return stored === SUPERSEDED_STEP_CAP ? DEFAULT_SETTINGS.agenticStepCap : stored
}

// ─── Health / API envelope ───────────────────────────────────────────────

export interface HealthResponse {
  status: 'ok' | 'degraded' | 'down'
  version: string
  port: number
  uptimeSeconds: number
  /** True when the host has seen a newer version on the update feed. */
  updateAvailable?: boolean
  /** Latest version available on the update feed, if a check has run. */
  latestVersion?: string
}

/** Server echoes the request id it assigned for a chat, so the pane can
 *  cancel. */
export interface ChatAcceptedResponse {
  requestId: string
}

export interface ApiError {
  code: string
  message: string
  retryable?: boolean
}

// ─── Update check / apply ─────────────────────────────────────────────────

export interface UpdateCheckResponse {
  updateAvailable: boolean
  currentVersion: string
  latestVersion: string
  /** ISO timestamp of the release, for the "released" line in the UI. */
  publishedAt?: string
  /** Short release notes body (already markdown-trimmed by the host). */
  releaseNotes?: string
}

export interface UpdateApplyResponse {
  ok: boolean
  message?: string
}
