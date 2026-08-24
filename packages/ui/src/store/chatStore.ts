// Chat store: messages, streaming state, the tool-calling agent loop,
// send/cancel/retry/edit-and-resend, plus conversation history persisted to
// localStorage (host-side history is a later phase).
//
// The agent loop lives here rather than in the host service because Office.js
// only exists inside the pane's webview. The host streams a tool_call back, the
// pane executes it against the live document, appends the result, and re-sends —
// so the loop necessarily runs on this side of the wire.

import { create } from 'zustand'
import type {
  ChatMessage,
  DocumentContext,
  EditMode,
  StreamEvent,
  ToolCall,
  ToolDefinition,
} from '@openofficellm/shared'
import { cancelChat, callMcpTool } from '../api/client'
import { openChatStream, type StreamHandle } from '../api/stream'
import { loadPersisted, savePersisted, isArray } from './persist'
import { useModelsStore } from './modelsStore'
import { useSettingsStore } from './settingsStore'
import { useContextStore } from './contextStore'
import { useProposalStore } from './proposalStore'
import { getHost, getAdapter, getDocumentKey, settingsHost, shell } from '../host/bridge'

export interface Conversation {
  id: string
  title: string
  messages: ChatMessage[]
  createdAt: number
  updatedAt: number
  /** Which document this chat belongs to (see `getDocumentKey`). Absent on
   *  conversations saved before chats were scoped per document; those are never
   *  auto-restored, only reachable from the sidebar. */
  docKey?: string
}

export interface ChatState {
  conversations: Conversation[]
  activeId: string | null
  streaming: boolean
  streamingRequestId: string | null
  /** Name of the tool currently executing, for the status line. */
  runningTool: string | null
  /** Accumulated session usage/cost. */
  usage: {
    promptTokens: number
    completionTokens: number
    totalTokens: number
    estimatedCostUsd: number
  }
  lastError: { code: string; message: string; retryable?: boolean } | null
  reconnecting: { attempt: number; backoffMs: number } | null
  /** Set when a turn ran out of tool steps. A quiet transcript note rather than
   *  an error: nothing failed, the model simply had more to do. */
  stepLimit: number | null

  /** Point the pane at the document it is running in, restoring that
   *  document's most recent chat and leaving other documents' chats alone. */
  adoptDocument: (docKey: string) => void
  newChat: () => void
  selectChat: (id: string) => void
  renameChat: (id: string, title: string) => void
  deleteChat: (id: string) => void

  send: (
    text: string,
    opts?: { mode?: EditMode; skillId?: string; agentId?: string },
  ) => Promise<void>
  cancel: () => Promise<void>
  retry: (messageId: string) => Promise<void>
  editAndResend: (messageId: string, newText: string) => Promise<void>
  clearError: () => void

  activeConversation: () => Conversation | null
}

const SESSION_KEY = 'chat.conversations'
const ACTIVE_KEY = 'chat.activeId'

function genId(prefix = 'id'): string {
  return `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`
}

function nowMs(): number {
  return Date.now()
}

function deriveTitle(text: string): string {
  const trimmed = text.trim().replace(/\s+/g, ' ')
  if (trimmed.length <= 40) return trimmed || 'New chat'
  return `${trimmed.slice(0, 40)}…`
}

function ensureConversation(state: ChatState): Conversation {
  const active = state.activeId ? state.conversations.find((c) => c.id === state.activeId) : null
  if (active) return active
  const conv: Conversation = {
    id: genId('conv'),
    title: 'New chat',
    messages: [],
    createdAt: nowMs(),
    updatedAt: nowMs(),
    docKey: getDocumentKey(),
  }
  state.conversations = [conv, ...state.conversations]
  state.activeId = conv.id
  persistConversations(state.conversations, conv.id)
  return conv
}

function persistConversations(conversations: Conversation[], activeId: string | null) {
  savePersisted(SESSION_KEY, conversations)
  savePersisted(ACTIVE_KEY, activeId)
}

let activeStream: StreamHandle | null = null
// Monotonic generation counter. `runTurn` captures the current value and bails
// out of its event loop if the value changed (a newer send/retry/edit started).
// This prevents a stale stream's late events from corrupting a newer
// conversation — e.g. user switches chats or sends again before the prior
// stream's `finally` runs.
let streamGeneration = 0

/** Resolve the edit mode for a turn. */
function resolveMode(override?: EditMode): EditMode {
  if (override) return override
  const settings = useSettingsStore.getState().settings
  return settings.defaultMode[settingsHost()]
}

/**
 * Gather the document context to attach to a turn.
 *
 * Explicit context chips win — if the user picked what to include, that is the
 * answer. Otherwise we read the host's default scope automatically. Auto-attach
 * is what makes the first message of a conversation work: without it the model
 * has to spend a round trip on a read tool before it can answer "summarize
 * this", and a model that decides not to bother answers about nothing at all.
 */
async function gatherContext(): Promise<DocumentContext | undefined> {
  const explicit = useContextStore.getState().toDocumentContext()
  if (explicit) return explicit

  const adapter = getAdapter()
  if (!adapter) return undefined
  const settings = useSettingsStore.getState().settings
  const scope = settings.defaultContext[adapter.host]
  if (scope === 'none') return undefined
  try {
    return await adapter.getContext(scope)
  } catch {
    // A document that cannot be read is not a reason to refuse to chat.
    return undefined
  }
}

export const useChatStore = create<ChatState>((set, get) => ({
  conversations: loadPersisted<Conversation[]>(
    SESSION_KEY,
    [],
    (v): v is Conversation[] =>
      isArray(v) &&
      v.every((c) => c && typeof c === 'object' && isArray((c as Conversation).messages)),
  ),
  activeId: loadPersisted<string | null>(
    ACTIVE_KEY,
    null,
    (v): v is string | null => v === null || typeof v === 'string',
  ),
  streaming: false,
  streamingRequestId: null,
  runningTool: null,
  usage: { promptTokens: 0, completionTokens: 0, totalTokens: 0, estimatedCostUsd: 0 },
  lastError: null,
  reconnecting: null,
  stepLimit: null,

  adoptDocument: (docKey) => {
    set((s) => {
      const active = s.conversations.find((c) => c.id === s.activeId)
      if (active?.docKey === docKey) return {}
      const mine = s.conversations
        .filter((c) => c.docKey === docKey)
        .sort((a, b) => b.updatedAt - a.updatedAt)
      const activeId = mine[0]?.id ?? null
      savePersisted(ACTIVE_KEY, activeId)
      return { activeId, lastError: null, stepLimit: null }
    })
  },

  newChat: () => {
    const conv: Conversation = {
      id: genId('conv'),
      title: 'New chat',
      messages: [],
      createdAt: nowMs(),
      updatedAt: nowMs(),
      docKey: getDocumentKey(),
    }
    set((s) => {
      const conversations = [conv, ...s.conversations]
      persistConversations(conversations, conv.id)
      return { conversations, activeId: conv.id, lastError: null, stepLimit: null }
    })
    useProposalStore.getState().discardAll()
  },

  selectChat: (id) => {
    savePersisted(ACTIVE_KEY, id)
    set({ activeId: id, lastError: null, stepLimit: null })
    useProposalStore.getState().discardAll()
  },

  renameChat: (id, title) =>
    set((s) => {
      const conversations = s.conversations.map((c) =>
        c.id === id ? { ...c, title: title.trim() || c.title, updatedAt: nowMs() } : c,
      )
      persistConversations(conversations, s.activeId)
      return { conversations }
    }),

  deleteChat: (id) =>
    set((s) => {
      const conversations = s.conversations.filter((c) => c.id !== id)
      const activeId = s.activeId === id ? (conversations[0]?.id ?? null) : s.activeId
      persistConversations(conversations, activeId)
      return { conversations, activeId }
    }),

  send: async (text, opts) => {
    const state = get()
    if (state.streaming) return
    const trimmed = text.trim()
    if (!trimmed) return

    const models = useModelsStore.getState()
    const modelId = models.selectedModelId
    if (!modelId) {
      set({ lastError: { code: 'no_model', message: 'Select a model first.' } })
      return
    }

    const mode = resolveMode(opts?.mode)
    const context = await gatherContext()

    // A new turn supersedes anything staged by the previous one. Leaving stale
    // proposals visible invites applying an edit the user has since talked the
    // model out of.
    useProposalStore.getState().discardAll()

    const conv = ensureConversation(get())
    const userMsg: ChatMessage = {
      id: genId('msg'),
      role: 'user',
      content: trimmed,
      createdAt: nowMs(),
    }
    const baseMessages = [...conv.messages, userMsg]
    const updatedConv: Conversation = {
      ...conv,
      title: conv.messages.length === 0 ? deriveTitle(trimmed) : conv.title,
      messages: baseMessages,
      updatedAt: nowMs(),
    }
    set((s) => {
      const conversations = s.conversations.map((c) => (c.id === updatedConv.id ? updatedConv : c))
      persistConversations(conversations, updatedConv.id)
      return {
        conversations,
        streaming: true,
        streamingRequestId: null,
        lastError: null,
        stepLimit: null,
      }
    })

    await runAgentLoop(set, get, {
      conversationId: conv.id,
      history: baseMessages,
      model: modelId,
      mode,
      context,
      skillId: opts?.skillId,
      agentId: opts?.agentId,
    })
  },

  cancel: async () => {
    const state = get()
    if (!state.streaming) return
    const reqId = state.streamingRequestId
    // Bump the generation first so an in-flight loop stops issuing new rounds,
    // then abort the local fetch so the stream exits promptly, then notify the
    // host. If reqId is null (start event hasn't arrived yet), the host tears
    // down its orphan stream when it sees the client disconnect.
    streamGeneration++
    activeStream?.abort()
    if (reqId) {
      try {
        await cancelChat(reqId)
      } catch {
        // best-effort
      }
    }
    set({ streaming: false, streamingRequestId: null, reconnecting: null, runningTool: null })
  },

  retry: async (messageId) => {
    const state = get()
    if (state.streaming) return
    const conv = state.activeConversation()
    if (!conv) return
    const idx = conv.messages.findIndex((m) => m.id === messageId)
    if (idx < 0) return
    const target = conv.messages[idx]
    if (target.role !== 'assistant') return
    // Regenerate the assistant answer using the full prior context. Resending
    // just the last user text would discard the entire preceding conversation.
    const prior = conv.messages.slice(0, idx)
    if (prior.length === 0) return
    const modelId = useModelsStore.getState().selectedModelId
    if (!modelId) {
      set({ lastError: { code: 'no_model', message: 'Select a model first.' } })
      return
    }
    const mode = resolveMode()
    const context = await gatherContext()
    useProposalStore.getState().discardAll()

    const updatedConv: Conversation = { ...conv, messages: prior, updatedAt: nowMs() }
    set((s) => {
      const conversations = s.conversations.map((c) => (c.id === conv.id ? updatedConv : c))
      persistConversations(conversations, conv.id)
      return {
        conversations,
        streaming: true,
        streamingRequestId: null,
        lastError: null,
        stepLimit: null,
      }
    })
    await runAgentLoop(set, get, {
      conversationId: conv.id,
      history: prior,
      model: modelId,
      mode,
      context,
    })
  },

  editAndResend: async (messageId, newText) => {
    const state = get()
    if (state.streaming) return
    const trimmed = newText.trim()
    if (!trimmed) return
    const conv = state.activeConversation()
    if (!conv) return
    const idx = conv.messages.findIndex((m) => m.id === messageId)
    if (idx < 0) return
    const target = conv.messages[idx]
    if (target.role !== 'user') return
    const modelId = useModelsStore.getState().selectedModelId
    if (!modelId) {
      set({ lastError: { code: 'no_model', message: 'Select a model first.' } })
      return
    }
    const truncated = conv.messages.slice(0, idx)
    const edited: ChatMessage = { ...target, content: trimmed, createdAt: nowMs() }
    const history = [...truncated, edited]
    const mode = resolveMode()
    const context = await gatherContext()
    useProposalStore.getState().discardAll()

    const updatedConv: Conversation = { ...conv, messages: history, updatedAt: nowMs() }
    set((s) => {
      const conversations = s.conversations.map((c) => (c.id === conv.id ? updatedConv : c))
      persistConversations(conversations, conv.id)
      return {
        conversations,
        streaming: true,
        streamingRequestId: null,
        lastError: null,
        stepLimit: null,
      }
    })
    await runAgentLoop(set, get, {
      conversationId: conv.id,
      history,
      model: modelId,
      mode,
      context,
    })
  },

  clearError: () => set({ lastError: null }),

  activeConversation: () => {
    const s = get()
    return s.conversations.find((c) => c.id === s.activeId) ?? null
  },
}))

type SetState = (partial: Partial<ChatState> | ((s: ChatState) => Partial<ChatState>)) => void
type GetState = () => ChatState

interface TurnOptions {
  conversationId: string
  /** Conversation so far, ending with the user message being answered. */
  history: ChatMessage[]
  model: string
  mode: EditMode
  context?: DocumentContext
  skillId?: string
  agentId?: string
}

/** Append a message to a conversation and persist. */
function appendMessage(set: SetState, conversationId: string, message: ChatMessage): void {
  set((s) => {
    const conv = s.conversations.find((c) => c.id === conversationId)
    if (!conv) return {}
    const messages = [...conv.messages, message]
    const conversations = s.conversations.map((c) =>
      c.id === conv.id ? { ...conv, messages, updatedAt: nowMs() } : c,
    )
    persistConversations(conversations, conv.id)
    return { conversations }
  })
}

function updateMessage(
  set: SetState,
  conversationId: string,
  messageId: string,
  mutate: (m: ChatMessage) => ChatMessage,
): void {
  set((s) => {
    // Apply deltas to the conversation that was active when the stream started,
    // NOT whatever is active now — switching chats mid-stream must not retarget
    // the stream onto a different conversation.
    const conv = s.conversations.find((c) => c.id === conversationId)
    if (!conv) return {}
    const messages = conv.messages.map((m) => (m.id === messageId ? mutate(m) : m))
    const conversations = s.conversations.map((c) =>
      c.id === conv.id ? { ...conv, messages, updatedAt: nowMs() } : c,
    )
    persistConversations(conversations, conv.id)
    return { conversations }
  })
}

/**
 * Run a turn to completion, including any rounds of tool calling.
 *
 * Each round appends its own assistant message, so the transcript reads as
 * "thought → called a tool → result → answer" rather than collapsing everything
 * into one bubble with no explanation of where the numbers came from.
 */
async function runAgentLoop(set: SetState, get: GetState, opts: TurnOptions): Promise<void> {
  const myGeneration = ++streamGeneration
  const { conversationId, model, mode, context, skillId, agentId } = opts

  const settings = useSettingsStore.getState().settings
  const agent = agentId ? settings.agents.find((a) => a.id === agentId && a.enabled) : undefined
  const effectiveMode = agent?.mode ?? mode
  const effectiveModel = agent?.model ?? model
  // One cap for both modes. Propose mode still drives the same tool loop — a
  // formatting pass reads the document, makes several edits and re-reads to
  // check itself — so a separate, tighter budget there only ever abandoned the
  // work partway through.
  const stepCap = Math.max(1, settings.agenticStepCap)

  const host = getHost()
  // Write tools are pointless without a document, and withheld entirely when
  // the user has set the context scope to none.
  const scope = settings.defaultContext[settingsHost(host)]
  const tools: ToolDefinition[] = shell().toolCatalog(
    host,
    scope !== 'none' && agent?.tools !== false,
  )

  let working = [...opts.history]

  try {
    for (let round = 0; round < stepCap; round++) {
      if (streamGeneration !== myGeneration) return

      const assistantId = genId('asst')
      appendMessage(set, conversationId, {
        id: assistantId,
        role: 'assistant',
        content: '',
        model: effectiveModel,
        createdAt: nowMs(),
      })

      const result = await runOneRound(set, get, myGeneration, conversationId, assistantId, {
        conversationId,
        // Only the first round carries the document snapshot: re-sending it on
        // every round would multiply a large document by the round count, and
        // the model already has it in history.
        context: round === 0 ? context : undefined,
        messages: working,
        model: effectiveModel,
        mode: effectiveMode,
        skillId: round === 0 ? skillId : undefined,
        tools: tools.length > 0 ? tools : undefined,
      })

      if (result.aborted) return
      if (result.error) return

      const assistantMsg: ChatMessage = {
        role: 'assistant',
        content: result.content,
        ...(result.toolCalls.length > 0 ? { toolCalls: result.toolCalls } : {}),
      }
      working = [...working, assistantMsg]

      if (result.toolCalls.length === 0) return

      // Execute every tool the model asked for, then loop with the results.
      const toolMessages: ChatMessage[] = []
      for (const call of result.toolCalls) {
        if (streamGeneration !== myGeneration) return
        set({ runningTool: call.name })
        const outcome = await executeToolCall(call, effectiveMode)
        const toolMsg: ChatMessage = {
          id: genId('tool'),
          role: 'tool',
          content: outcome.content,
          toolCallId: call.id,
          toolName: call.name,
          createdAt: nowMs(),
        }
        appendMessage(set, conversationId, toolMsg)
        toolMessages.push({
          role: 'tool',
          content: outcome.content,
          toolCallId: call.id,
          toolName: call.name,
        })
      }
      set({ runningTool: null })
      working = [...working, ...toolMessages]
    }

    // Ran out of rounds with the model still wanting tools. Not an error — the
    // work so far stands — so it goes in the transcript, not the error banner.
    if (streamGeneration === myGeneration) set({ stepLimit: stepCap })
  } finally {
    if (streamGeneration === myGeneration) {
      set({ streaming: false, streamingRequestId: null, reconnecting: null, runningTool: null })
      activeStream = null
    }
  }
}

interface RoundResult {
  content: string
  toolCalls: ToolCall[]
  aborted: boolean
  error: boolean
}

async function runOneRound(
  set: SetState,
  get: GetState,
  myGeneration: number,
  conversationId: string,
  assistantId: string,
  payload: {
    conversationId?: string
    messages: ChatMessage[]
    model: string
    mode: EditMode
    context?: DocumentContext
    skillId?: string
    tools?: ToolDefinition[]
  },
): Promise<RoundResult> {
  const wire = {
    ...payload,
    messages: payload.messages.map((m) => ({
      role: m.role,
      content: m.content,
      ...(m.toolCalls ? { toolCalls: m.toolCalls } : {}),
      ...(m.toolCallId ? { toolCallId: m.toolCallId } : {}),
      ...(m.toolName ? { toolName: m.toolName } : {}),
    })),
  }

  const handle = openChatStream(
    wire,
    { maxReconnects: 2 },
    {
      onReconnect: (attempt, backoffMs) => {
        if (streamGeneration === myGeneration) set(() => ({ reconnecting: { attempt, backoffMs } }))
      },
      onReconnectFailed: () => {
        if (streamGeneration === myGeneration) set(() => ({ reconnecting: null }))
      },
    },
  )
  activeStream = handle

  let content = ''
  const toolCalls: ToolCall[] = []

  for await (const ev of handle.events) {
    if (streamGeneration !== myGeneration)
      return { content, toolCalls, aborted: true, error: false }
    const streamEv = ev as StreamEvent
    switch (streamEv.type) {
      case 'start':
        set({ streamingRequestId: streamEv.requestId, reconnecting: null })
        updateMessage(set, conversationId, assistantId, (m) => ({ ...m, model: streamEv.model }))
        break
      case 'delta':
        content += streamEv.text
        updateMessage(set, conversationId, assistantId, (m) => ({
          ...m,
          content: m.content + streamEv.text,
        }))
        break
      case 'reasoning':
        // Kept out of `content` entirely. It renders in its own collapsed block
        // and is stripped before history goes back upstream.
        updateMessage(set, conversationId, assistantId, (m) => ({
          ...m,
          reasoning: (m.reasoning ?? '') + streamEv.text,
        }))
        break
      case 'tool_call':
        toolCalls.push(streamEv.toolCall)
        updateMessage(set, conversationId, assistantId, (m) => ({
          ...m,
          toolCalls: [...(m.toolCalls ?? []), streamEv.toolCall],
        }))
        break
      case 'usage':
        set((s) => ({
          usage: {
            promptTokens: s.usage.promptTokens + streamEv.promptTokens,
            completionTokens: s.usage.completionTokens + streamEv.completionTokens,
            totalTokens: s.usage.totalTokens + streamEv.totalTokens,
            estimatedCostUsd: s.usage.estimatedCostUsd,
          },
        }))
        break
      case 'cost':
        set((s) => ({
          usage: {
            ...s.usage,
            estimatedCostUsd: s.usage.estimatedCostUsd + streamEv.estimatedCostUsd,
          },
        }))
        break
      case 'done':
        return { content, toolCalls, aborted: false, error: false }
      case 'error':
        if (streamGeneration === myGeneration) {
          set({
            streaming: false,
            streamingRequestId: null,
            reconnecting: null,
            runningTool: null,
            lastError: {
              code: streamEv.code,
              message: streamEv.message,
              retryable: streamEv.retryable,
            },
          })
          updateMessage(set, conversationId, assistantId, (m) =>
            m.content.length === 0 ? { ...m, content: `⚠️ ${streamEv.message}` } : m,
          )
          activeStream = null
        }
        return { content, toolCalls, aborted: false, error: true }
    }
  }
  return { content, toolCalls, aborted: false, error: false }
}

/** Route a tool call to the pane (document tools) or the host (MCP tools). */
async function executeToolCall(
  call: ToolCall,
  mode: EditMode,
): Promise<{ content: string; isError: boolean }> {
  if (call.name.startsWith('mcp__')) {
    // MCP tools run in the host service, which enforces consent. MCP output is
    // untrusted data: it comes back as a tool result for the model to read, and
    // can only reach the document by the model calling a document tool, which
    // still passes through the active edit mode.
    try {
      const result = await callMcpTool(call.name, safeParse(call.arguments))
      return { content: result.content, isError: result.isError }
    } catch (e) {
      return { content: `MCP call failed: ${(e as Error).message}`, isError: true }
    }
  }

  const adapter = getAdapter()
  if (!adapter) {
    return {
      content:
        'No document is open. This add-in is running outside Word or Excel, so document tools are unavailable.',
      isError: true,
    }
  }
  return shell().executeDocumentTool(call.name, call.arguments, {
    adapter,
    mode,
    propose: (edit, description) => {
      useProposalStore.getState().add({ edit, description, turnId: String(streamGeneration) })
    },
  })
}

function safeParse(json: string): Record<string, unknown> {
  try {
    const v = JSON.parse(json || '{}')
    return v && typeof v === 'object' ? (v as Record<string, unknown>) : {}
  } catch {
    return {}
  }
}
