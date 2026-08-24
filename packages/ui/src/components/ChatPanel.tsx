import { useEffect, useMemo, useRef, useState } from 'react'
import { useChatStore } from '../store/chatStore'
import { useSettingsStore } from '../store/settingsStore'
import { MessageBubble } from './MessageBubble'
import { ToolActivity, buildTranscript } from './ToolActivity'
import { ArrowDownIcon, PlusIcon } from './icons'
import { EmptyState } from './EmptyState'

const EMPTY_MESSAGES: readonly never[] = []

export function ChatPanel() {
  const conv = useChatStore((s) => s.conversations.find((c) => c.id === s.activeId))
  const streaming = useChatStore((s) => s.streaming)
  const stepLimit = useChatStore((s) => s.stepLimit)
  const newChat = useChatStore((s) => s.newChat)
  const showReasoning = useSettingsStore((s) => s.settings.showReasoning)
  const [editingId, setEditingId] = useState<string | null>(null)
  const scrollRef = useRef<HTMLDivElement | null>(null)
  const [atBottom, setAtBottom] = useState(true)
  const [showJump, setShowJump] = useState(false)
  const scrollPending = useRef(false)

  const messages = useMemo(() => conv?.messages ?? EMPTY_MESSAGES, [conv])
  const hasMessages = messages.length > 0

  // Streaming assistant message — used to mark it with role="status" for
  // screen readers, separate from the scrollable log (which must NOT be
  // aria-live or every token would be announced).
  const liveAssistant = useMemo(() => {
    if (!streaming) return null
    const last = messages[messages.length - 1]
    return last && last.role === 'assistant' ? last : null
  }, [streaming, messages])

  // Tool calls and their results are folded out of the transcript into
  // collapsed activity rows; see ToolActivity for why.
  const rows = useMemo(
    () => buildTranscript(messages, { showReasoning, streamingId: liveAssistant?.id }),
    [messages, showReasoning, liveAssistant],
  )

  // Attach the scroll listener via a callback ref so it re-binds whenever the
  // scroll container mounts (e.g. when transitioning from EmptyState to the
  // populated view). The old [] deps approach attached once to a null element.
  const setScrollRef = (el: HTMLDivElement | null) => {
    scrollRef.current = el
    if (!el) return
    const onScroll = () => {
      const nearBottom = el.scrollHeight - el.scrollTop - el.clientHeight < 80
      setAtBottom(nearBottom)
      setShowJump(!nearBottom)
    }
    el.addEventListener('scroll', onScroll, { passive: true })
    // Store cleanup on the element via a dataset marker so the ref callback
    // can remove the old listener if the element re-mounts.
    const prev = (el as HTMLDivElement & { __cleanup?: () => void }).__cleanup
    if (prev) prev()
    ;(el as HTMLDivElement & { __cleanup?: () => void }).__cleanup = () => {
      el.removeEventListener('scroll', onScroll)
    }
  }

  // Auto-scroll on new content, but throttle to one scroll per animation frame
  // so fast streaming (100s of tokens/sec) doesn't jank the layout.
  useEffect(() => {
    if (!atBottom) return
    if (scrollPending.current) return
    scrollPending.current = true
    requestAnimationFrame(() => {
      scrollPending.current = false
      const el = scrollRef.current
      if (el) el.scrollTop = el.scrollHeight
    })
  }, [messages, atBottom])

  const jumpToLatest = () => {
    const el = scrollRef.current
    if (el) {
      el.scrollTop = el.scrollHeight
      setAtBottom(true)
      setShowJump(false)
    }
  }

  if (!hasMessages) {
    return (
      <div className="flex h-full flex-col">
        <div className="flex-1 overflow-hidden">
          <EmptyState />
        </div>
      </div>
    )
  }

  return (
    <div className="relative flex h-full flex-col">
      <div className="flex items-center justify-between border-b border-surface-border px-3 py-1.5">
        <span className="truncate text-xs font-medium text-muted">{conv?.title ?? 'Chat'}</span>
        <button
          className="btn btn-ghost h-7 px-2 text-xs"
          onClick={newChat}
          aria-label="New chat"
          title="New chat"
        >
          <PlusIcon size={14} /> New
        </button>
      </div>
      <div
        ref={setScrollRef}
        className="flex-1 overflow-y-auto px-2"
        role="log"
        aria-busy={streaming}
      >
        {rows.map((row, i) =>
          row.kind === 'message' ? (
            <MessageBubble
              key={row.key}
              message={row.message}
              streaming={streaming && liveAssistant?.id === row.message.id}
              onEdit={(id) => setEditingId(id)}
            />
          ) : (
            <ToolActivity
              key={row.key}
              group={row.group}
              active={streaming && i === rows.length - 1}
            />
          ),
        )}
        {stepLimit !== null && (
          <div
            className="my-2 rounded-lg border border-surface-border bg-surface-muted px-2 py-1.5
            text-[0.72rem] text-muted"
          >
            Paused after {stepLimit} tool steps. Everything done so far stands — send another
            message to keep going.
          </div>
        )}
        {/* Dedicated live region for the streaming message — throttled by
            React's render batching and only announces the streaming message,
            not every historical message on mount. */}
        {liveAssistant && (
          <div role="status" aria-live="polite" className="sr-only">
            {streaming ? 'Generating response…' : ''}
          </div>
        )}
      </div>
      {showJump && (
        <button
          className="btn btn-subtle absolute bottom-3 left-1/2 -translate-x-1/2 px-3 py-1 text-xs shadow-pop"
          onClick={jumpToLatest}
          aria-label="Jump to latest message"
        >
          <ArrowDownIcon size={14} /> Latest
        </button>
      )}
      {editingId && <EditBanner id={editingId} onClose={() => setEditingId(null)} />}
    </div>
  )
}

function EditBanner({ id, onClose }: { id: string; onClose: () => void }) {
  const editAndResend = useChatStore((s) => s.editAndResend)
  const conv = useChatStore((s) => s.conversations.find((c) => c.id === s.activeId))
  const msg = conv?.messages.find((m) => m.id === id)
  const [text, setText] = useState(msg?.content ?? '')
  if (!msg) return null
  const canResend = text.trim().length > 0
  return (
    <div className="absolute inset-x-0 bottom-0 border-t border-surface-border bg-surface p-2 shadow-pane">
      <div className="mb-1 text-xs text-muted">
        Editing your message — resubmit sends a new turn.
      </div>
      <textarea
        className="field min-h-[60px] resize-y"
        value={text}
        onChange={(e) => setText(e.target.value)}
        autoFocus
      />
      <div className="mt-2 flex justify-end gap-2">
        <button className="btn btn-ghost px-3 py-1 text-xs" onClick={onClose}>
          Cancel
        </button>
        <button
          className="btn btn-primary px-3 py-1 text-xs"
          disabled={!canResend}
          onClick={() => {
            if (!canResend) return
            void editAndResend(id, text)
            onClose()
          }}
        >
          Resend
        </button>
      </div>
    </div>
  )
}
