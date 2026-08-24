import { useEffect, useRef, useState } from 'react'
import { useChatStore } from '../store/chatStore'
import { SendIcon, StopIcon } from './icons'
import { ModelSelector } from './ModelSelector'
import { ModeToggle } from './ModeToggle'
import { AgentSelector } from './AgentSelector'
import { estimateTokens } from '../util/tokens'

// The pane is ~320px wide. Anything placed *beside* the textarea comes straight
// out of the typing area, and with a model button, a mode button and a send
// button alongside it the input collapsed to roughly ten characters wide. So
// the controls live under the text instead, inside the same bordered surface —
// the textarea gets the full width, and the row beneath reads as one control
// strip rather than three widgets competing with the input.
const MIN_HEIGHT = 60
const MAX_HEIGHT = 220

export function Composer() {
  const [text, setText] = useState('')
  const streaming = useChatStore((s) => s.streaming)
  const send = useChatStore((s) => s.send)
  const cancel = useChatStore((s) => s.cancel)
  const reconnecting = useChatStore((s) => s.reconnecting)
  const [activeAgentId, setActiveAgentId] = useState<string | null>(null)
  const textareaRef = useRef<HTMLTextAreaElement>(null)

  useEffect(() => {
    const el = textareaRef.current
    if (!el) return
    el.style.height = 'auto'
    el.style.height = `${Math.min(Math.max(el.scrollHeight, MIN_HEIGHT), MAX_HEIGHT)}px`
    // Once the box stops growing the content has to scroll, or the tail of a
    // long message becomes untypeable.
    el.style.overflowY = el.scrollHeight > MAX_HEIGHT ? 'auto' : 'hidden'
  }, [text])

  const onSubmit = () => {
    if (streaming) return
    const t = text.trim()
    if (!t) return
    setText('')
    void send(t, activeAgentId ? { agentId: activeAgentId } : undefined)
  }

  const onKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Enter' && !e.shiftKey && !e.nativeEvent.isComposing) {
      e.preventDefault()
      onSubmit()
    }
  }

  const tokenEstimate = estimateTokens(text)

  return (
    <div className="border-t border-surface-border p-2">
      {reconnecting && (
        <div className="mb-1 px-1 text-[0.7rem] text-warn" role="status">
          Reconnecting (attempt {reconnecting.attempt})…
        </div>
      )}

      <div
        className="rounded-xl2 border border-surface-border bg-surface-muted transition-colors
          focus-within:border-accent"
      >
        <textarea
          ref={textareaRef}
          className="block w-full resize-none border-0 bg-transparent px-3 pb-1 pt-2.5 text-sm
            leading-snug text-fg placeholder:text-faint focus:outline-none focus:ring-0"
          style={{ height: MIN_HEIGHT }}
          placeholder={streaming ? 'Generating…' : 'Ask anything…'}
          title="Enter to send · Shift+Enter for a new line"
          value={text}
          onChange={(e) => setText(e.target.value)}
          onKeyDown={onKeyDown}
          aria-label="Message composer"
          rows={1}
        />

        <div className="flex items-center gap-1 px-1.5 pb-1.5">
          <AgentSelector activeId={activeAgentId} onSelect={setActiveAgentId} />
          <ModelSelector compact />
          <span className="h-4 w-px shrink-0 bg-surface-border" aria-hidden />
          <ModeToggle />
          <div className="min-w-0 flex-1" />
          {text.length > 0 && (
            <span className="shrink-0 whitespace-nowrap px-1 text-[0.65rem] text-faint" aria-hidden>
              ≈ {tokenEstimate} tokens
            </span>
          )}
          {streaming ? (
            <button
              className="btn btn-danger h-8 w-8 shrink-0"
              onClick={() => void cancel()}
              aria-label="Stop generating"
              title="Stop"
            >
              <StopIcon size={15} />
            </button>
          ) : (
            <button
              className="btn btn-primary h-8 w-8 shrink-0"
              onClick={onSubmit}
              disabled={!text.trim()}
              aria-label="Send message"
              title="Send (Enter)"
            >
              <SendIcon size={15} />
            </button>
          )}
        </div>
      </div>
    </div>
  )
}
