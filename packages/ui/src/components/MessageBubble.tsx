import { memo, useState } from 'react'
import type { ChatMessage } from '@openofficellm/shared'
import { Markdown } from './Markdown'
import { CopyIcon, RetryIcon, EditIcon } from './icons'
import { useChatStore } from '../store/chatStore'
import { useSettingsStore } from '../store/settingsStore'
import { textOf } from '../util/content'

// Tool calls and tool results never reach this component: ChatPanel folds them
// into a collapsed ToolActivity row before rendering. This draws prose only.
interface Props {
  message: ChatMessage
  streaming?: boolean
  onEdit?: (id: string) => void
}

function timeLabel(ts?: number): string {
  if (!ts) return ''
  const d = new Date(ts)
  return d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
}

export const MessageBubble = memo(function MessageBubble({ message, streaming, onEdit }: Props) {
  const [copied, setCopied] = useState(false)
  const retry = useChatStore((s) => s.retry)
  const showReasoning = useSettingsStore((s) => s.settings.showReasoning)
  const isUser = message.role === 'user'

  const onCopy = async () => {
    try {
      await navigator.clipboard.writeText(textOf(message.content))
      setCopied(true)
      setTimeout(() => setCopied(false), 1500)
    } catch {
      // clipboard may be unavailable
    }
  }

  return (
    <div
      className={`group my-2 flex flex-col ${isUser ? 'items-end' : 'items-start'}`}
      role="article"
      aria-label={`${message.role} message`}
    >
      <div
        className={`max-w-[92%] rounded-xl2 px-3 py-2 text-sm ${
          isUser ? 'bg-accent text-white' : 'bg-surface-muted text-fg'
        }`}
      >
        {!isUser && message.reasoning && showReasoning && (
          <ReasoningBlock text={message.reasoning} streaming={streaming} />
        )}
        {isUser ? (
          <div className="whitespace-pre-wrap break-words">{textOf(message.content)}</div>
        ) : (
          <Markdown
            content={textOf(message.content) || (streaming ? '…' : '')}
            streaming={streaming}
            className="prose-sm max-w-none break-words"
          />
        )}
      </div>
      <div className="mt-1 flex items-center gap-2 px-1 text-[0.7rem] text-faint opacity-0 transition-opacity group-hover:opacity-100 focus-within:opacity-100">
        <span>{timeLabel(message.createdAt)}</span>
        {message.model && <span className="text-faint">· {message.model}</span>}
        <button
          className="icon-btn h-5 w-5 hover:text-fg"
          onClick={onCopy}
          aria-label="Copy message"
          title="Copy"
        >
          <CopyIcon size={13} />
        </button>
        {!isUser && !streaming && textOf(message.content) && (
          <button
            className="icon-btn h-5 w-5 hover:text-fg"
            onClick={() => void retry(message.id!)}
            aria-label="Retry"
            title="Retry"
          >
            <RetryIcon size={13} />
          </button>
        )}
        {isUser && onEdit && !streaming && (
          <button
            className="icon-btn h-5 w-5 hover:text-fg"
            onClick={() => onEdit(message.id!)}
            aria-label="Edit and resend"
            title="Edit and resend"
          >
            <EditIcon size={13} />
          </button>
        )}
        {copied && <span className="text-ok">copied</span>}
      </div>
    </div>
  )
})

/**
 * The model's chain-of-thought, collapsed.
 *
 * Reasoning models emit their scratchpad before the answer. It used to stream
 * into the message body, so the pane showed the model planning its response and
 * then giving it — the user's actual answer buried under a numbered list of
 * what the model was about to say.
 */
function ReasoningBlock({ text, streaming }: { text: string; streaming?: boolean }) {
  return (
    <details className="mb-2 rounded-lg border border-surface-border bg-surface/60 text-xs">
      <summary className="cursor-pointer select-none px-2 py-1 text-faint marker:text-faint">
        {streaming ? 'Thinking…' : 'Thought process'}
      </summary>
      <div className="max-h-56 overflow-y-auto whitespace-pre-wrap break-words border-t border-surface-border px-2 py-1.5 text-[0.72rem] leading-relaxed text-muted">
        {text}
      </div>
    </details>
  )
}
