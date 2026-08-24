// One collapsed block for everything the model did between two answers.
//
// A formatting request can easily run a dozen rounds — read the document,
// restyle four headings, bullet a list, re-read to check. Rendered one bubble
// per call plus one per result, that pushed the actual answer off the top of a
// 320px-wide pane. So the whole run collapses into a single "Working" row that
// the user can open when they want to audit what happened, and ignore when they
// do not.

import { useState } from 'react'
import type { ChatMessage } from '@openofficellm/shared'
import { ChevronRightIcon } from './icons'

export interface ToolStep {
  /** The tool call id, or a synthetic one for an orphan result. */
  id: string
  name: string
  /** Raw JSON arguments as the model emitted them. */
  arguments: string
  /** Tool output, absent while the call is still running. */
  result?: string
}

export interface ToolActivityGroup {
  id: string
  steps: ToolStep[]
}

export type TranscriptRow =
  | { kind: 'message'; key: string; message: ChatMessage }
  | { kind: 'activity'; key: string; group: ToolActivityGroup }

/**
 * Fold a message list into renderable rows, pulling tool traffic out of the
 * transcript and into activity groups.
 *
 * Assistant messages keep their prose; only their `toolCalls` move. An
 * assistant message with nothing but tool calls disappears entirely into the
 * group, which is the common case — most models emit no commentary alongside a
 * call.
 */
export function buildTranscript(
  messages: readonly ChatMessage[],
  opts: { showReasoning: boolean; streamingId?: string },
): TranscriptRow[] {
  const rows: TranscriptRow[] = []
  // The group currently being filled. Consecutive tool traffic accumulates
  // here; any message with a visible body closes it, so the transcript reads
  // "question → working → answer" instead of interleaving the two.
  let open: ToolActivityGroup | null = null

  for (let i = 0; i < messages.length; i++) {
    const msg = messages[i]
    const key = msg.id ?? `msg_${i}`

    if (msg.role === 'tool') {
      if (open === null) {
        open = { id: `act_${key}`, steps: [] }
        rows.push({ kind: 'activity', key: open.id, group: open })
      }
      const step = open.steps.find((s) => s.id === msg.toolCallId && s.result === undefined)
      if (step) {
        step.result = msg.content
      } else {
        // A result with no matching call still has to be shown — dropping it
        // would hide a tool the model ran.
        open.steps.push({
          id: key,
          name: msg.toolName ?? 'tool',
          arguments: '',
          result: msg.content,
        })
      }
      continue
    }

    if (msg.role !== 'assistant' || hasVisibleBody(msg, opts)) {
      open = null
      rows.push({ kind: 'message', key, message: msg })
    }

    const calls = msg.toolCalls
    if (msg.role === 'assistant' && calls && calls.length > 0) {
      if (open === null) {
        open = { id: `act_${key}`, steps: [] }
        rows.push({ kind: 'activity', key: open.id, group: open })
      }
      for (const call of calls) {
        open.steps.push({ id: call.id, name: call.name, arguments: call.arguments })
      }
    }
  }

  return rows
}

/** Whether an assistant message has anything left to render once its tool calls
 *  have been moved into an activity group. */
function hasVisibleBody(
  msg: ChatMessage,
  opts: { showReasoning: boolean; streamingId?: string },
): boolean {
  if (msg.content.trim().length > 0) return true
  if (opts.showReasoning && (msg.reasoning ?? '').trim().length > 0) return true
  // The streaming placeholder: an empty message that has not yet decided
  // whether it is going to answer or call a tool still needs to show that
  // something is happening. Once tool calls arrive the activity row says so,
  // and a bubble containing only "…" is noise.
  return msg.id === opts.streamingId && !(msg.toolCalls && msg.toolCalls.length > 0)
}

const MAX_RESULT_CHARS = 8000

export function ToolActivity({ group, active }: { group: ToolActivityGroup; active?: boolean }) {
  const [open, setOpen] = useState(false)
  const steps = group.steps
  const running = active ? steps.find((s) => s.result === undefined) : undefined

  return (
    <div className="my-1.5 overflow-hidden rounded-lg border border-surface-border bg-surface-muted">
      <button
        type="button"
        className="flex w-full items-center gap-1.5 px-2 py-1.5 text-left text-xs
          text-muted hover:bg-surface-hover focus-visible:outline-none
          focus-visible:ring-2 focus-visible:ring-accent"
        onClick={() => setOpen((v) => !v)}
        aria-expanded={open}
      >
        <span className={`shrink-0 transition-transform ${open ? 'rotate-90' : ''}`}>
          <ChevronRightIcon size={12} />
        </span>
        <span className={`shrink-0 font-medium ${active ? 'text-accent-fg' : 'text-muted'}`}>
          {active ? 'Working…' : summarizeCount(steps.length)}
        </span>
        <span className="truncate font-mono text-[0.7rem] text-faint">
          {running ? running.name : toolNames(steps)}
        </span>
      </button>

      {open && (
        <ul className="border-t border-surface-border">
          {steps.map((step, i) => (
            <ToolStepRow key={`${step.id}_${i}`} step={step} />
          ))}
        </ul>
      )}
    </div>
  )
}

function ToolStepRow({ step }: { step: ToolStep }) {
  const [open, setOpen] = useState(false)
  const hasResult = step.result !== undefined
  const args = summarizeArgs(step.arguments)

  return (
    <li className="border-b border-surface-border/60 last:border-b-0">
      <button
        type="button"
        className="flex w-full items-start gap-1.5 px-2 py-1 text-left hover:bg-surface-hover
          focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent"
        onClick={() => setOpen((v) => !v)}
        aria-expanded={open}
        disabled={!hasResult && !args}
      >
        <span className="shrink-0 pt-0.5 font-mono text-[0.7rem] text-accent-fg">{step.name}</span>
        <span className="min-w-0 flex-1 truncate font-mono text-[0.7rem] text-faint">{args}</span>
        {!hasResult && <span className="shrink-0 text-[0.65rem] text-faint">running…</span>}
      </button>
      {open && (
        <div className="px-2 pb-1.5">
          {args && (
            <pre
              className="max-h-32 overflow-auto whitespace-pre-wrap break-words rounded
              bg-surface/60 px-1.5 py-1 font-mono text-[0.68rem] text-muted"
            >
              {step.arguments || args}
            </pre>
          )}
          {hasResult && (
            <pre
              className="mt-1 max-h-48 overflow-auto whitespace-pre-wrap break-words rounded
              bg-surface/60 px-1.5 py-1 font-mono text-[0.68rem] text-fg"
            >
              {truncate(step.result ?? '')}
            </pre>
          )}
        </div>
      )}
    </li>
  )
}

function truncate(text: string): string {
  return text.length > MAX_RESULT_CHARS ? `${text.slice(0, MAX_RESULT_CHARS)}\n…[truncated]` : text
}

function summarizeCount(n: number): string {
  return n === 1 ? '1 step' : `${n} steps`
}

/** Distinct tool names in call order — repeats say nothing extra in a label
 *  that has about forty characters to work with. */
function toolNames(steps: readonly ToolStep[]): string {
  const seen: string[] = []
  for (const s of steps) if (!seen.includes(s.name)) seen.push(s.name)
  return seen.join(', ')
}

/** One-line rendering of tool arguments. The raw JSON of a replace_selection
 *  call is the entire replacement text, which is not a useful label. */
export function summarizeArgs(json: string): string {
  if (!json) return ''
  try {
    const args = JSON.parse(json) as Record<string, unknown>
    return Object.entries(args)
      .filter(([, v]) => v !== null && v !== undefined && v !== '')
      .map(([k, v]) => {
        const s = typeof v === 'string' ? v : JSON.stringify(v)
        return `${k}=${s.length > 40 ? `${s.slice(0, 40)}…` : s}`
      })
      .join(' ')
  } catch {
    return json.length > 60 ? `${json.slice(0, 60)}…` : json
  }
}
