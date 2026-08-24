// Splits inline chain-of-thought out of a model's visible output.
//
// Two different things both surface as "thinking", and both leaked into the
// answer before this existed:
//
//   1. A dedicated field on the wire — Ollama's `message.thinking`, OpenAI's
//      `reasoning_content`, Anthropic's `thinking` blocks. Adapters map those
//      straight onto `reasoning` events.
//   2. Tags inlined in the content stream — `<think>…</think>` and friends,
//      emitted by GLM, DeepSeek-R1, Qwen3, and most local reasoning finetunes.
//      Nothing on the wire distinguishes it from the answer.
//
// Case 2 is what this file handles. It ran invisibly before: the pane's
// markdown renderer sanitizes unknown HTML, so `<think>` vanished and its
// contents rendered as prose — the model appeared to answer with its own
// scratchpad.
//
// Applied centrally in server.ts rather than per-adapter, so a new provider
// cannot forget it.

import type { StreamEvent } from '@openofficellm/shared'

/** Recognised reasoning tag names. Matched case-insensitively. */
const TAG_NAMES = ['think', 'thinking', 'reason', 'reasoning', 'thought'] as const

const TAG_RE = new RegExp(`<(/?)(?:${TAG_NAMES.join('|')})\\s*>`, 'i')

/** Every literal that could begin at a `<`, for the hold-back check. */
const TAG_LITERALS: string[] = TAG_NAMES.flatMap((n) => [`<${n}>`, `</${n}>`])
const MAX_TAG_LEN = Math.max(...TAG_LITERALS.map((t) => t.length))

/**
 * How much of the tail must be withheld because it might be the start of a tag
 * split across two deltas. Returns an index into `buf`: everything before it is
 * safe to emit.
 *
 * Only withholds a tail that is genuinely a prefix of a known tag, so ordinary
 * prose containing `<` (or a code fence full of HTML) streams without a stall.
 */
function safeEmitLength(buf: string): number {
  const lt = buf.lastIndexOf('<')
  if (lt === -1) return buf.length
  const tail = buf.slice(lt)
  // A complete tag would already have been matched by TAG_RE, and a `>` means
  // this `<…>` is something else entirely.
  if (tail.includes('>')) return buf.length
  if (tail.length > MAX_TAG_LEN) return buf.length
  const lower = tail.toLowerCase()
  return TAG_LITERALS.some((t) => t.startsWith(lower)) ? lt : buf.length
}

export interface ReasoningSplit {
  /** Text belonging to the answer. */
  content: string
  /** Text belonging to the model's scratchpad. */
  reasoning: string
}

/**
 * Incremental splitter. Feed it deltas; it returns the portion of each that is
 * safely classifiable. Call `flush()` once the stream ends to drain whatever
 * was held back for a possible tag.
 */
export class ReasoningSplitter {
  private buf = ''
  private inReasoning = false

  /** True if any reasoning tag was seen. Lets callers tell "no reasoning" from
   *  "reasoning we filtered". */
  sawReasoning = false

  push(chunk: string): ReasoningSplit {
    this.buf += chunk
    let content = ''
    let reasoning = ''

    for (;;) {
      const m = TAG_RE.exec(this.buf)
      if (!m) break
      const before = this.buf.slice(0, m.index)
      if (this.inReasoning) reasoning += before
      else content += before
      // A stray `</think>` with no opener still means "reasoning ends here",
      // which is what several finetunes emit after a prefilled opener.
      this.inReasoning = m[1] !== '/'
      this.sawReasoning = true
      this.buf = this.buf.slice(m.index + m[0].length)
    }

    const safe = safeEmitLength(this.buf)
    if (safe > 0) {
      const emit = this.buf.slice(0, safe)
      if (this.inReasoning) reasoning += emit
      else content += emit
      this.buf = this.buf.slice(safe)
    }
    return { content, reasoning }
  }

  /** Drain the hold-back buffer. Anything still pending was not a tag. */
  flush(): ReasoningSplit {
    const rest = this.buf
    this.buf = ''
    if (!rest) return { content: '', reasoning: '' }
    return this.inReasoning ? { content: '', reasoning: rest } : { content: rest, reasoning: '' }
  }
}

/**
 * Wrap a provider stream so inline reasoning tags in `delta` events become
 * `reasoning` events. Events of every other type pass through untouched.
 */
export async function* splitReasoning(
  source: AsyncIterable<StreamEvent>,
): AsyncIterable<StreamEvent> {
  const splitter = new ReasoningSplitter()
  for await (const ev of source) {
    if (ev.type !== 'delta') {
      // Drain before any terminal event so held-back text is not lost to an
      // early return in the consumer.
      if (ev.type === 'done' || ev.type === 'error') {
        const tail = splitter.flush()
        if (tail.reasoning) yield { type: 'reasoning', text: tail.reasoning }
        if (tail.content) yield { type: 'delta', text: tail.content }
      }
      yield ev
      continue
    }
    const { content, reasoning } = splitter.push(ev.text)
    if (reasoning) yield { type: 'reasoning', text: reasoning }
    if (content) yield { type: 'delta', text: content }
  }
  const tail = splitter.flush()
  if (tail.reasoning) yield { type: 'reasoning', text: tail.reasoning }
  if (tail.content) yield { type: 'delta', text: tail.content }
}
