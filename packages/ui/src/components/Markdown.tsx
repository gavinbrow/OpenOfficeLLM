// Streaming-aware markdown renderer.
//
// Hard part: while streaming, the partial text may contain an open code fence
// with no closing ```. react-markdown treats an unterminated fence as a code
// block up to EOF, which is fine visually, but we additionally:
//  - close an open fence with a sentinel ``` so the parser renders it as a
//    code block instead of a dangling literal
//  - strip the sentinel on the next delta so the content keeps growing
//
// Sanitization: rehype-sanitize with a permissive schema (code, pre, tables,
// links with safe protocols). HTML is never rendered raw.
//
// Highlighting: rehype-highlight (highlight.js, bundle-friendly). The CSS is
// in src/index.css.

import { useMemo } from 'react'
import ReactMarkdown from 'react-markdown'
import remarkGfm from 'remark-gfm'
import rehypeHighlight from 'rehype-highlight'
import rehypeSanitize, { defaultSchema } from 'rehype-sanitize'
import type { Schema } from 'hast-util-sanitize'

const sanitizeSchema: Schema = {
  ...defaultSchema,
  attributes: {
    ...defaultSchema.attributes,
    // Keep className regex-restricted for code/pre so a model can't inject
    // arbitrary Tailwind/utility classes (e.g. `class="fixed inset-0"`) that
    // could overlay the pane. Allow only language- and hljs- prefixed classes.
    code: [...(defaultSchema.attributes?.code ?? []), ['className', /^language-/]],
    pre: [...(defaultSchema.attributes?.pre ?? []), ['className', /^hljs-|language-/]],
    span: [...(defaultSchema.attributes?.span ?? []), ['className', /^hljs-/]],
    a: [...(defaultSchema.attributes?.a ?? []), 'href', 'target', 'rel'],
    th: [...(defaultSchema.attributes?.th ?? []), 'align'],
    td: [...(defaultSchema.attributes?.td ?? []), 'align'],
  },
  protocols: {
    ...defaultSchema.protocols,
    href: ['http', 'https', 'mailto'],
  },
  tagNames: Array.from(
    new Set([
      ...(defaultSchema.tagNames ?? []),
      'a',
      'code',
      'pre',
      'span',
      'table',
      'thead',
      'tbody',
      'tr',
      'th',
      'td',
    ]),
  ),
}

/** Close an unterminated ``` fence so the parser renders a code block.
 *  Tracks the fence marker length so 4-backtick fences (which GFM allows for
 *  wrapping content that contains ```) are matched correctly. */
function patchUnclosedFence(text: string): string {
  if (!text.includes('```')) return text
  const lines = text.split('\n')
  let openFenceLen = 0
  for (const line of lines) {
    const trimmed = line.trim()
    const m = trimmed.match(/^`{3,}/)
    if (!m) continue
    const len = m[0].length
    if (openFenceLen === 0) {
      openFenceLen = len
    } else if (len >= openFenceLen) {
      openFenceLen = 0
    }
  }
  if (openFenceLen > 0) return `${text}\n${'`'.repeat(openFenceLen)}`
  return text
}

interface MarkdownProps {
  content: string
  /** While streaming, we patch unterminated fences. */
  streaming?: boolean
  className?: string
}

export function Markdown({ content, streaming = false, className }: MarkdownProps) {
  const patched = useMemo(
    () => (streaming ? patchUnclosedFence(content) : content),
    [content, streaming],
  )
  return (
    <div className={className}>
      <ReactMarkdown
        remarkPlugins={[remarkGfm]}
        rehypePlugins={[
          [rehypeSanitize, sanitizeSchema],
          [rehypeHighlight, { detect: false, ignoreMissing: true }],
        ]}
        components={{
          a: (props) => (
            <a
              {...props}
              target="_blank"
              rel="noopener noreferrer"
              className="text-accent underline hover:text-accent-fg"
            />
          ),
          pre: ({ children, ...props }) => (
            <pre {...props} className="my-2 overflow-x-auto rounded-lg">
              {children}
            </pre>
          ),
          code: ({ className: cls, children, ...props }) => {
            const isInline = !cls
            return isInline ? (
              <code
                {...props}
                className="rounded bg-surface-muted px-1 py-0.5 font-mono text-[0.85em] text-accent-fg"
              >
                {children}
              </code>
            ) : (
              <code {...props} className={cls}>
                {children}
              </code>
            )
          },
          table: ({ children }) => (
            <div className="my-2 overflow-x-auto">
              <table className="w-full border-collapse text-sm">{children}</table>
            </div>
          ),
          th: ({ children }) => (
            <th className="border border-surface-border px-2 py-1 text-left font-semibold">
              {children}
            </th>
          ),
          td: ({ children }) => (
            <td className="border border-surface-border px-2 py-1 align-top">{children}</td>
          ),
        }}
      >
        {patched}
      </ReactMarkdown>
    </div>
  )
}
