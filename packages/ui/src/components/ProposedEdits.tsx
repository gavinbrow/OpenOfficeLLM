// Review card for edits staged in `propose` mode (P4.18).
//
// This is the surface that makes "propose" a real mode rather than a label. If
// the model has staged changes, nothing has touched the document until the user
// presses Apply here.

import type { Edit, TextFormatting } from '@openofficellm/shared'
import { useProposalStore } from '../store/proposalStore'
import { shell } from '../host/bridge'
import { CheckIcon, CloseIcon } from './icons'

/** Formatting in the user's terms, via the shell's own vocabulary when it has
 *  one. Word turns `{ styleBuiltIn: 'Heading1' }` into "Heading 1"; a shell
 *  without that mapping falls back to the raw keys, which is uglier but still
 *  tells the user what they are approving. */
function describeFormatting(formatting: TextFormatting): string {
  const describe = shell().describeFormatting
  if (describe) return describe(formatting)
  return Object.entries(formatting)
    .map(([k, v]) => `${k}: ${String(v)}`)
    .join(', ')
}

/** What the change will do, in the user's terms. A proposal the user cannot
 *  read is one they will approve without understanding, which defeats the
 *  point of propose mode. */
function previewOf(edit: Edit): string {
  switch (edit.kind) {
    case 'replaceSelection':
    case 'insertAfter':
    case 'insertBefore':
      return edit.text
    case 'replaceRange':
      return edit.text
    case 'addComment':
      return edit.text
    case 'setCellValues':
      return edit.cells.map((c) => `${c.cell} = ${String(c.value)}`).join('\n')
    case 'setCellFormulas':
      return edit.cells.map((c) => `${c.cell} = ${c.formula}`).join('\n')
    case 'applyFormatting':
      return Object.entries(edit.formatting)
        .map(([k, v]) => `${k}: ${String(v)}`)
        .join(', ')
    case 'formatText':
      return describeFormatting(edit.formatting)
    case 'insertParagraph':
      return edit.style ? `${edit.style}: ${edit.text}` : edit.text
    case 'deleteParagraphs':
      return `Paragraph${edit.paragraphs.length === 1 ? '' : 's'} ${edit.paragraphs.join(', ')}`
    case 'setList':
      return edit.listType === 'none'
        ? `Remove list formatting, paragraphs ${edit.paragraphs.join(', ')}`
        : `${edit.listType === 'bullet' ? 'Bulleted' : 'Numbered'} list, level ${edit.level ?? 1}, paragraphs ${edit.paragraphs.join(', ')}`
    case 'insertTable':
      return edit.rows
        .slice(0, 6)
        .map((r) => r.join(' | '))
        .join('\n')
    case 'insertBreak':
      return `${edit.breakType} break`
    case 'insertHyperlink':
      return edit.text ? `${edit.text} → ${edit.url}` : edit.url
    case 'replaceAll':
      return `"${edit.find}" → "${edit.replace}"`
    case 'setHeaderFooter':
      return edit.pageNumber ? `${edit.text} [page number]`.trim() : edit.text
    case 'setPageSetup': {
      const { orientation, pageSize, margins } = edit.setup
      const marginText = margins
        ? Object.entries(margins)
            .filter(([, v]) => v !== undefined)
            .map(([k, v]) => `${k} ${v}"`)
            .join(', ')
        : ''
      return [orientation, pageSize, marginText].filter(Boolean).join(', ')
    }
    default:
      return ''
  }
}

export function ProposedEdits() {
  const proposals = useProposalStore((s) => s.proposals)
  const applying = useProposalStore((s) => s.applying)
  const lastResult = useProposalStore((s) => s.lastResult)
  const applyAll = useProposalStore((s) => s.applyAll)
  const applyOne = useProposalStore((s) => s.applyOne)
  const discard = useProposalStore((s) => s.discard)
  const discardAll = useProposalStore((s) => s.discardAll)
  const clearResult = useProposalStore((s) => s.clearResult)

  if (proposals.length === 0) {
    if (!lastResult) return null
    return (
      <div
        className={`mx-2 mb-1.5 flex items-center gap-2 rounded-lg border px-2 py-1.5 text-xs ${
          lastResult.ok
            ? 'border-ok/40 bg-ok/10 text-ok'
            : 'border-danger/40 bg-danger/10 text-danger'
        }`}
        role="status"
      >
        <span className="flex-1">{lastResult.message}</span>
        <button className="icon-btn h-5 w-5" onClick={clearResult} aria-label="Dismiss">
          <CloseIcon size={11} />
        </button>
      </div>
    )
  }

  return (
    <section
      className="mx-2 mb-1.5 rounded-xl2 border border-accent-line bg-accent-soft"
      aria-label="Proposed changes"
    >
      <header className="flex items-center justify-between px-2.5 py-1.5">
        <span className="text-xs font-semibold">
          {proposals.length} proposed change{proposals.length === 1 ? '' : 's'}
        </span>
        <span className="text-[0.68rem] text-muted">Nothing applied yet</span>
      </header>

      <ul className="max-h-[min(13rem,32vh)] overflow-y-auto px-2 pb-1">
        {proposals.map((p) => (
          <li key={p.id} className="mb-1.5 rounded-lg border border-surface-border bg-surface p-2">
            <div className="flex items-start justify-between gap-2">
              <span className="text-xs font-medium">{p.description}</span>
              <div className="flex shrink-0 gap-0.5">
                <button
                  className="icon-btn h-5 w-5 text-ok hover:bg-ok/10"
                  onClick={() => void applyOne(p.id)}
                  disabled={applying}
                  aria-label={`Apply: ${p.description}`}
                  title="Apply this change"
                >
                  <CheckIcon size={12} />
                </button>
                <button
                  className="icon-btn h-5 w-5 text-danger hover:bg-danger/10"
                  onClick={() => discard(p.id)}
                  disabled={applying}
                  aria-label={`Discard: ${p.description}`}
                  title="Discard this change"
                >
                  <CloseIcon size={12} />
                </button>
              </div>
            </div>
            <pre className="mt-1 max-h-24 overflow-y-auto whitespace-pre-wrap break-words text-[0.72rem] leading-snug text-muted">
              {previewOf(p.edit).slice(0, 600)}
            </pre>
          </li>
        ))}
      </ul>

      <footer className="flex justify-end gap-2 border-t border-accent-line px-2.5 py-1.5">
        <button
          className="btn btn-ghost px-2.5 py-1 text-xs"
          onClick={discardAll}
          disabled={applying}
        >
          Discard all
        </button>
        <button
          className="btn btn-primary px-2.5 py-1 text-xs"
          onClick={() => void applyAll()}
          disabled={applying}
        >
          {/* One batch, one undo record — see proposalStore.applyAll. */}
          {applying ? 'Applying…' : `Apply all (${proposals.length})`}
        </button>
      </footer>
    </section>
  )
}
