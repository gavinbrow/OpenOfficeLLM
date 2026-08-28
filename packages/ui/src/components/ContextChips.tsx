// What the model can see, and a control to change it (P3.13).
//
// This bar exists as much to reassure as to configure. The most common failure
// report for an Office assistant is "it can't see my document" — often when it
// can. Naming the attached scope and its token cost makes the answer visible
// before the user has to ask.

import { useCallback, useEffect, useState } from 'react'
import type { ContextScope } from '@openofficellm/shared'
import { useContextStore } from '../store/contextStore'
import { useSettingsStore } from '../store/settingsStore'
import { getHost, getAdapter, settingsHost } from '../host/bridge'
import { SCOPES_FOR_HOST } from '../host/adapter'
import { CloseIcon, RefreshIcon, FileIcon } from './icons'

const SCOPE_LABEL: Record<ContextScope, string> = {
  none: 'Nothing',
  selection: 'Selection',
  paragraph: 'Current paragraph',
  document: 'Whole document',
  sheet: 'Whole sheet',
  range: 'Selected range',
  page: 'Whole page',
}

export function ContextChips() {
  const items = useContextStore((s) => s.items)
  const remove = useContextStore((s) => s.remove)
  const removeAttachment = useContextStore((s) => s.removeAttachment)
  const settings = useSettingsStore((s) => s.settings)
  const save = useSettingsStore((s) => s.save)

  const host = getHost()
  const [preview, setPreview] = useState<{ scope: ContextScope; tokens: number } | null>(null)
  const [checking, setChecking] = useState(false)

  const scope = settings.defaultContext[settingsHost(host)]

  const refresh = useCallback(async () => {
    const adapter = getAdapter()
    if (!adapter || scope === 'none') {
      setPreview(null)
      return
    }
    setChecking(true)
    try {
      const ctx = await adapter.getContext(scope)
      // getContext reports the scope it actually used, which may differ from
      // the requested one — asking for the selection with nothing selected
      // falls back to the document, and the chip should say so.
      setPreview({ scope: ctx.scope, tokens: ctx.tokenEstimate ?? 0 })
    } catch {
      setPreview(null)
    } finally {
      setChecking(false)
    }
  }, [scope])

  // Read the document when the scope changes. Deliberately not calling
  // `refresh` here: it sets state synchronously before awaiting, which is a
  // cascading render inside an effect. Reading the document is the external
  // system this effect subscribes to, so state only moves in the continuation.
  useEffect(() => {
    let cancelled = false
    const adapter = getAdapter()
    if (!adapter || scope === 'none') return
    adapter
      .getContext(scope)
      .then((ctx) => {
        if (!cancelled) setPreview({ scope: ctx.scope, tokens: ctx.tokenEstimate ?? 0 })
      })
      .catch(() => {
        if (!cancelled) setPreview(null)
      })
    return () => {
      cancelled = true
    }
  }, [scope])

  const onScopeChange = (next: ContextScope) => {
    const key = host === 'excel' ? 'excel' : 'word'
    void save({ ...settings, defaultContext: { ...settings.defaultContext, [key]: next } })
  }

  // Outside Office there is no document to attach, and offering scopes that
  // cannot resolve would be a lie.
  if (host === 'none' && items.length === 0) return null

  const overBudget = (preview?.tokens ?? 0) > settings.contextTrimWarningTokens

  return (
    <div
      className="flex shrink-0 flex-wrap items-center gap-1.5 border-t border-surface-border px-2 py-1.5 text-xs"
      aria-label="Included context"
    >
      {host !== 'none' && (
        <>
          <label className="flex items-center gap-1.5 text-muted">
            <span className="text-faint">Sees:</span>
            {/* `.select-bare` rather than a hand-rolled set of utilities: the
                option list is drawn by the OS, so the explicit option colours
                in index.css are the only thing standing between dark mode and
                white text on a white popup. */}
            <select
              className="select-bare"
              value={scope}
              onChange={(e) => onScopeChange(e.target.value as ContextScope)}
              aria-label="What the assistant can read"
            >
              {SCOPES_FOR_HOST[settingsHost(host)].map((s) => (
                <option key={s} value={s}>
                  {SCOPE_LABEL[s]}
                </option>
              ))}
            </select>
          </label>

          {scope !== 'none' && (
            <span
              className={`text-[0.68rem] ${overBudget ? 'text-warn' : 'text-faint'}`}
              title={
                overBudget
                  ? 'Larger than the context warning threshold in settings.'
                  : 'Approximate size of what will be sent.'
              }
            >
              {checking
                ? 'reading…'
                : preview
                  ? `${SCOPE_LABEL[preview.scope].toLowerCase()} · ~${preview.tokens.toLocaleString()} tokens`
                  : 'nothing to read'}
            </span>
          )}

          <button
            className="icon-btn h-5 w-5 text-faint hover:text-fg"
            onClick={() => void refresh()}
            aria-label="Re-read the document"
            title="Re-read the document"
          >
            <RefreshIcon size={11} />
          </button>
        </>
      )}

      {items.map((item) => {
        const isAttachment =
          item.kind === 'text-attachment' || item.kind === 'image-attachment'
        return (
          <span
            key={item.id}
            className="chip"
            title={
              isAttachment
                ? `${item.kind === 'image-attachment' ? 'Image' : 'File'}: ${item.label} · ~${item.tokenEstimate} tokens`
                : `${item.scope} · ~${item.tokenEstimate} tokens`
            }
          >
            {isAttachment && <FileIcon size={11} />}
            <span className="max-w-[160px] truncate">{item.label}</span>
            <span className="text-[0.6rem] text-faint">·{item.tokenEstimate}t</span>
            <button
              className="ml-0.5 inline-flex h-4 w-4 items-center justify-center rounded-full hover:bg-surface-border"
              onClick={() =>
                isAttachment ? void removeAttachment(item.id) : remove(item.id)
              }
              aria-label={`Remove ${item.label} from context`}
            >
              <CloseIcon size={10} />
            </button>
          </span>
        )
      })}
    </div>
  )
}
