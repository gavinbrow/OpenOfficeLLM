// What the panel shows before it can do anything, and the one control it needs
// that the shared UI has no concept of.
//
// Two situations, and they must not be conflated. Pairing is about the *host
// service* — the extension has no token and cannot reach any model. Site access
// is about the *current tab* — chat works fine, but page tools do not. The
// first blocks everything; the second is a strip above a working assistant.

import { useState } from 'react'
import type { PairingResult } from '../browser/pairing'
import { getTab, requestAccessTo, isScriptableUrl } from '../browser/bootstrap'

export function PairingScreen({ result, onRetry }: { result: PairingResult; onRetry: () => void }) {
  if (result.ok) return null
  const notPaired = result.reason === 'not_paired'

  return (
    <div className="flex h-full flex-col items-center justify-center gap-4 p-6 text-center">
      <div className="rounded-full bg-warn/15 p-3 text-warn">
        <svg
          width="26"
          height="26"
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="1.75"
          strokeLinecap="round"
          strokeLinejoin="round"
          aria-hidden
        >
          <rect x="3" y="11" width="18" height="11" rx="2" />
          <path d="M7 11V7a5 5 0 0 1 10 0v4" />
        </svg>
      </div>

      <div className="space-y-1">
        <h1 className="text-sm font-semibold text-fg">
          {notPaired ? 'Not paired with the host' : 'Cannot reach the host service'}
        </h1>
        <p className="text-xs leading-relaxed text-muted">
          {notPaired
            ? 'The local host service holds your model providers and their keys. It will not talk to an extension you have not explicitly allowed.'
            : result.message}
        </p>
      </div>

      {notPaired && (
        <div className="w-full space-y-2 text-left">
          <p className="text-[0.7rem] text-faint">Run this once, in a terminal:</p>
          <CopyableCommand command={`openofficellm --pair ${chrome.runtime.id}`} />
        </div>
      )}

      <button className="btn btn-primary h-8 px-3 text-xs" onClick={onRetry}>
        Try again
      </button>
    </div>
  )
}

function CopyableCommand({ command }: { command: string }) {
  const [copied, setCopied] = useState(false)
  return (
    <div className="flex items-stretch gap-1">
      <code className="min-w-0 flex-1 select-all overflow-x-auto whitespace-nowrap rounded-lg border border-surface-border bg-surface-muted px-2 py-1.5 font-mono text-[0.68rem] text-fg">
        {command}
      </code>
      <button
        className="btn btn-ghost shrink-0 px-2 text-[0.7rem]"
        onClick={() => {
          void navigator.clipboard.writeText(command).then(
            () => {
              setCopied(true)
              setTimeout(() => setCopied(false), 1500)
            },
            () => {
              // Clipboard denied. The command is select-all above, so the user
              // still has a way to take it.
            },
          )
        }}
      >
        {copied ? 'copied' : 'copy'}
      </button>
    </div>
  )
}

/**
 * The strip offering access to the current site.
 *
 * Site access is requested per-origin and on demand rather than declared up
 * front. An extension that asks for every site at install time is asking to
 * read the user's banking session, and "it only reads when you press send" is
 * not something the user can verify. Per-site, on a click they initiated, is.
 */
export function SiteAccessBar({ onGranted }: { onGranted: () => void }) {
  const [busy, setBusy] = useState(false)
  const [denied, setDenied] = useState(false)
  const tab = getTab()

  if (!tab || !isScriptableUrl(tab.url)) return null

  let origin: string
  try {
    origin = new URL(tab.url).host
  } catch {
    return null
  }

  return (
    <div className="flex shrink-0 items-center gap-2 border-b border-accent-line bg-accent-soft px-3 py-1.5 text-[0.7rem]">
      <span className="min-w-0 flex-1 truncate text-muted">
        {denied
          ? `Access to ${origin} was declined — page tools stay unavailable.`
          : `Let the assistant read ${origin}?`}
      </span>
      <button
        className="btn btn-primary h-6 shrink-0 px-2 text-[0.7rem]"
        disabled={busy}
        onClick={() => {
          setBusy(true)
          setDenied(false)
          void requestAccessTo(tab.url).then((granted) => {
            setBusy(false)
            if (granted) onGranted()
            else setDenied(true)
          })
        }}
      >
        {busy ? '…' : denied ? 'Retry' : 'Allow'}
      </button>
    </div>
  )
}
