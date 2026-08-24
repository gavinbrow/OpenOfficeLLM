import { useChatStore } from '../store/chatStore'
import { useUiStore } from '../store/uiStore'
import { useContextStore } from '../store/contextStore'
import { AlertIcon, CloseIcon } from './icons'

export function ErrorBanner() {
  const lastError = useChatStore((s) => s.lastError)
  const clearError = useChatStore((s) => s.clearError)
  const openSettings = useUiStore((s) => s.openSettings)
  const clearContext = useContextStore((s) => s.clear)
  const openModelSelectorHint = useUiStore((s) => s.openModelSelectorHint)
  if (!lastError) return null
  const isAuth = lastError.code === 'forbidden' || lastError.code === 'auth'
  const isContext = lastError.code === 'context_too_long' || lastError.code === 'context_length'
  const isModel = lastError.code === 'model_not_found' || lastError.code === 'model_unreachable'
  return (
    <div
      className="flex items-start gap-2 border-t border-danger/30 bg-danger/10 px-2 py-1.5 text-xs text-danger"
      role="alert"
    >
      <AlertIcon size={14} className="mt-0.5 shrink-0" />
      <div className="flex-1">
        <div>{lastError.message}</div>
        <div className="mt-1 flex gap-2">
          {isAuth && (
            <button
              className="underline"
              onClick={() => {
                openSettings()
              }}
            >
              Open Settings to add a key
            </button>
          )}
          {isContext && (
            <button
              className="underline"
              onClick={() => {
                // Drop all context chips so the next send is smaller; the user
                // can then retry from the composer. Previously this button just
                // dismissed the error without changing the request size.
                clearContext()
                clearError()
              }}
            >
              Drop context and dismiss
            </button>
          )}
          {isModel && (
            <button
              className="underline"
              onClick={() => {
                openModelSelectorHint()
                clearError()
              }}
            >
              Pick another model
            </button>
          )}
          {lastError.retryable && (
            <button className="underline" onClick={clearError}>
              Dismiss
            </button>
          )}
        </div>
      </div>
      <button className="icon-btn h-5 w-5 shrink-0" onClick={clearError} aria-label="Dismiss error">
        <CloseIcon size={12} />
      </button>
    </div>
  )
}

export function Toasts() {
  const toasts = useUiStore((s) => s.toasts)
  const dismiss = useUiStore((s) => s.dismissToast)
  if (toasts.length === 0) return null
  return (
    <div className="pointer-events-none fixed bottom-3 right-3 z-50 flex flex-col gap-1.5">
      {toasts.map((t) => (
        <div
          key={t.id}
          className={`pointer-events-auto flex items-center gap-2 rounded-lg px-3 py-2 text-xs shadow-pop ${
            t.kind === 'error'
              ? 'bg-danger text-fg-inverse'
              : t.kind === 'success'
                ? 'bg-ok text-fg-inverse'
                : t.kind === 'warn'
                  ? 'bg-warn text-fg-inverse'
                  : 'bg-surface-inverse text-fg-inverse'
          }`}
          role="status"
        >
          <span>{t.message}</span>
          <button className="icon-btn h-4 w-4" onClick={() => dismiss(t.id)} aria-label="Dismiss">
            <CloseIcon size={10} />
          </button>
        </div>
      ))}
    </div>
  )
}
