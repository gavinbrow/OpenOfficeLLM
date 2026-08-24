import { useModelsStore } from '../store/modelsStore'
import { useUiStore } from '../store/uiStore'
import type { DetectedHost } from '@openofficellm/shared'
import { getHost } from '../host/bridge'
import { useChatStore } from '../store/chatStore'
import { RefreshIcon } from './icons'

const WORD_SUGGESTIONS = [
  'Summarize this document',
  'Rewrite the selected paragraph more formally',
  'Proofread for clarity and grammar',
  'Draft a follow-up email about this topic',
]

const EXCEL_SUGGESTIONS = [
  'Explain the formula in the selected cell',
  'Write a formula to sum this column',
  'Find anomalies in this sheet',
  'Suggest a chart for this data',
]

const BROWSER_SUGGESTIONS = [
  'Summarize this page',
  'What is the main argument here?',
  'Pull out the key facts and figures',
  'Is anything on this page unsupported?',
]

const GENERIC_SUGGESTIONS = [
  'What can you help me with?',
  'Summarize the current context',
  'Explain a concept',
]

const SUGGESTIONS_FOR_HOST: Record<DetectedHost, string[]> = {
  word: WORD_SUGGESTIONS,
  excel: EXCEL_SUGGESTIONS,
  browser: BROWSER_SUGGESTIONS,
  // Detached: nothing is attached, so nothing host-specific is worth offering.
  none: GENERIC_SUGGESTIONS,
}

export function EmptyState() {
  const models = useModelsStore((s) => s.models)
  const loading = useModelsStore((s) => s.loading)
  const error = useModelsStore((s) => s.error)
  const reload = useModelsStore((s) => s.load)
  const openSettings = useUiStore((s) => s.openSettings)
  const send = useChatStore((s) => s.send)
  const host = getHost()

  const localModels = models.filter((m) => m.kind === 'local')
  const suggestions = SUGGESTIONS_FOR_HOST[host]

  return (
    <div className="flex h-full flex-col items-center justify-center gap-4 p-6 text-center">
      <div className="rounded-full bg-accent-soft p-3 text-accent">
        <svg
          width="28"
          height="28"
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="1.75"
          strokeLinecap="round"
          strokeLinejoin="round"
          aria-hidden
        >
          <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2Z" />
        </svg>
      </div>
      <div>
        <h1 className="text-lg font-semibold text-fg">OpenOfficeLLM</h1>
        <p className="mt-1 text-sm text-muted">
          {host === 'word'
            ? 'Your AI assistant in Word.'
            : host === 'excel'
              ? 'Your AI assistant in Excel.'
              : 'Your AI assistant for Office.'}
        </p>
      </div>

      {loading ? (
        <div className="text-xs text-faint">Looking for models…</div>
      ) : error ? (
        /* The model list failing to load and there genuinely being no models
           are different problems with the same symptom. Saying "start a local
           server" when the real cause was a failed request sends the user to
           fix something that isn't broken. */
        <div className="max-w-xs rounded-lg border border-danger/40 bg-danger/10 p-3 text-xs text-danger">
          <p>Could not reach the local host service.</p>
          <p className="mt-1 text-[0.68rem] opacity-80">{error}</p>
          <button
            className="btn btn-subtle mt-2 px-2 py-1 text-xs"
            onClick={() => void reload(true)}
          >
            <RefreshIcon size={12} /> Try again
          </button>
        </div>
      ) : models.length === 0 ? (
        <div className="max-w-xs rounded-lg border border-surface-border bg-surface-muted p-3 text-xs text-muted">
          <p>
            No models found. Start a local server (Ollama, LM Studio) or add a provider key in
            Settings.
          </p>
          <div className="mt-2 flex justify-center gap-2">
            <button className="btn btn-subtle px-2 py-1 text-xs" onClick={openSettings}>
              Open Settings
            </button>
            <button className="btn btn-ghost px-2 py-1 text-xs" onClick={() => void reload(true)}>
              <RefreshIcon size={12} /> Re-scan
            </button>
          </div>
        </div>
      ) : (
        <div className="w-full max-w-xs">
          {localModels.length > 0 && (
            <div className="mb-2 text-xs text-ok">
              ✓ {localModels.length} local model{localModels.length === 1 ? '' : 's'} detected · no
              cost
            </div>
          )}
          <div className="space-y-1.5">
            {suggestions.map((s) => (
              <button
                key={s}
                className="btn btn-subtle w-full justify-start px-3 py-2 text-left text-xs"
                onClick={() => void send(s)}
              >
                {s}
              </button>
            ))}
          </div>
        </div>
      )}
      <div className="text-[0.7rem] text-faint">Token estimates use ~4 characters per token.</div>
    </div>
  )
}
