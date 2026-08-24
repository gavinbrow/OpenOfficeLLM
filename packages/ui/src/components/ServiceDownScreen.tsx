import { useEffect, useState } from 'react'
import { getHealth } from '../api/client'
import { useUiStore } from '../store/uiStore'
import { PowerIcon, RefreshIcon } from './icons'

const POLL_INTERVAL = 5000

export function ServiceDownScreen() {
  const setServiceDown = useUiStore((s) => s.setServiceDown)
  const [lastError, setLastError] = useState<string>('')
  const [retryIn, setRetryIn] = useState<number>(POLL_INTERVAL / 1000)

  useEffect(() => {
    let cancelled = false
    let timer: ReturnType<typeof setTimeout> | null = null
    let ticker: ReturnType<typeof setInterval> | null = null

    const poll = async () => {
      try {
        const h = await getHealth()
        if (cancelled) return
        if (h.status === 'ok' || h.status === 'degraded') {
          setServiceDown(false, h.version)
          return
        }
      } catch (e) {
        if (cancelled) return
        setLastError((e as Error).message)
      }
      setRetryIn(POLL_INTERVAL / 1000)
      let countdown = POLL_INTERVAL / 1000
      ticker = setInterval(() => {
        countdown -= 1
        setRetryIn(Math.max(0, countdown))
      }, 1000)
      timer = setTimeout(() => {
        if (ticker) clearInterval(ticker)
        void poll()
      }, POLL_INTERVAL)
    }
    void poll()
    return () => {
      cancelled = true
      if (timer) clearTimeout(timer)
      if (ticker) clearInterval(ticker)
    }
  }, [setServiceDown])

  const retryNow = () => {
    setRetryIn(0)
    void getHealth()
      .then((h) => {
        if (h.status === 'ok' || h.status === 'degraded') setServiceDown(false, h.version)
      })
      .catch((e) => setLastError((e as Error).message))
  }

  return (
    <div
      className="flex h-full w-full flex-col items-center justify-center gap-4 overflow-y-auto p-6
        text-center"
      role="alertdialog"
      aria-labelledby="sds-title"
      aria-describedby="sds-desc"
    >
      <div className="rounded-full bg-danger/15 p-3 text-danger">
        <PowerIcon size={28} />
      </div>
      <h1 id="sds-title" className="text-lg font-semibold">
        Host service is not running
      </h1>
      <p id="sds-desc" className="max-w-xs text-sm text-muted">
        OpenOfficeLLM needs a small background service on this PC. It serves this pane, brokers
        model calls to local and cloud providers, and keeps your API keys out of the browser. It
        usually starts automatically when you log in.
      </p>
      <div className="flex flex-col gap-2">
        <button className="btn btn-primary px-4 py-2 text-sm" onClick={retryNow}>
          <RefreshIcon size={14} /> Retry now
        </button>
        <div className="text-xs text-faint">
          {retryIn > 0 ? `Retrying in ${retryIn}s…` : 'Retrying…'}
        </div>
      </div>
      <details className="max-w-xs text-left text-xs text-muted">
        <summary className="cursor-pointer text-muted">How to start it manually</summary>
        <ul className="mt-2 list-disc space-y-1 pl-4">
          <li>
            Run <code className="font-mono text-accent">npm start</code> in the project folder.
          </li>
          <li>
            Check the install with <code className="font-mono text-accent">npm run diagnose</code>.
          </li>
          <li>
            If that reports anything missing,{' '}
            <code className="font-mono text-accent">npm run setup</code> re-provisions everything.
          </li>
        </ul>
      </details>
      {lastError && (
        <div className="max-w-xs rounded border border-danger/30 bg-danger/10 p-2 text-xs text-danger">
          {lastError}
        </div>
      )}
    </div>
  )
}
