// Side panel entrypoint.
//
// Compare packages/addin/src/main.tsx. The steps are the same — work out what
// we are attached to, register a Shell, adopt that surface's conversation,
// mount — with two extra ones that only a cross-origin shell needs: find the
// host service, and get a token from it.

import { StrictMode, useCallback, useEffect, useState } from 'react'
import { createRoot } from 'react-dom/client'
import '@openofficellm/ui/index.css'
import { App, ErrorBoundary, registerShell, useChatStore, getDocumentKey } from '@openofficellm/ui'
import { browserShell } from './browser/shell'
import {
  attachToActiveTab,
  getHost,
  getTab,
  hasAccessTo,
  isScriptableUrl,
  watchColorScheme,
  watchTabs,
} from './browser/bootstrap'
import { pairWithHost, type PairingResult } from './browser/pairing'
import { PairingScreen, SiteAccessBar } from './components/PairingGate'

function Panel() {
  const [pairing, setPairing] = useState<PairingResult | null>(null)
  // Bumped to retry pairing. A counter rather than a callback so the effect
  // owns the request and can abandon its own result on unmount.
  const [attempt, setAttempt] = useState(0)
  // Bumped whenever the attached tab changes, to re-evaluate the access bar.
  const [tabEpoch, setTabEpoch] = useState(0)
  const [needsAccess, setNeedsAccess] = useState(false)

  const retry = useCallback(() => {
    setPairing(null)
    setAttempt((n) => n + 1)
  }, [])

  useEffect(() => {
    let cancelled = false
    void pairWithHost().then((result) => {
      if (!cancelled) setPairing(result)
    })
    return () => {
      cancelled = true
    }
  }, [attempt])

  // Follow the active tab. `changed` is false for a re-check that landed on the
  // same tab, which is what keeps an ordinary window focus event from resetting
  // a conversation mid-turn.
  useEffect(() => {
    const adopt = () => {
      useChatStore.getState().adoptDocument(getDocumentKey())
      setTabEpoch((n) => n + 1)
    }
    void attachToActiveTab().then(adopt)
    return watchTabs((changed) => {
      if (changed) adopt()
      else setTabEpoch((n) => n + 1)
    })
  }, [])

  // Whether to offer the site-access prompt: the tab is one we could read, but
  // have not been granted. Recomputed per tab change rather than held in the
  // bootstrap module, because the answer changes when the user clicks Allow.
  //
  // Both branches resolve through a promise so the effect never sets state
  // synchronously — a sync setState here would re-render before the effect
  // that scheduled it had finished committing.
  useEffect(() => {
    let cancelled = false
    const tab = getTab()
    const eligible = tab !== null && isScriptableUrl(tab.url) && getHost() !== 'browser'
    const granted = eligible ? hasAccessTo(tab.url) : Promise.resolve(true)
    void granted.then((has) => {
      if (!cancelled) setNeedsAccess(!has)
    })
    return () => {
      cancelled = true
    }
  }, [tabEpoch])

  useEffect(() => {
    watchColorScheme(() => {
      // The class is applied inside watchColorScheme; this callback exists so a
      // future preference override has somewhere to hook in.
    })
  }, [])

  if (pairing === null) {
    return (
      <div className="flex h-full items-center justify-center text-xs text-faint">
        Connecting to the host service…
      </div>
    )
  }

  if (!pairing.ok) {
    return <PairingScreen result={pairing} onRetry={retry} />
  }

  return (
    <div className="flex h-full w-full flex-col overflow-hidden">
      {needsAccess && (
        <SiteAccessBar
          onGranted={() => {
            void attachToActiveTab().then(() => {
              setNeedsAccess(false)
              setTabEpoch((n) => n + 1)
            })
          }}
        />
      )}
      <div className="min-h-0 flex-1">
        <App />
      </div>
    </div>
  )
}

registerShell(browserShell)

const rootEl = document.getElementById('root')
if (!rootEl) throw new Error('root element missing')
createRoot(rootEl).render(
  <StrictMode>
    <ErrorBoundary>
      <Panel />
    </ErrorBoundary>
  </StrictMode>,
)
