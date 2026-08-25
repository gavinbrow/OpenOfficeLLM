import { useEffect } from 'react'
import { useUiStore } from './store/uiStore'
import { useModelsStore } from './store/modelsStore'
import { useSettingsStore } from './store/settingsStore'
import { getHealth } from './api/client'
import { ServiceDownScreen } from './components/ServiceDownScreen'
import { HostBanner } from './components/HostBanner'
import { ChatPanel } from './components/ChatPanel'
import { Composer } from './components/Composer'
import { ContextChips } from './components/ContextChips'
import { ProposedEdits } from './components/ProposedEdits'
import { SkillBar } from './components/SkillBar'
import { UsageBar } from './components/UsageBar'
import { Sidebar, SidebarToggle } from './components/Sidebar'
import { SettingsPanel } from './components/SettingsPanel'
import { ErrorBanner, Toasts } from './components/ErrorBanner'
import { SettingsIcon, CloseIcon } from './components/icons'
import { UpdateBadge } from './components/UpdateBadge'
import { applyTheme } from './theme'

export function App() {
  const serviceDown = useUiStore((s) => s.serviceDown)
  const sidebarOpen = useUiStore((s) => s.sidebarOpen)
  const openSettings = useUiStore((s) => s.openSettings)
  const theme = useUiStore((s) => s.theme)
  const setServiceDown = useUiStore((s) => s.setServiceDown)
  const setUpdateAvailable = useUiStore((s) => s.setUpdateAvailable)
  const loadModels = useModelsStore((s) => s.load)
  const loadSettings = useSettingsStore((s) => s.load)
  const setTheme = useUiStore((s) => s.setTheme)
  const savedTheme = useSettingsStore((s) => s.settings.theme)
  const updateCountdown = useUiStore((s) => s.updateCountdown)
  const latestVersion = useUiStore((s) => s.latestVersion)
  const cancelUpdateCountdown = useUiStore((s) => s.cancelUpdateCountdown)

  useEffect(() => {
    applyTheme(theme)
  }, [theme])

  // The host-side setting is the durable choice; the UI store is what actually
  // paints. Adopt the saved value once it arrives so the pane looks the same on
  // a machine where localStorage was cleared. `undefined` means "follow Office".
  useEffect(() => {
    setTheme(savedTheme ?? 'office')
  }, [savedTheme, setTheme])

  useEffect(() => {
    let cancelled = false
    void getHealth()
      .then((h) => {
        if (cancelled) return
        if (h.status === 'ok' || h.status === 'degraded') {
          setServiceDown(false, h.version)
          setUpdateAvailable(h.updateAvailable ?? false, h.latestVersion ?? null)
          void loadModels()
          void loadSettings()
        } else {
          setServiceDown(true)
        }
      })
      .catch(() => {
        if (!cancelled) setServiceDown(true)
      })
    return () => {
      cancelled = true
    }
  }, [setServiceDown, setUpdateAvailable, loadModels, loadSettings])

  if (serviceDown) {
    return (
      <>
        <ServiceDownScreen />
        <Toasts />
      </>
    )
  }

  return (
    <div className="flex h-full w-full flex-col overflow-hidden">
      <HostBanner />
      <div className="flex min-h-0 flex-1 overflow-hidden">
        {sidebarOpen && (
          <aside
            className="w-48 shrink-0 border-r border-surface-border"
            aria-label="Conversation history"
          >
            <Sidebar />
          </aside>
        )}
        <main className="flex min-w-0 flex-1 flex-col overflow-hidden" aria-label="Chat">
          <header className="flex shrink-0 items-center gap-1.5 border-b border-surface-border p-1.5">
            <SidebarToggle />
            <div className="flex-1" />
            <UpdateBadge />
            <button
              className="btn btn-ghost h-8 w-8"
              onClick={openSettings}
              aria-label="Open settings"
              title="Settings"
            >
              <SettingsIcon size={16} />
            </button>
          </header>
          <ContextChips />
          <div className="min-h-0 flex-1 overflow-hidden">
            <ChatPanel />
          </div>
          {/* Above the composer, below the transcript: staged edits are the
              thing to act on next, so they sit where the eye lands after
              reading the answer.

              `shrink-0` on the whole bottom stack is deliberate — the
              transcript is what gives up height when this grows, and each of
              these caps its own. Without it the composer was the thing that got
              squeezed out of view. */}
          <div className="shrink-0">
            <ProposedEdits />
            <UsageBar />
            <ErrorBanner />
            <SkillBar />
            <Composer />
          </div>
        </main>
      </div>
      <SettingsPanel />
      {updateCountdown !== null && (
        <div className="fixed bottom-3 right-3 z-50 flex items-center gap-2 rounded-lg bg-surface-inverse px-3 py-2 text-xs text-fg-inverse shadow-pop">
          <span>
            Updating to v{latestVersion ?? ''} — restarting in {updateCountdown}…{' '}
          </span>
          <button
            className="icon-btn h-4 w-4"
            onClick={cancelUpdateCountdown}
            aria-label="Cancel update"
          >
            <CloseIcon size={10} />
          </button>
        </div>
      )}
      <Toasts />
    </div>
  )
}
