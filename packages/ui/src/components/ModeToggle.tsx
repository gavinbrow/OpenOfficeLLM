import { useState, useRef, useEffect } from 'react'
import type { EditMode } from '@openofficellm/shared'
import { useSettingsStore } from '../store/settingsStore'
import { settingsHost } from '../host/bridge'
import { loadPersisted, savePersisted } from '../store/persist'
import { useAnchoredMenu } from '../util/useAnchoredMenu'
import { ChevronDownIcon } from './icons'

const AGENTIC_CONFIRMED_KEY = 'mode.agentic.confirmed'

const MODES: { id: EditMode; label: string; hint: string }[] = [
  {
    id: 'propose',
    label: 'Propose',
    hint: 'Render a diff for review. Apply with one click. (Default.)',
  },
  {
    id: 'direct',
    label: 'Direct',
    hint: 'Apply edits immediately in one batch. A single Ctrl+Z reverts everything.',
  },
  {
    id: 'agentic',
    label: 'Agentic',
    hint: 'Multi-step tool loop. Snapshots before the run; "Revert run" undoes it all.',
  },
]

export function ModeToggle() {
  const host = settingsHost()
  const settings = useSettingsStore((s) => s.settings)
  const setDefaultMode = useSettingsStore((s) => s.setDefaultMode)
  const current = settings.defaultMode[host]
  const [open, setOpen] = useState(false)
  const [showConfirm, setShowConfirm] = useState(false)
  const confirmedAgentic = loadPersisted<boolean>(AGENTIC_CONFIRMED_KEY, false)
  const containerRef = useRef<HTMLDivElement>(null)
  const { anchorRef, menuRef } = useAnchoredMenu<HTMLButtonElement>(open, { width: 224 })

  useEffect(() => {
    if (!open) return
    const onDown = (e: MouseEvent) => {
      if (!containerRef.current?.contains(e.target as Node)) setOpen(false)
    }
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOpen(false)
    }
    document.addEventListener('mousedown', onDown)
    document.addEventListener('keydown', onKey)
    return () => {
      document.removeEventListener('mousedown', onDown)
      document.removeEventListener('keydown', onKey)
    }
  }, [open])

  const onSelect = (mode: EditMode) => {
    setOpen(false)
    if (mode === 'agentic' && !confirmedAgentic) {
      setShowConfirm(true)
      return
    }
    setDefaultMode(host, mode)
  }

  const confirmAgentic = () => {
    savePersisted(AGENTIC_CONFIRMED_KEY, true)
    setShowConfirm(false)
    setDefaultMode(host, 'agentic')
  }

  const currentLabel = MODES.find((m) => m.id === current)?.label ?? 'Mode'

  return (
    <div className="relative" ref={containerRef}>
      <button
        ref={anchorRef}
        className="btn btn-ghost h-7 px-2 text-xs"
        onClick={() => setOpen((v) => !v)}
        aria-haspopup="menu"
        aria-expanded={open}
        aria-label="Edit mode"
        title={MODES.find((m) => m.id === current)?.hint}
      >
        {currentLabel}
        <ChevronDownIcon size={14} />
      </button>
      {open && (
        <div
          ref={menuRef}
          className="menu-floating panel z-30 overflow-y-auto p-1"
          role="menu"
          aria-label="Edit mode options"
        >
          {MODES.map((m) => (
            <button
              key={m.id}
              role="menuitem"
              className={`flex w-full flex-col rounded-lg px-2 py-1.5 text-left text-xs ${
                current === m.id ? 'bg-accent-soft text-accent-fg' : 'hover:bg-surface-hover'
              }`}
              onClick={() => onSelect(m.id)}
            >
              <span className="font-medium">{m.label}</span>
              <span className="text-[0.68rem] text-faint">{m.hint}</span>
            </button>
          ))}
        </div>
      )}
      {showConfirm && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
          role="dialog"
          aria-modal="true"
          aria-labelledby="agentic-confirm-title"
        >
          <div className="panel w-full max-w-sm p-4">
            <h2 id="agentic-confirm-title" className="text-sm font-semibold">
              Enable Agentic mode?
            </h2>
            <p className="mt-2 text-xs text-muted">
              Agentic mode lets the assistant take multiple actions in a row — reading and editing
              your document across several steps. A snapshot is taken before each run so you can
              revert everything with one click. Continue?
            </p>
            <div className="mt-3 flex justify-end gap-2">
              <button
                className="btn btn-ghost px-3 py-1.5 text-xs"
                onClick={() => setShowConfirm(false)}
              >
                Cancel
              </button>
              <button className="btn btn-primary px-3 py-1.5 text-xs" onClick={confirmAgentic}>
                Enable Agentic
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
