import { useState, useRef, useEffect } from 'react'
import { useSettingsStore } from '../store/settingsStore'
import { useAnchoredMenu } from '../util/useAnchoredMenu'
import { ChevronDownIcon, RobotIcon } from './icons'

export function AgentSelector({
  activeId,
  onSelect,
}: {
  activeId: string | null
  onSelect: (id: string | null) => void
}) {
  const agents = useSettingsStore((s) => s.settings.agents.filter((a) => a.enabled))
  const [open, setOpen] = useState(false)
  const containerRef = useRef<HTMLDivElement>(null)
  const { anchorRef, menuRef } = useAnchoredMenu<HTMLButtonElement>(open, { width: 256 })
  const active = agents.find((a) => a.id === activeId)
  // Stable signature of enabled agent ids, so the cleanup effect doesn't fire
  // on every render just because `.filter()` returns a new array reference.
  const agentIds = agents.map((a) => a.id).join('|')

  // Clear the active selection if the agent was disabled or removed.
  useEffect(() => {
    if (activeId && !agentIds.split('|').includes(activeId)) {
      onSelect(null)
    }
  }, [activeId, agentIds, onSelect])

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

  if (agents.length === 0) return null

  return (
    <div className="relative" ref={containerRef}>
      <button
        ref={anchorRef}
        className={`btn h-7 px-2 text-xs ${active ? 'btn-primary' : 'btn-ghost'}`}
        onClick={() => setOpen((v) => !v)}
        aria-haspopup="menu"
        aria-expanded={open}
        aria-label="Select agent"
        title={
          active
            ? `${active.name}: ${active.description || 'No description'}\nModel: ${active.model}`
            : 'Select an agent'
        }
      >
        <RobotIcon size={14} />
        <span className="max-w-[80px] truncate">{active?.name ?? 'Agent'}</span>
        <ChevronDownIcon size={14} />
      </button>
      {open && (
        <div
          ref={menuRef}
          className="menu-floating panel z-30 overflow-y-auto p-1"
          role="menu"
          aria-label="Agents"
        >
          <button
            role="menuitem"
            className={`flex w-full flex-col rounded-lg px-2 py-1.5 text-left text-xs ${
              !active ? 'bg-accent-soft text-accent-fg' : 'hover:bg-surface-hover'
            }`}
            onClick={() => {
              onSelect(null)
              setOpen(false)
            }}
          >
            <span className="font-medium">No agent</span>
            <span className="text-[0.68rem] text-faint">Use the selected model directly.</span>
          </button>
          {agents.map((a) => (
            <button
              key={a.id}
              role="menuitem"
              className={`flex w-full flex-col rounded-lg px-2 py-1.5 text-left text-xs ${
                active?.id === a.id ? 'bg-accent-soft text-accent-fg' : 'hover:bg-surface-hover'
              }`}
              onClick={() => {
                onSelect(a.id)
                setOpen(false)
              }}
              title={`${a.description || 'No description'}\nModel: ${a.model}\nSource: ${a.source}`}
            >
              <span className="font-medium">{a.name}</span>
              <span className="text-[0.68rem] text-faint">{a.model}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  )
}
