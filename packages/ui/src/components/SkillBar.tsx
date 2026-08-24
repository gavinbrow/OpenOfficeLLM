// One-click skills above the composer (P5.5).
//
// A skill is a prompt template plus a context scope, so invoking one is just a
// send with a skillId — the host resolves the prompt and folds it into the
// system message.

import { useEffect, useRef, useState, useMemo } from 'react'
import { useSkillsStore } from '../store/skillsStore'
import { useSettingsStore } from '../store/settingsStore'
import { useChatStore } from '../store/chatStore'
import { useAnchoredMenu } from '../util/useAnchoredMenu'
import { ChevronDownIcon } from './icons'

/** Skills shown inline before the rest move into the overflow menu. */
const INLINE_COUNT = 3

export function SkillBar() {
  const allSkills = useSkillsStore((s) => s.skills)
  const load = useSkillsStore((s) => s.load)
  const disabledSkills = useSettingsStore((s) => s.settings.disabledSkills)
  const showImportedSkills = useSettingsStore((s) => s.settings.showImportedSkills)
  const send = useChatStore((s) => s.send)
  const streaming = useChatStore((s) => s.streaming)
  const [overflowOpen, setOverflowOpen] = useState(false)
  const overflowRef = useRef<HTMLDivElement | null>(null)
  const { anchorRef, menuRef } = useAnchoredMenu<HTMLButtonElement>(overflowOpen, {
    width: 224,
    maxHeight: 256,
  })

  // Derive visible skills reactively from the settings toggles so changes in
  // the Settings panel take effect as soon as settings save, without needing
  // a skills reload.
  const skills = useMemo(
    () =>
      allSkills.filter((s) => {
        if (disabledSkills.includes(s.id)) return false
        if (!showImportedSkills && s.source === 'opencode') return false
        return true
      }),
    [allSkills, disabledSkills, showImportedSkills],
  )

  useEffect(() => {
    void load()
  }, [load])

  useEffect(() => {
    if (!overflowOpen) return
    const onDown = (e: MouseEvent) => {
      if (!overflowRef.current?.contains(e.target as Node)) setOverflowOpen(false)
    }
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOverflowOpen(false)
    }
    document.addEventListener('mousedown', onDown)
    document.addEventListener('keydown', onKey)
    return () => {
      document.removeEventListener('mousedown', onDown)
      document.removeEventListener('keydown', onKey)
    }
  }, [overflowOpen])

  if (skills.length === 0) return null

  const invoke = (id: string, name: string, mode?: string) => {
    setOverflowOpen(false)
    const validMode = ['propose', 'direct', 'agentic'].includes(mode ?? '')
    // The visible user message is the skill's name, not its prompt — the
    // transcript should read "Proofread", not three paragraphs of instructions.
    void send(name, {
      skillId: id,
      mode: validMode ? (mode as import('@openofficellm/shared').EditMode) : undefined,
    })
  }

  const inline = skills.slice(0, INLINE_COUNT)
  const overflow = skills.slice(INLINE_COUNT)

  return (
    <div
      className="flex items-center gap-1 overflow-x-auto border-t border-surface-border px-2 py-1"
      aria-label="Skills"
    >
      {inline.map((s) => (
        <button
          key={s.id}
          className="chip shrink-0 hover:bg-surface-hover disabled:opacity-45"
          onClick={() => invoke(s.id, s.name, s.mode)}
          disabled={streaming}
          title={s.description || s.name}
        >
          {s.icon && <span aria-hidden>{s.icon}</span>}
          <span>{s.name}</span>
        </button>
      ))}

      {overflow.length > 0 && (
        <div className="relative shrink-0" ref={overflowRef}>
          <button
            ref={anchorRef}
            className="chip hover:bg-surface-hover"
            onClick={() => setOverflowOpen((v) => !v)}
            disabled={streaming}
            aria-expanded={overflowOpen}
            aria-haspopup="menu"
            title="More skills"
          >
            <span>More</span>
            <ChevronDownIcon size={11} />
          </button>
          {overflowOpen && (
            <div ref={menuRef} className="menu-floating panel z-30 overflow-y-auto p-1" role="menu">
              {overflow.map((s) => (
                <button
                  key={s.id}
                  className="flex w-full items-start gap-2 rounded-lg px-2 py-1.5 text-left text-xs hover:bg-surface-hover"
                  onClick={() => invoke(s.id, s.name, s.mode)}
                  role="menuitem"
                >
                  {s.icon && <span aria-hidden>{s.icon}</span>}
                  <span className="min-w-0">
                    <span className="block font-medium">{s.name}</span>
                    {s.description && (
                      <span className="block truncate text-[0.68rem] text-faint">
                        {s.description}
                      </span>
                    )}
                  </span>
                </button>
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  )
}
