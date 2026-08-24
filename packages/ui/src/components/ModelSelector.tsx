import { useEffect, useMemo, useRef, useState } from 'react'
import type { ModelInfo } from '@openofficellm/shared'
import { useModelsStore } from '../store/modelsStore'
import { useUiStore } from '../store/uiStore'
import { ChevronDownIcon, ChevronRightIcon, SearchIcon, StarIcon, RefreshIcon } from './icons'
import { loadPersisted, savePersisted } from '../store/persist'
import { useAnchoredMenu } from '../util/useAnchoredMenu'

const EXPANDED_KEY = 'models.expandedProviders'

function isStringArrayOrNull(v: unknown): v is string[] | null {
  return v === null || (Array.isArray(v) && v.every((x) => typeof x === 'string'))
}

export interface ModelSelectorProps {
  /** Render as a compact bottom-bar button instead of the full header control. */
  compact?: boolean
}

export function ModelSelector({ compact }: ModelSelectorProps) {
  const models = useModelsStore((s) => s.models)
  const hidden = useModelsStore((s) => s.hidden)
  const favorites = useModelsStore((s) => s.favorites)
  const providers = useModelsStore((s) => s.providers)
  const selected = useModelsStore((s) => s.selectedModelId)
  const select = useModelsStore((s) => s.selectModel)
  const toggleFavorite = useModelsStore((s) => s.toggleFavorite)
  const load = useModelsStore((s) => s.load)
  const loading = useModelsStore((s) => s.loading)
  const [open, setOpen] = useState(false)
  const [query, setQuery] = useState('')
  // `null` means "the user has not chosen yet" — distinct from "chose to
  // collapse everything", which is an empty array.
  const [expanded, setExpanded] = useState<string[] | null>(() =>
    loadPersisted<string[] | null>(EXPANDED_KEY, null, isStringArrayOrNull),
  )
  const containerRef = useRef<HTMLDivElement>(null)
  const searchInputRef = useRef<HTMLInputElement>(null)
  const { anchorRef: triggerRef, menuRef } = useAnchoredMenu<HTMLButtonElement>(open, {
    width: 290,
    align: 'start',
  })
  const modelSelectorHint = useUiStore((s) => s.modelSelectorHint)

  // External "pick another model" signal from the error banner — open the
  // dropdown when the hint counter increments. Use a ref to capture the
  // previous value so we only open on a change, not on initial mount.
  const prevHint = useRef(0)
  useEffect(() => {
    if (modelSelectorHint !== prevHint.current && modelSelectorHint > 0) {
      prevHint.current = modelSelectorHint
      setOpen(true)
    }
  }, [modelSelectorHint])

  const visible = useMemo(() => {
    if (hidden.length === 0) return models
    const hiddenSet = new Set(hidden)
    return models.filter((m) => !hiddenSet.has(m.id))
  }, [models, hidden])

  const q = query.trim().toLowerCase()

  const grouped = useMemo(() => {
    const byProvider = new Map<string, ModelInfo[]>()
    for (const m of visible) {
      const arr = byProvider.get(m.providerId) ?? []
      arr.push(m)
      byProvider.set(m.providerId, arr)
    }
    const entries = Array.from(byProvider.entries())
    const filtered = q
      ? entries
          .map(
            ([pid, arr]) =>
              [
                pid,
                arr.filter(
                  (m) => m.name.toLowerCase().includes(q) || m.id.toLowerCase().includes(q),
                ),
              ] as const,
          )
          .filter(([, arr]) => arr.length > 0)
      : entries.map(([pid, arr]) => [pid, arr] as const)
    return filtered.sort((a, b) => a[0].localeCompare(b[0]))
  }, [visible, q])

  const favoriteModels = useMemo(
    () =>
      favorites.map((id) => visible.find((m) => m.id === id)).filter((m): m is ModelInfo => !!m),
    [favorites, visible],
  )

  const providerInfo = (id: string) => providers.find((p) => p.id === id)
  const selectedModel = models.find((m) => m.id === selected) ?? null

  // Until the user touches a group, the provider holding the current model is
  // open — and a lone provider is opened too, since collapsing the only group
  // would leave the menu empty.
  const openByDefault = useMemo(() => {
    if (grouped.length === 1) return [grouped[0][0]]
    return selectedModel ? [selectedModel.providerId] : []
  }, [grouped, selectedModel])
  const openGroups = expanded ?? openByDefault

  // A search is a request to see the matches, so it opens every group that has
  // one without disturbing the user's choices.
  const isExpanded = (pid: string) => q !== '' || openGroups.includes(pid)

  const toggleGroup = (pid: string) => {
    const next = openGroups.includes(pid)
      ? openGroups.filter((p) => p !== pid)
      : [...openGroups, pid]
    savePersisted(EXPANDED_KEY, next)
    setExpanded(next)
  }

  const choose = (id: string) => {
    select(id)
    setOpen(false)
  }

  // Close on outside click and Escape. Focus the search input on open and
  // return focus to the trigger on close.
  useEffect(() => {
    if (!open) return
    const trigger = triggerRef.current
    const onPointerDown = (e: PointerEvent) => {
      const target = e.target as Node
      if (
        containerRef.current &&
        !containerRef.current.contains(target) &&
        !menuRef.current?.contains(target)
      ) {
        setOpen(false)
      }
    }
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation()
        setOpen(false)
      }
    }
    document.addEventListener('pointerdown', onPointerDown)
    document.addEventListener('keydown', onKeyDown)
    // Defer focus to next tick so the input is mounted.
    const t = setTimeout(() => searchInputRef.current?.focus(), 0)
    return () => {
      document.removeEventListener('pointerdown', onPointerDown)
      document.removeEventListener('keydown', onKeyDown)
      clearTimeout(t)
      // Use the captured trigger ref, not triggerRef.current (which may have
      // changed by cleanup time).
      trigger?.focus()
    }
  }, [open, triggerRef, menuRef])

  // Compact lives in the composer's control strip, where a long model id has to
  // truncate rather than push the mode and send buttons out of the pane.
  const triggerClass = compact
    ? 'btn btn-ghost h-7 min-w-0 max-w-[150px] px-1.5 text-xs'
    : 'btn btn-subtle w-full justify-between px-2.5 py-1.5 text-xs'

  return (
    <div className={compact ? 'relative min-w-0' : 'relative'} ref={containerRef}>
      <button
        ref={triggerRef}
        className={triggerClass}
        onClick={() => setOpen((v) => !v)}
        aria-haspopup="listbox"
        aria-expanded={open}
        aria-label="Select model"
        title={
          selectedModel ? `${selectedModel.providerName} · ${selectedModel.name}` : 'Select a model'
        }
      >
        <span className="flex min-w-0 items-center gap-1.5 truncate">
          <span className="truncate">{selectedModel?.name ?? 'Select a model'}</span>
          {selectedModel && !compact && (
            <span className="text-[0.6rem] text-faint">{selectedModel.providerName}</span>
          )}
        </span>
        <ChevronDownIcon size={14} className="shrink-0" />
      </button>
      {open && (
        <div
          ref={menuRef}
          // The search row stays pinned: the panel itself does not scroll, the
          // model list inside it does.
          className="menu-floating panel z-30 flex flex-col overflow-hidden p-2"
          role="listbox"
          aria-label="Models"
        >
          <div className="mb-2 flex shrink-0 items-center gap-1.5">
            <div className="relative flex-1">
              <SearchIcon
                size={13}
                className="pointer-events-none absolute left-2 top-1/2 -translate-y-1/2 text-faint"
              />
              <input
                ref={searchInputRef}
                className="field h-8 pl-7 text-xs"
                placeholder="Search models"
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                aria-label="Search models"
              />
            </div>
            <button
              className="btn btn-ghost h-8 w-8"
              onClick={() => void load(true)}
              disabled={loading}
              aria-label="Refresh models"
              title="Refresh"
            >
              <RefreshIcon size={14} />
            </button>
          </div>
          <div className="min-h-0 flex-1 overflow-y-auto">
            {grouped.length === 0 && (
              <div className="p-3 text-center text-xs text-faint">
                {loading
                  ? 'Loading…'
                  : models.length > 0
                    ? 'Nothing matches. Hidden models can be re-enabled in Settings.'
                    : 'No models. Try refreshing.'}
              </div>
            )}

            {favoriteModels.length > 0 && q === '' && (
              <div className="mb-1.5" role="group" aria-label="Favorites">
                <div className="flex items-center gap-1.5 px-1 py-0.5 text-[0.7rem] uppercase tracking-wide text-muted">
                  <StarIcon size={11} filled className="text-accent" />
                  Favorites
                </div>
                <div className="space-y-0.5">
                  {favoriteModels.map((m) => (
                    <ModelRow
                      key={`fav-${m.id}`}
                      model={m}
                      showProvider
                      favorite
                      selected={m.id === selected}
                      onSelect={() => choose(m.id)}
                      onFavorite={() => toggleFavorite(m.id)}
                    />
                  ))}
                </div>
              </div>
            )}

            {grouped.map(([pid, arr]) => {
              const info = providerInfo(pid)
              const unreachable = info && !info.reachable
              const expandedNow = isExpanded(pid)
              return (
                <ProviderGroup
                  key={pid}
                  label={info?.name ?? pid}
                  count={arr.length}
                  badge={info?.kind === 'local' ? 'Local · no cost' : info?.kind}
                  muted={!!unreachable}
                  reason={unreachable ? 'unreachable' : undefined}
                  holdsSelection={arr.some((m) => m.id === selected)}
                  expanded={expandedNow}
                  onToggle={() => toggleGroup(pid)}
                >
                  {arr.map((m) => (
                    <ModelRow
                      key={m.id}
                      model={m}
                      favorite={favorites.includes(m.id)}
                      selected={m.id === selected}
                      onSelect={() => choose(m.id)}
                      onFavorite={() => toggleFavorite(m.id)}
                    />
                  ))}
                </ProviderGroup>
              )
            })}
          </div>
        </div>
      )}
    </div>
  )
}

function ProviderGroup({
  label,
  count,
  badge,
  muted,
  reason,
  holdsSelection,
  expanded,
  onToggle,
  children,
}: {
  label: string
  count: number
  badge?: string
  muted?: boolean
  reason?: string
  holdsSelection: boolean
  expanded: boolean
  onToggle: () => void
  children: React.ReactNode
}) {
  return (
    <div className="mb-1.5" role="group" aria-label={label}>
      <button
        type="button"
        className={`flex w-full items-center gap-1 rounded px-1 py-1 text-left text-[0.7rem] uppercase tracking-wide hover:bg-surface-hover ${
          muted ? 'text-faint' : 'text-muted'
        }`}
        onClick={onToggle}
        aria-expanded={expanded}
      >
        {expanded ? (
          <ChevronDownIcon size={12} className="shrink-0" />
        ) : (
          <ChevronRightIcon size={12} className="shrink-0" />
        )}
        <span className="truncate">{label}</span>
        {reason && <span className="shrink-0 text-danger">· {reason}</span>}
        {/* A collapsed group that holds the running model still has to be
            findable, so it keeps a marker of its own. */}
        {holdsSelection && !expanded && (
          <span className="h-1.5 w-1.5 shrink-0 rounded-full bg-accent" aria-hidden />
        )}
        <span className="ml-auto shrink-0 tabular-nums text-faint">{count}</span>
        {badge && (
          <span
            className={`shrink-0 rounded px-1.5 py-0.5 text-[0.6rem] ${
              badge.startsWith('Local') ? 'bg-ok/15 text-ok' : 'bg-surface-muted text-muted'
            }`}
          >
            {badge}
          </span>
        )}
      </button>
      {expanded && <div className="space-y-0.5">{children}</div>}
    </div>
  )
}

function ModelRow({
  model,
  showProvider,
  favorite,
  selected,
  onSelect,
  onFavorite,
}: {
  model: ModelInfo
  showProvider?: boolean
  favorite: boolean
  selected: boolean
  onSelect: () => void
  onFavorite: () => void
}) {
  const { name, kind, contextWindow, sizeBytes, quantization, providerName } = model
  return (
    <div
      className={`group flex cursor-pointer items-center justify-between rounded px-2 py-1 text-xs ${
        selected ? 'bg-accent-soft text-accent-fg' : 'hover:bg-surface-hover'
      }`}
      role="option"
      aria-selected={selected}
      onClick={onSelect}
      onKeyDown={(e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault()
          onSelect()
        }
      }}
      tabIndex={0}
    >
      <span className="flex min-w-0 flex-1 items-center gap-1.5 text-left">
        <span className="truncate">{name}</span>
        {showProvider && <span className="shrink-0 text-[0.65rem] text-faint">{providerName}</span>}
        {sizeBytes && (
          <span className="shrink-0 text-[0.65rem] text-faint">
            {(sizeBytes / 1e9).toFixed(1)}GB
          </span>
        )}
        {quantization && <span className="shrink-0 text-[0.65rem] text-faint">{quantization}</span>}
        {contextWindow && (
          <span className="shrink-0 text-[0.65rem] text-faint">
            {(contextWindow / 1000).toFixed(0)}k ctx
          </span>
        )}
        {kind === 'local' && (
          <span className="shrink-0 rounded bg-ok/15 px-1 text-[0.6rem] text-ok">local</span>
        )}
      </span>
      <button
        className={`icon-btn h-5 w-5 shrink-0 opacity-0 ${
          favorite ? 'text-accent opacity-100' : 'group-hover:opacity-100'
        }`}
        onClick={(e) => {
          e.stopPropagation()
          onFavorite()
        }}
        aria-label={favorite ? `Unfavorite ${name}` : `Favorite ${name}`}
        title={favorite ? 'Remove from favorites' : 'Add to favorites'}
      >
        <StarIcon size={12} filled={favorite} />
      </button>
    </div>
  )
}
