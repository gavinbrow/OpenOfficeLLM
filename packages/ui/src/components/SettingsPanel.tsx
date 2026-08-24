import { useEffect, useRef, useState } from 'react'
import {
  type ContextScope,
  type HostKind,
  type McpServerInfo,
  type ProviderInfo,
  type Settings,
  DEFAULT_SETTINGS,
} from '@openofficellm/shared'
import { getSurface } from '../host/bridge'
import { SCOPES_FOR_HOST } from '../host/adapter'

import { useModelsStore } from '../store/modelsStore'
import { useSettingsStore, mergeDefaults } from '../store/settingsStore'
import { useUiStore } from '../store/uiStore'
import { useMcpStore } from '../store/mcpStore'
import { useSkillsStore } from '../store/skillsStore'
import {
  putProviderKey,
  deleteProviderKey,
  testProvider,
  getSettings,
  previewOpencodeImport,
  runOpencodeImport,
  emptyImportResult,
  type ProviderTestResponse,
  type SkillDraft,
  type OpencodeImportResult,
} from '../api/client'
import {
  CloseIcon,
  CheckIcon,
  AlertIcon,
  RefreshIcon,
  PlusIcon,
  TrashIcon,
  ChevronDownIcon,
  ChevronRightIcon,
  StarIcon,
} from './icons'

export function SettingsPanel() {
  const open = useUiStore((s) => s.settingsOpen)
  if (!open) return null
  return <SettingsPanelInner />
}

function SettingsPanelInner() {
  // Which shell this is. The settings that apply are not the same in a Word
  // task pane and a browser side panel, and showing a Chrome user an "Excel
  // edit mode" row is clutter that also implies a capability they do not have.
  const surface = getSurface()
  const close = useUiStore((s) => s.closeSettings)
  const settings = useSettingsStore((s) => s.settings)
  const save = useSettingsStore((s) => s.save)
  const models = useModelsStore((s) => s.models)
  const hidden = useModelsStore((s) => s.hidden)
  const providers = useModelsStore((s) => s.providers)
  const loadModels = useModelsStore((s) => s.load)
  const [draft, setDraft] = useState<Settings>(settings)
  const [saving, setSaving] = useState(false)
  const [saveError, setSaveError] = useState<string | null>(null)
  const panelRef = useRef<HTMLDivElement>(null)

  // Escape closes; focus is trapped within the panel by intercepting Tab at
  // the boundary. Focus the panel on open and return focus to the previously
  // focused element on close.
  useEffect(() => {
    const previouslyFocused = document.activeElement as HTMLElement | null
    const panel = panelRef.current
    panel?.focus()
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation()
        close()
        return
      }
      if (e.key === 'Tab' && panel) {
        const focusables = panel.querySelectorAll<HTMLElement>(
          'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])',
        )
        if (focusables.length === 0) return
        const first = focusables[0]
        const last = focusables[focusables.length - 1]
        if (e.shiftKey && document.activeElement === first) {
          e.preventDefault()
          last.focus()
        } else if (!e.shiftKey && document.activeElement === last) {
          e.preventDefault()
          first.focus()
        }
      }
    }
    document.addEventListener('keydown', onKeyDown)
    return () => {
      document.removeEventListener('keydown', onKeyDown)
      previouslyFocused?.focus?.()
    }
  }, [close])

  const onSave = async () => {
    setSaving(true)
    setSaveError(null)
    const saved = await save(draft)
    setSaving(false)
    // Only close on success; keep the modal open with the error visible on
    // failure so the user doesn't lose their draft. Without the message the
    // panel simply refused to close and gave no reason — the usual cause is a
    // number field cleared to 0, which the host schema rejects.
    if (saved) close()
    else setSaveError(useSettingsStore.getState().error ?? 'Could not save settings.')
  }

  return (
    <div
      className="fixed inset-0 z-40 flex items-center justify-center bg-black/40 p-4"
      role="dialog"
      aria-modal="true"
      aria-labelledby="settings-title"
      onClick={close}
    >
      <div
        ref={panelRef}
        className="panel flex max-h-[88vh] w-full max-w-lg flex-col outline-none"
        tabIndex={-1}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between border-b border-surface-border p-3">
          <h2 id="settings-title" className="text-sm font-semibold">
            Settings
          </h2>
          <button className="btn btn-ghost h-7 w-7" onClick={close} aria-label="Close settings">
            <CloseIcon size={14} />
          </button>
        </div>
        <div className="flex-1 overflow-y-auto p-3">
          <section className="mb-4">
            <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-muted">
              Appearance
            </h3>
            {/* `settings.theme` existed in the schema and in the UI store but
                nothing ever connected them, so the pane could only follow
                Office's theme and the user had no way to override a bad guess.
                Applied immediately as well as saved, so the choice is visible
                before pressing Save. */}
            <label className="flex flex-col gap-1 text-xs text-muted">
              <span>Theme</span>
              <select
                className="field h-8"
                value={draft.theme ?? 'office'}
                onChange={(e) => {
                  const v = e.target.value as 'office' | 'light' | 'dark'
                  setDraft({ ...draft, theme: v === 'office' ? undefined : v })
                  useUiStore.getState().setTheme(v)
                }}
              >
                <option value="office">Follow Word / Excel</option>
                <option value="light">Light</option>
                <option value="dark">Dark</option>
              </select>
            </label>
          </section>

          <section className="mb-4">
            <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-muted">
              Default model
            </h3>
            {/* No overlaid chevron: `select.field` draws one itself, and the
                second copy sat exactly on top of the first. */}
            <select
              className="field h-9"
              value={draft.defaultModel ?? ''}
              onChange={(e) => setDraft({ ...draft, defaultModel: e.target.value || undefined })}
              aria-label="Default model"
            >
              <option value="">(none)</option>
              {/* Hidden models are not offered, but a default that is already
                  set to one stays listed so opening Settings does not silently
                  change it. */}
              {models
                .filter((m) => !hidden.includes(m.id) || m.id === draft.defaultModel)
                .map((m) => (
                  <option key={m.id} value={m.id}>
                    {m.providerName} · {m.name}
                    {hidden.includes(m.id) ? ' (hidden)' : ''}
                  </option>
                ))}
            </select>
            <button
              className="btn btn-ghost mt-1 h-7 px-2 text-xs"
              onClick={() => void loadModels(true)}
            >
              <RefreshIcon size={12} /> Refresh models
            </button>
          </section>

          {/* Office only. The browser host has no write tools, so an edit-mode
              default there would be a control that changes nothing. */}
          {surface === 'office' && (
            <section className="mb-4">
              <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-muted">
                Edit mode defaults
              </h3>
              <div className="grid grid-cols-2 gap-3">
                <ModeDefaultSelect
                  label="Word"
                  value={draft.defaultMode.word}
                  onChange={(v) =>
                    setDraft({ ...draft, defaultMode: { ...draft.defaultMode, word: v } })
                  }
                />
                <ModeDefaultSelect
                  label="Excel"
                  value={draft.defaultMode.excel}
                  onChange={(v) =>
                    setDraft({ ...draft, defaultMode: { ...draft.defaultMode, excel: v } })
                  }
                />
              </div>
            </section>
          )}

          <section className="mb-4">
            <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-muted">
              Context defaults
            </h3>
            <div className="grid grid-cols-2 gap-3">
              {surface === 'browser' ? (
                <ContextDefaultSelect
                  label="Browser"
                  host="browser"
                  value={draft.defaultContext.browser}
                  onChange={(v) =>
                    setDraft({ ...draft, defaultContext: { ...draft.defaultContext, browser: v } })
                  }
                />
              ) : (
                <>
                  <ContextDefaultSelect
                    label="Word"
                    host="word"
                    value={draft.defaultContext.word}
                    onChange={(v) =>
                      setDraft({ ...draft, defaultContext: { ...draft.defaultContext, word: v } })
                    }
                  />
                  <ContextDefaultSelect
                    label="Excel"
                    host="excel"
                    value={draft.defaultContext.excel}
                    onChange={(v) =>
                      setDraft({ ...draft, defaultContext: { ...draft.defaultContext, excel: v } })
                    }
                  />
                </>
              )}
            </div>
          </section>

          <section className="mb-4">
            <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-muted">
              Context trim warning
            </h3>
            <label className="flex items-center gap-2 text-xs text-muted">
              <span className="shrink-0">Warn above</span>
              <input
                type="number"
                min={100}
                className="field h-8 w-28"
                value={draft.contextTrimWarningTokens}
                onChange={(e) =>
                  setDraft({
                    ...draft,
                    contextTrimWarningTokens: Number(e.target.value) || 0,
                  })
                }
              />
              <span className="shrink-0">tokens</span>
            </label>
          </section>

          <section className="mb-4">
            <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-muted">
              Step cap
            </h3>
            <label className="flex items-center gap-2 text-xs text-muted">
              <input
                type="number"
                min={1}
                max={100}
                className="field h-8 w-28"
                value={draft.agenticStepCap}
                onChange={(e) =>
                  setDraft({ ...draft, agenticStepCap: Number(e.target.value) || 0 })
                }
                aria-label="Step cap"
              />
              <span className="shrink-0">tool steps per turn</span>
            </label>
            <p className="mt-1 text-[0.7rem] text-faint">
              Applies to both modes. A formatting pass reads the document, makes its edits and
              re-reads to check itself, so this needs headroom.
            </p>
          </section>

          <section className="mb-4">
            <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-muted">
              Providers
            </h3>
            <ProviderList providers={providers} />
          </section>

          <section className="mb-4">
            <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-muted">
              Responses
            </h3>
            <label className="flex items-start gap-2 text-xs text-muted">
              <input
                type="checkbox"
                className="mt-0.5"
                checked={draft.showReasoning}
                onChange={(e) => setDraft({ ...draft, showReasoning: e.target.checked })}
              />
              <span>
                Show the model’s thought process in a collapsed block.
                <span className="block text-faint">
                  Reasoning never appears in the answer itself either way — unchecking this discards
                  it entirely.
                </span>
              </span>
            </label>
          </section>

          <section className="mb-4">
            <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-muted">
              MCP servers
            </h3>
            <McpSection />
          </section>

          <section className="mb-4">
            <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-muted">
              Skills
            </h3>
            <SkillsSection draft={draft} setDraft={setDraft} />
          </section>

          <section className="mb-4">
            <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-muted">
              Import from opencode
            </h3>
            <OpencodeImportSection draft={draft} setDraft={setDraft} />
          </section>

          <section className="mb-4">
            <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-muted">
              Ollama host
            </h3>
            <label className="flex flex-col gap-1 text-xs text-muted">
              Local Ollama base URL
              <input
                type="url"
                className="field h-8 text-xs"
                placeholder="http://127.0.0.1:11434"
                value={draft.providerOptions.ollamaBaseUrl ?? ''}
                onChange={(e) =>
                  setDraft({
                    ...draft,
                    providerOptions: { ...draft.providerOptions, ollamaBaseUrl: e.target.value },
                  })
                }
              />
              <span className="text-[0.65rem] text-faint">
                Change this if Ollama runs on a different host or port.
              </span>
            </label>
          </section>

          <section className="mb-4">
            <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-muted">
              Agents
            </h3>
            <AgentsSection draft={draft} setDraft={setDraft} />
          </section>

          <button
            className="btn btn-ghost text-xs"
            onClick={() => setDraft({ ...DEFAULT_SETTINGS })}
          >
            Reset to defaults
          </button>
        </div>
        <div className="flex items-center justify-end gap-2 border-t border-surface-border p-3">
          {saveError && (
            <span
              className="mr-auto flex items-center gap-1 text-[0.7rem] text-danger"
              role="alert"
            >
              <AlertIcon size={12} /> {saveError}
            </span>
          )}
          <button className="btn btn-ghost px-3 py-1.5 text-xs" onClick={close}>
            Cancel
          </button>
          <button
            className="btn btn-primary px-3 py-1.5 text-xs"
            onClick={onSave}
            disabled={saving}
          >
            {saving ? 'Saving…' : 'Save'}
          </button>
        </div>
      </div>
    </div>
  )
}

function ModeDefaultSelect({
  label,
  value,
  onChange,
}: {
  label: string
  value: Settings['defaultMode']['word']
  onChange: (v: Settings['defaultMode']['word']) => void
}) {
  return (
    <label className="flex flex-col gap-1 text-xs text-muted">
      <span>{label}</span>
      <select
        className="field h-8"
        value={value}
        onChange={(e) => onChange(e.target.value as typeof value)}
      >
        <option value="propose">Propose</option>
        <option value="direct">Direct</option>
        <option value="agentic">Agentic</option>
      </select>
    </label>
  )
}

const SCOPE_OPTION_LABEL: Record<ContextScope, string> = {
  none: 'None',
  selection: 'Selection',
  paragraph: 'Paragraph',
  document: 'Document',
  sheet: 'Sheet',
  range: 'Range',
  page: 'Page',
}

function ContextDefaultSelect({
  label,
  host,
  value,
  onChange,
}: {
  label: string
  host: HostKind
  value: ContextScope
  onChange: (v: ContextScope) => void
}) {
  return (
    <label className="flex flex-col gap-1 text-xs text-muted">
      <span>{label}</span>
      <select
        className="field h-8"
        value={value}
        onChange={(e) => onChange(e.target.value as ContextScope)}
      >
        {/* Only the scopes this host has. Offering "Sheet" for Word was always
            wrong; offering it for a web page is nonsense. */}
        {SCOPES_FOR_HOST[host].map((s) => (
          <option key={s} value={s}>
            {SCOPE_OPTION_LABEL[s]}
          </option>
        ))}
      </select>
    </label>
  )
}

// ─── MCP ───────────────────────────────────────────────────────────────────

function McpSection() {
  const servers = useMcpStore((s) => s.servers)
  const loading = useMcpStore((s) => s.loading)
  const error = useMcpStore((s) => s.error)
  const load = useMcpStore((s) => s.load)
  const remove = useMcpStore((s) => s.remove)
  const setToolEnabled = useMcpStore((s) => s.setToolEnabled)
  const [adding, setAdding] = useState(false)

  useEffect(() => {
    void load()
  }, [load])

  return (
    <div>
      <p className="mb-2 text-xs text-faint">
        Tools from MCP servers are <strong>off by default</strong>. Enable each one you want the
        model to be able to call.
      </p>

      {error && <p className="mb-2 text-xs text-danger">{error}</p>}

      {servers.length === 0 && !loading && (
        <p className="mb-2 text-xs text-faint">No servers configured.</p>
      )}

      <ul className="space-y-2">
        {servers.map((server) => (
          <li key={server.id} className="rounded-lg border border-surface-border p-2.5">
            <div className="flex items-center justify-between gap-2">
              <div className="flex min-w-0 items-center gap-2">
                <span className="truncate text-sm font-medium">{server.name}</span>
                <StatusPill status={server.status} />
              </div>
              <div className="flex shrink-0 items-center gap-1">
                <button
                  className="btn btn-ghost h-6 px-1.5 text-[0.68rem]"
                  onClick={() => void load(true)}
                  title="Reconnect and re-read the tool list"
                >
                  <RefreshIcon size={11} />
                </button>
                <button
                  className="btn btn-danger h-6 px-1.5 text-[0.68rem]"
                  onClick={() => void remove(server.id)}
                  aria-label={`Remove ${server.name}`}
                >
                  Remove
                </button>
              </div>
            </div>

            <div className="mt-1 truncate font-mono text-[0.66rem] text-faint">
              {server.transport === 'stdio'
                ? `${server.command ?? ''} ${(server.args ?? []).join(' ')}`.trim()
                : server.url}
            </div>

            {server.error && <p className="mt-1 text-[0.68rem] text-danger">{server.error}</p>}

            {server.tools.length > 0 && (
              <ul className="mt-2 space-y-1 border-t border-surface-border pt-2">
                {server.tools.map((tool) => (
                  <li key={tool.name} className="flex items-start gap-2">
                    <input
                      type="checkbox"
                      className="mt-0.5"
                      checked={tool.enabled}
                      onChange={(e) => void setToolEnabled(server.id, tool.name, e.target.checked)}
                      id={`tool-${server.id}-${tool.name}`}
                    />
                    <label
                      htmlFor={`tool-${server.id}-${tool.name}`}
                      className="min-w-0 cursor-pointer text-xs"
                    >
                      <span className="font-mono text-[0.7rem]">{tool.name}</span>
                      {tool.description && (
                        <span className="block text-[0.66rem] text-faint">{tool.description}</span>
                      )}
                    </label>
                  </li>
                ))}
              </ul>
            )}

            {server.status === 'ready' && server.tools.length === 0 && (
              <p className="mt-1 text-[0.68rem] text-faint">This server exposes no tools.</p>
            )}
          </li>
        ))}
      </ul>

      {adding ? (
        <AddMcpServerForm onDone={() => setAdding(false)} />
      ) : (
        <button className="btn btn-ghost mt-2 h-7 px-2 text-xs" onClick={() => setAdding(true)}>
          <PlusIcon size={12} /> Add server
        </button>
      )}
    </div>
  )
}

function StatusPill({ status }: { status: McpServerInfo['status'] }) {
  const styles: Record<McpServerInfo['status'], string> = {
    ready: 'bg-ok/15 text-ok',
    starting: 'bg-warn/15 text-warn',
    error: 'bg-danger/15 text-danger',
    stopped: 'bg-surface-muted text-muted',
  }
  return (
    <span className={`shrink-0 rounded px-1.5 py-0.5 text-[0.6rem] ${styles[status]}`}>
      {status}
    </span>
  )
}

function AddMcpServerForm({ onDone }: { onDone: () => void }) {
  const save = useMcpStore((s) => s.save)
  const [transport, setTransport] = useState<'stdio' | 'http'>('stdio')
  const [name, setName] = useState('')
  const [command, setCommand] = useState('')
  const [args, setArgs] = useState('')
  const [url, setUrl] = useState('')
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)

  const submit = async () => {
    setErr(null)
    const id = name
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '')
    if (!id) {
      setErr('Give the server a name.')
      return
    }
    setBusy(true)
    const ok = await save({
      id,
      name: name.trim(),
      transport,
      // Split on whitespace but keep quoted groups together — an arg like
      // --root "C:\My Documents" is one argument, not three.
      ...(transport === 'stdio'
        ? { command: command.trim(), args: splitArgs(args) }
        : { url: url.trim() }),
      enabled: true,
    })
    setBusy(false)
    if (ok) onDone()
    else setErr('Could not save. Check the command or URL.')
  }

  return (
    <div className="mt-2 rounded-lg border border-surface-border p-2.5">
      <div className="mb-2 flex gap-2">
        {(['stdio', 'http'] as const).map((t) => (
          <button
            key={t}
            className={`btn h-7 px-2 text-xs ${transport === t ? 'btn-primary' : 'btn-ghost'}`}
            onClick={() => setTransport(t)}
          >
            {t === 'stdio' ? 'Local command' : 'HTTP'}
          </button>
        ))}
      </div>
      <input
        className="field mb-1.5 h-8 text-xs"
        placeholder="Name (e.g. Filesystem)"
        value={name}
        onChange={(e) => setName(e.target.value)}
      />
      {transport === 'stdio' ? (
        <>
          <input
            className="field mb-1.5 h-8 font-mono text-xs"
            placeholder="Command (e.g. npx)"
            value={command}
            onChange={(e) => setCommand(e.target.value)}
          />
          <input
            className="field mb-1.5 h-8 font-mono text-xs"
            placeholder="Arguments (e.g. -y @modelcontextprotocol/server-filesystem C:\Docs)"
            value={args}
            onChange={(e) => setArgs(e.target.value)}
          />
        </>
      ) : (
        <input
          className="field mb-1.5 h-8 font-mono text-xs"
          placeholder="https://example.com/mcp"
          value={url}
          onChange={(e) => setUrl(e.target.value)}
        />
      )}
      {err && <p className="mb-1.5 text-xs text-danger">{err}</p>}
      <div className="flex justify-end gap-2">
        <button className="btn btn-ghost px-2 py-1 text-xs" onClick={onDone}>
          Cancel
        </button>
        <button
          className="btn btn-primary px-2 py-1 text-xs"
          onClick={() => void submit()}
          disabled={busy}
        >
          {busy ? 'Connecting…' : 'Add'}
        </button>
      </div>
    </div>
  )
}

/** Whitespace-separated, honouring double quotes around paths with spaces. */
function splitArgs(raw: string): string[] {
  const out: string[] = []
  const re = /"([^"]*)"|(\S+)/g
  let m: RegExpExecArray | null
  while ((m = re.exec(raw)) !== null) out.push(m[1] ?? m[2])
  return out
}

// ─── Skills ────────────────────────────────────────────────────────────────

function SkillsSection({ draft, setDraft }: { draft: Settings; setDraft: (s: Settings) => void }) {
  const skills = useSkillsStore((s) => s.skills)
  const userDir = useSkillsStore((s) => s.userDir)
  const load = useSkillsStore((s) => s.load)
  const remove = useSkillsStore((s) => s.remove)
  const save = useSkillsStore((s) => s.save)
  const [editing, setEditing] = useState<string | null>(null)

  useEffect(() => {
    void load()
  }, [load])

  const toggleSkill = (id: string) => {
    const next = draft.disabledSkills.includes(id)
      ? draft.disabledSkills.filter((x) => x !== id)
      : [...draft.disabledSkills, id]
    setDraft({ ...draft, disabledSkills: next })
  }

  return (
    <div>
      <p className="mb-2 text-xs text-faint">
        Drop markdown files in <code className="font-mono text-[0.66rem]">{userDir}</code> to add
        your own. Skills from opencode are picked up automatically.
      </p>

      <label className="mb-2 flex items-center gap-2 text-xs text-muted">
        <input
          type="checkbox"
          checked={draft.showImportedSkills}
          onChange={(e) => setDraft({ ...draft, showImportedSkills: e.target.checked })}
        />
        Show imported opencode skills
      </label>

      <ul className="space-y-1">
        {skills.map((s) => {
          const disabled = draft.disabledSkills.includes(s.id)
          return (
            <li
              key={s.id}
              className="flex items-center justify-between gap-2 rounded-lg border border-surface-border px-2 py-1.5"
            >
              <div className="flex min-w-0 items-center gap-2">
                {s.icon && <span aria-hidden>{s.icon}</span>}
                <label className="flex cursor-pointer items-center gap-2">
                  <input type="checkbox" checked={!disabled} onChange={() => toggleSkill(s.id)} />
                  <span className="truncate text-xs font-medium">{s.name}</span>
                </label>
                <span className="shrink-0 rounded bg-surface-muted px-1 py-0.5 text-[0.6rem] text-muted">
                  {s.source}
                </span>
              </div>
              {s.source === 'user' && (
                <button
                  className="btn btn-ghost h-6 px-1.5 text-[0.68rem] text-danger"
                  onClick={() => void remove(s.id)}
                  aria-label={`Delete ${s.name}`}
                >
                  <TrashIcon size={11} />
                </button>
              )}
            </li>
          )
        })}
      </ul>

      {editing === 'new' ? (
        <NewSkillForm
          onCancel={() => setEditing(null)}
          onSave={async (draft) => {
            const ok = await save(draft)
            if (ok) setEditing(null)
            return ok
          }}
        />
      ) : (
        <button className="btn btn-ghost mt-2 h-7 px-2 text-xs" onClick={() => setEditing('new')}>
          <PlusIcon size={12} /> New skill
        </button>
      )}
    </div>
  )
}

function NewSkillForm({
  onCancel,
  onSave,
}: {
  onCancel: () => void
  onSave: (draft: SkillDraft) => Promise<boolean>
}) {
  const [name, setName] = useState('')
  const [description, setDescription] = useState('')
  const [prompt, setPrompt] = useState('')
  const [hosts, setHosts] = useState<HostKind[]>([])
  const [busy, setBusy] = useState(false)

  const id = name
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
  const valid = id.length > 0 && prompt.trim().length > 0

  return (
    <div className="mt-2 rounded-lg border border-surface-border p-2.5">
      <input
        className="field mb-1.5 h-8 text-xs"
        placeholder="Name"
        value={name}
        onChange={(e) => setName(e.target.value)}
      />
      <input
        className="field mb-1.5 h-8 text-xs"
        placeholder="Short description"
        value={description}
        onChange={(e) => setDescription(e.target.value)}
      />
      <textarea
        className="field mb-1.5 min-h-[80px] resize-y text-xs"
        placeholder="Instructions given to the model when this skill runs…"
        value={prompt}
        onChange={(e) => setPrompt(e.target.value)}
      />
      <div className="mb-1.5 flex gap-3 text-xs text-muted">
        {(['word', 'excel', 'browser'] as const).map((h) => (
          <label key={h} className="flex items-center gap-1">
            <input
              type="checkbox"
              checked={hosts.includes(h)}
              onChange={(e) =>
                setHosts(e.target.checked ? [...hosts, h] : hosts.filter((x) => x !== h))
              }
            />
            {h === 'word' ? 'Word' : h === 'excel' ? 'Excel' : 'Browser'}
          </label>
        ))}
        <span className="text-faint">(neither = both)</span>
      </div>
      <div className="flex justify-end gap-2">
        <button className="btn btn-ghost px-2 py-1 text-xs" onClick={onCancel}>
          Cancel
        </button>
        <button
          className="btn btn-primary px-2 py-1 text-xs"
          disabled={!valid || busy}
          onClick={async () => {
            setBusy(true)
            await onSave({ id, name: name.trim(), description, hosts, prompt })
            setBusy(false)
          }}
        >
          Save
        </button>
      </div>
    </div>
  )
}

function ProviderList({ providers }: { providers: ProviderInfo[] }) {
  if (providers.length === 0) {
    return <p className="text-xs text-faint">No providers detected yet.</p>
  }
  return (
    <>
      <p className="mb-2 text-xs text-faint">
        Expand a provider to add its key or choose which of its models the picker offers. Visibility
        and favourites apply immediately — they are not part of Save.
      </p>
      <ul className="space-y-2">
        {providers.map((p) => (
          <ProviderRow key={p.id} provider={p} />
        ))}
      </ul>
    </>
  )
}

/** Show/hide and favourite toggles for one provider's models. */
function ProviderModelList({ providerId }: { providerId: string }) {
  const models = useModelsStore((s) => s.models)
  const hidden = useModelsStore((s) => s.hidden)
  const favorites = useModelsStore((s) => s.favorites)
  const setHidden = useModelsStore((s) => s.setHidden)
  const toggleFavorite = useModelsStore((s) => s.toggleFavorite)

  const mine = models.filter((m) => m.providerId === providerId)
  if (mine.length === 0) {
    return (
      <p className="mt-2 border-t border-surface-border pt-2 text-[0.68rem] text-faint">
        No models listed. Refresh models once this provider is reachable.
      </p>
    )
  }

  const hiddenSet = new Set(hidden)
  const shownCount = mine.filter((m) => !hiddenSet.has(m.id)).length
  const ids = mine.map((m) => m.id)

  return (
    <div className="mt-2 border-t border-surface-border pt-2">
      <div className="mb-1 flex items-center justify-between">
        <span className="text-[0.68rem] text-faint">
          {shownCount} of {mine.length} shown in the picker
        </span>
        <button
          className="btn btn-ghost h-6 px-1.5 text-[0.68rem]"
          onClick={() => setHidden(ids, shownCount > 0)}
        >
          {shownCount > 0 ? 'Hide all' : 'Show all'}
        </button>
      </div>
      <ul className="space-y-0.5">
        {mine.map((m) => {
          const isHidden = hiddenSet.has(m.id)
          const isFavorite = favorites.includes(m.id)
          return (
            <li key={m.id} className="flex items-center gap-2">
              <input
                type="checkbox"
                id={`model-${m.id}`}
                checked={!isHidden}
                onChange={(e) => setHidden([m.id], !e.target.checked)}
              />
              <label
                htmlFor={`model-${m.id}`}
                className={`min-w-0 flex-1 cursor-pointer truncate text-xs ${
                  isHidden ? 'text-faint line-through' : ''
                }`}
                title={m.id}
              >
                {m.name}
              </label>
              {m.contextWindow && (
                <span className="shrink-0 text-[0.62rem] text-faint">
                  {(m.contextWindow / 1000).toFixed(0)}k
                </span>
              )}
              <button
                className={`icon-btn h-5 w-5 shrink-0 ${isFavorite ? 'text-accent' : 'text-faint'}`}
                onClick={() => toggleFavorite(m.id)}
                aria-label={isFavorite ? `Unfavorite ${m.name}` : `Favorite ${m.name}`}
                title={isFavorite ? 'Remove from favorites' : 'Add to favorites'}
              >
                <StarIcon size={12} filled={isFavorite} />
              </button>
            </li>
          )
        })}
      </ul>
    </div>
  )
}

// ─── Opencode import ───────────────────────────────────────────────────────

function OpencodeImportSection({ setDraft }: { draft: Settings; setDraft: (s: Settings) => void }) {
  const [preview, setPreview] = useState<OpencodeImportResult | null>(null)
  // Starts true: the scan below fires on mount, so "busy" is the correct
  // initial state. Setting it inside the effect instead would be a synchronous
  // setState in an effect body — a cascading render for a value that is already
  // knowable at construction time.
  const [busy, setBusy] = useState(true)
  const [done, setDone] = useState<OpencodeImportResult | null>(null)

  // Scan on mount. Reading opencode's config is a cheap local file read, and
  // making the user press "Preview" first only served to hide the answer to the
  // question the section exists to answer: is there anything here to import?
  useEffect(() => {
    let cancelled = false
    previewOpencodeImport()
      .then((res) => {
        if (!cancelled) setPreview(res)
      })
      .catch((e: Error) => {
        if (!cancelled) setPreview(emptyImportResult(e.message))
      })
      .finally(() => {
        if (!cancelled) setBusy(false)
      })
    return () => {
      cancelled = true
    }
  }, [])

  const rescan = async () => {
    setBusy(true)
    setDone(null)
    try {
      setPreview(await previewOpencodeImport())
    } catch (e) {
      setPreview(emptyImportResult((e as Error).message))
    }
    setBusy(false)
  }

  const runImport = async () => {
    setBusy(true)
    try {
      const res = await runOpencodeImport()
      setDone(res)
      if (res.ok) {
        // Refresh the draft from the server so the newly imported
        // providers/agents/MCP appear in the panel that is still open.
        const fresh = await getSettings()
        setDraft(mergeDefaults(fresh))
        // Force a model list refresh so newly registered providers show up.
        void useModelsStore.getState().load(true)
      }
    } catch (e) {
      setDone(emptyImportResult((e as Error).message))
    }
    setBusy(false)
  }

  const result = done ?? preview
  const nothingFound =
    result != null &&
    result.providerCount === 0 &&
    result.mcpCount === 0 &&
    result.agentCount === 0 &&
    result.linkedProviders.length === 0

  return (
    <div className="text-xs">
      {busy && !result && <p className="text-faint">Looking for an opencode install…</p>}

      {result && (
        <>
          {result.sources.length > 0 ? (
            <p className="text-faint">
              Read <span className="font-mono text-[0.66rem]">{result.sources.join(', ')}</span>
            </p>
          ) : (
            <p className="text-faint">
              No opencode config found. Searched{' '}
              <span className="font-mono text-[0.66rem]">{result.searched.join(', ')}</span>
            </p>
          )}

          {!nothingFound && (
            <ul className="mt-2 space-y-1.5">
              {result.providers.map((p) => (
                <li key={p.id} className="rounded-lg border border-surface-border px-2 py-1.5">
                  <div className="flex items-center justify-between gap-2">
                    <span className="truncate font-medium">{p.name}</span>
                    <span className="shrink-0 text-[0.65rem] text-faint">
                      {p.modelCount > 0 ? `${p.modelCount} model hint(s)` : 'lists models live'}
                    </span>
                  </div>
                  <div className="truncate font-mono text-[0.64rem] text-faint">{p.baseUrl}</div>
                </li>
              ))}
              {result.linkedProviders.length > 0 && (
                <li className="rounded-lg border border-surface-border px-2 py-1.5">
                  <span className="font-medium">Keys for built-in providers</span>
                  <div className="truncate text-[0.65rem] text-faint">
                    {result.linkedProviders.join(', ')}
                  </div>
                </li>
              )}
              {result.mcpServers.length > 0 && (
                <li className="rounded-lg border border-surface-border px-2 py-1.5">
                  <span className="font-medium">
                    {result.mcpServers.length} MCP server
                    {result.mcpServers.length === 1 ? '' : 's'}
                  </span>
                  <div className="truncate text-[0.65rem] text-faint">
                    {result.mcpServers.map((s) => s.name).join(', ')} · imported disabled
                  </div>
                </li>
              )}
              {result.agents.length > 0 && (
                <li className="rounded-lg border border-surface-border px-2 py-1.5">
                  <span className="font-medium">
                    {result.agents.length} agent{result.agents.length === 1 ? '' : 's'}
                  </span>
                  <div className="truncate text-[0.65rem] text-faint">
                    {result.agents.map((a) => a.name).join(', ')} · enable below
                  </div>
                </li>
              )}
            </ul>
          )}

          {nothingFound && result.sources.length > 0 && (
            <p className="mt-1 text-warn">The config parsed but held nothing importable.</p>
          )}

          {result.errors.length > 0 && (
            <ul className="mt-2 space-y-0.5 text-[0.68rem] text-warn">
              {result.errors.map((err) => (
                <li key={err}>{err}</li>
              ))}
            </ul>
          )}

          {done?.ok && (
            <p className="mt-2 flex items-center gap-1 text-ok">
              <CheckIcon size={12} /> Imported. API keys are stored in the host&apos;s secret store.
            </p>
          )}
        </>
      )}

      <div className="mt-2 flex gap-2">
        <button
          className="btn btn-primary h-7 px-2 text-xs"
          onClick={() => void runImport()}
          disabled={busy || nothingFound}
        >
          {busy ? 'Working…' : done?.ok ? 'Import again' : 'Import'}
        </button>
        <button
          className="btn btn-ghost h-7 px-2 text-xs"
          onClick={() => void rescan()}
          disabled={busy}
        >
          <RefreshIcon size={12} /> Rescan
        </button>
      </div>
    </div>
  )
}

// ─── Agents ──────────────────────────────────────────────────────────────────

function AgentsSection({ draft, setDraft }: { draft: Settings; setDraft: (s: Settings) => void }) {
  const toggle = (id: string) => {
    const next = draft.agents.map((a) => (a.id === id ? { ...a, enabled: !a.enabled } : a))
    setDraft({ ...draft, agents: next })
  }

  if (draft.agents.length === 0) {
    return (
      <p className="text-xs text-faint">
        No agents imported yet. Use “Import from opencode” above.
      </p>
    )
  }

  return (
    <ul className="space-y-1">
      {draft.agents.map((a) => (
        <li
          key={a.id}
          className="flex items-center justify-between gap-2 rounded-lg border border-surface-border px-2 py-1.5"
          title={`${a.description || 'No description'}\nModel: ${a.model}\nSource: ${a.source}`}
        >
          <div className="flex min-w-0 items-center gap-2">
            <input type="checkbox" checked={a.enabled} onChange={() => toggle(a.id)} />
            <span className="truncate text-xs font-medium">{a.name}</span>
            <span className="shrink-0 rounded bg-surface-muted px-1 py-0.5 text-[0.6rem] text-muted">
              {a.source}
            </span>
          </div>
          <span className="truncate text-[0.65rem] text-faint">{a.model}</span>
        </li>
      ))}
    </ul>
  )
}

function ProviderRow({ provider }: { provider: ProviderInfo }) {
  const [expanded, setExpanded] = useState(false)
  const [keyInput, setKeyInput] = useState('')
  const [status, setStatus] = useState<{ kind: 'idle' | 'saving' | 'ok' | 'err'; msg?: string }>({
    kind: 'idle',
  })
  const [testResult, setTestResult] = useState<ProviderTestResponse | null>(null)
  const [testing, setTesting] = useState(false)

  const onSaveKey = async () => {
    if (!keyInput) return
    setStatus({ kind: 'saving' })
    try {
      await putProviderKey(provider.id, keyInput)
      setKeyInput('')
      setStatus({ kind: 'ok', msg: 'Saved' })
      setTimeout(() => setStatus({ kind: 'idle' }), 2000)
    } catch (e) {
      setStatus({ kind: 'err', msg: (e as Error).message })
    }
  }

  const onTest = async () => {
    setTesting(true)
    setTestResult(null)
    try {
      const r = await testProvider(provider.id)
      setTestResult(r)
    } catch (e) {
      setTestResult({ ok: false, reachable: false, error: (e as Error).message })
    }
    setTesting(false)
  }

  const onDeleteKey = async () => {
    try {
      await deleteProviderKey(provider.id)
      setStatus({ kind: 'ok', msg: 'Key removed' })
      setTimeout(() => setStatus({ kind: 'idle' }), 2000)
    } catch (e) {
      setStatus({ kind: 'err', msg: (e as Error).message })
    }
  }

  const needsKey = provider.kind === 'cloud' && !provider.configured
  return (
    <li className="rounded-lg border border-surface-border p-2.5">
      {/* Twelve always-expanded providers turned this section into most of the
          panel's scroll height, so the body is behind a disclosure and the row
          carries enough status to make expanding a deliberate choice. */}
      <button
        type="button"
        className="flex w-full items-center justify-between gap-2 text-left"
        onClick={() => setExpanded((v) => !v)}
        aria-expanded={expanded}
      >
        <span className="flex min-w-0 items-center gap-2">
          {expanded ? (
            <ChevronDownIcon size={13} className="shrink-0 text-faint" />
          ) : (
            <ChevronRightIcon size={13} className="shrink-0 text-faint" />
          )}
          <span className="truncate text-sm font-medium">{provider.name}</span>
          <span
            className={`shrink-0 rounded px-1.5 py-0.5 text-[0.6rem] ${
              provider.kind === 'local' ? 'bg-ok/15 text-ok' : 'bg-surface-muted text-muted'
            }`}
          >
            {provider.kind}
          </span>
          {/* "unreachable" only means something once there is a key to try.
              A cloud provider with no key always probes as unreachable, and
              flagging that in red made twelve providers the user has never
              configured look broken. */}
          {!provider.reachable && !needsKey && (
            <span
              className="shrink-0 text-[0.65rem] text-danger"
              title="Host cannot reach this provider"
            >
              unreachable
            </span>
          )}
        </span>
        <span className="flex shrink-0 items-center gap-2">
          <ProviderModelCount providerId={provider.id} />
          <span
            className={`text-[0.65rem] ${
              provider.configured ? 'text-ok' : needsKey ? 'text-warn' : 'text-faint'
            }`}
          >
            {provider.kind === 'local'
              ? provider.reachable
                ? 'ready'
                : 'off'
              : provider.configured
                ? 'key set'
                : 'no key'}
          </span>
        </span>
      </button>
      {!expanded ? null : (
        <>
          {provider.kind === 'cloud' && (
            <div className="mt-2 flex items-center gap-1.5">
              <input
                type="password"
                className="field h-8 flex-1 text-xs"
                placeholder={
                  provider.configured
                    ? '•••••••• (configured — enter new key to replace)'
                    : 'Paste API key'
                }
                value={keyInput}
                onChange={(e) => setKeyInput(e.target.value)}
                aria-label={`API key for ${provider.name}`}
                autoComplete="off"
              />
              <button
                className="btn btn-primary h-8 px-2 text-xs"
                onClick={onSaveKey}
                disabled={!keyInput || status.kind === 'saving'}
              >
                Save
              </button>
              {provider.configured && (
                <button
                  className="btn btn-danger h-8 px-2 text-xs"
                  onClick={onDeleteKey}
                  aria-label={`Remove key for ${provider.name}`}
                >
                  Remove
                </button>
              )}
            </div>
          )}
          <div className="mt-2 flex items-center gap-2">
            <button className="btn btn-ghost h-7 px-2 text-xs" onClick={onTest} disabled={testing}>
              {testing ? 'Testing…' : 'Test connection'}
            </button>
            {status.kind === 'ok' && (
              <span className="flex items-center gap-1 text-[0.7rem] text-ok">
                <CheckIcon size={12} /> {status.msg}
              </span>
            )}
            {status.kind === 'err' && (
              <span className="flex items-center gap-1 text-[0.7rem] text-danger">
                <AlertIcon size={12} /> {status.msg}
              </span>
            )}
            {testResult && (
              <span className={`text-[0.7rem] ${testResult.ok ? 'text-ok' : 'text-danger'}`}>
                {testResult.ok
                  ? `reachable · ${testResult.modelCount ?? 0} models`
                  : `failed${testResult.error ? ` · ${testResult.error}` : ''}`}
              </span>
            )}
          </div>
          <ProviderModelList providerId={provider.id} />
        </>
      )}
    </li>
  )
}

/** "3/19" on a collapsed provider row, so hiding models is visible without
 *  expanding every provider to find where they went. */
function ProviderModelCount({ providerId }: { providerId: string }) {
  const models = useModelsStore((s) => s.models)
  const hidden = useModelsStore((s) => s.hidden)
  const mine = models.filter((m) => m.providerId === providerId)
  if (mine.length === 0) return null
  const hiddenSet = new Set(hidden)
  const shown = mine.filter((m) => !hiddenSet.has(m.id)).length
  return (
    <span
      className={`tabular-nums text-[0.65rem] ${shown === mine.length ? 'text-faint' : 'text-warn'}`}
      title={`${shown} of ${mine.length} models shown in the picker`}
    >
      {shown}/{mine.length}
    </span>
  )
}
