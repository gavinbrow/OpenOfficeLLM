import { useState } from 'react'
import { useChatStore } from '../store/chatStore'
import { PlusIcon, TrashIcon, EditIcon, CloseIcon, MenuIcon } from './icons'
import { useUiStore } from '../store/uiStore'

export function Sidebar() {
  const conversations = useChatStore((s) => s.conversations)
  const activeId = useChatStore((s) => s.activeId)
  const select = useChatStore((s) => s.selectChat)
  const newChat = useChatStore((s) => s.newChat)
  const rename = useChatStore((s) => s.renameChat)
  const del = useChatStore((s) => s.deleteChat)
  const setSidebarOpen = useUiStore((s) => s.setSidebarOpen)
  const [renamingId, setRenamingId] = useState<string | null>(null)
  const [draft, setDraft] = useState('')

  return (
    <div className="flex h-full w-full flex-col bg-surface-muted">
      <div className="flex items-center justify-between border-b border-surface-border p-2">
        <span className="text-xs font-semibold">History</span>
        <div className="flex items-center gap-1">
          <button
            className="btn btn-ghost h-7 w-7"
            onClick={newChat}
            aria-label="New chat"
            title="New chat"
          >
            <PlusIcon size={14} />
          </button>
          <button
            className="btn btn-ghost h-7 w-7"
            onClick={() => setSidebarOpen(false)}
            aria-label="Close sidebar"
            title="Close"
          >
            <CloseIcon size={14} />
          </button>
        </div>
      </div>
      <div className="flex-1 overflow-y-auto p-1.5">
        {conversations.length === 0 ? (
          <div className="p-4 text-center text-xs text-faint">No conversations yet.</div>
        ) : (
          <ul className="space-y-0.5">
            {conversations.map((c) => (
              <li
                key={c.id}
                className={`group flex items-center gap-1 rounded px-1.5 py-1.5 text-xs ${
                  c.id === activeId ? 'bg-accent-soft text-accent-fg' : 'hover:bg-surface-hover'
                }`}
              >
                {renamingId === c.id ? (
                  <input
                    className="field h-6 flex-1 text-xs"
                    value={draft}
                    autoFocus
                    onChange={(e) => setDraft(e.target.value)}
                    onBlur={() => {
                      rename(c.id, draft)
                      setRenamingId(null)
                    }}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        rename(c.id, draft)
                        setRenamingId(null)
                      }
                      if (e.key === 'Escape') setRenamingId(null)
                    }}
                    aria-label="Rename conversation"
                  />
                ) : (
                  <button
                    className="min-w-0 flex-1 truncate text-left"
                    onClick={() => select(c.id)}
                    title={c.title}
                  >
                    {c.title}
                  </button>
                )}
                <div className="flex shrink-0 items-center opacity-0 group-hover:opacity-100 focus-within:opacity-100">
                  <button
                    className="icon-btn h-5 w-5"
                    onClick={() => {
                      setRenamingId(c.id)
                      setDraft(c.title)
                    }}
                    aria-label={`Rename ${c.title}`}
                    title="Rename"
                  >
                    <EditIcon size={11} />
                  </button>
                  <button
                    className="icon-btn h-5 w-5 text-danger"
                    onClick={() => del(c.id)}
                    aria-label={`Delete ${c.title}`}
                    title="Delete"
                  >
                    <TrashIcon size={11} />
                  </button>
                </div>
              </li>
            ))}
          </ul>
        )}
      </div>
    </div>
  )
}

export function SidebarToggle() {
  const toggle = useUiStore((s) => s.toggleSidebar)
  return (
    <button
      className="btn btn-ghost h-8 w-8"
      onClick={toggle}
      aria-label="Toggle conversation history"
      title="History"
    >
      <MenuIcon size={16} />
    </button>
  )
}
