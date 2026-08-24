// The strip that appears when the assistant has nothing to act on.
//
// Both shells can end up here and they mean different things by it. In Office
// it is a development situation — the pane was opened in a plain browser, so
// Office.js never loaded. In the extension it is routine: the side panel is
// open but the tab in front of it is a PDF viewer, a `chrome://` page, or one
// the user has not granted access to. Saying "running outside Office" there
// would be nonsense, so the copy is chosen per surface.

import { getHost, getSurface } from '../host/bridge'

export function HostBanner() {
  const host = getHost()
  if (host !== 'none') return null
  const browser = getSurface() === 'browser'

  return (
    <div
      className="flex shrink-0 items-center gap-2 border-b border-warn/30 bg-warn/10 px-3 py-1.5 text-[0.7rem] text-warn"
      role="banner"
    >
      <svg
        width="14"
        height="14"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.75"
        strokeLinecap="round"
        strokeLinejoin="round"
        aria-hidden
      >
        <path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0Z" />
        <path d="M12 9v4" />
        <path d="M12 17h.01" />
      </svg>
      <span>
        {browser
          ? 'No readable page attached — this tab is restricted or access has not been granted. Chat still works; page tools are unavailable.'
          : 'Running outside Office — document tools are unavailable. Use this mode only for development.'}
      </span>
    </div>
  )
}
