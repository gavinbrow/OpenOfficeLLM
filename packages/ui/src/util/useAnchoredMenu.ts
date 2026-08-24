// Viewport-clamped positioning for the pane's pop-up menus.
//
// The task pane is ~320px wide and the model/mode/agent buttons sit in the
// middle of a control strip, so a plain `absolute right-0 w-56` menu is wider
// than the space to the left of its trigger and runs straight off the edge of
// the pane — in Word there is no page behind it to spill onto, so the menu is
// simply cut in half. The skill strip clips its own menu the same way, because
// `overflow-x-auto` makes the vertical axis scroll too.
//
// Both problems go away with `position: fixed` plus an explicit clamp: fixed
// escapes every clipping ancestor, and the clamp keeps the menu inside the pane
// wherever the trigger happens to be. The menu stays a DOM child of its
// container, so outside-click handling is unaffected.
//
// Placement is written straight to the node rather than returned as state: it
// is a measurement of the DOM being fed back into the DOM, and routing it
// through a render only adds a pass that has to complete before the menu can be
// shown.

import { useCallback, useEffect, useLayoutEffect, useRef } from 'react'
import type { RefObject } from 'react'

/** Clearance kept between the menu and the edge of the pane. */
const EDGE = 8
/** Gap between the trigger and the menu. */
const GAP = 4
/** Below this, the preferred side is too cramped to be worth using. */
const MIN_ROOM = 140

export interface AnchoredMenuOptions {
  /** Desired width in px. Narrowed when the pane cannot fit it. */
  width: number
  /** Side of the trigger to open on, space permitting. */
  side?: 'top' | 'bottom'
  /** Trigger edge the menu lines up with, space permitting. */
  align?: 'start' | 'end'
  /** Upper bound on height. Available room lowers it further. */
  maxHeight?: number
}

export interface AnchoredMenu<A extends HTMLElement> {
  /** Attach to the trigger. */
  anchorRef: RefObject<A>
  /** Attach to the menu, which must also carry the `menu-floating` class. */
  menuRef: RefObject<HTMLDivElement>
}

export function useAnchoredMenu<A extends HTMLElement>(
  open: boolean,
  { width, side = 'top', align = 'end', maxHeight }: AnchoredMenuOptions,
): AnchoredMenu<A> {
  const anchorRef = useRef<A>(null)
  const menuRef = useRef<HTMLDivElement>(null)

  const place = useCallback(() => {
    const anchor = anchorRef.current
    const menu = menuRef.current
    if (!anchor || !menu) return

    const r = anchor.getBoundingClientRect()
    const vw = window.innerWidth
    const vh = window.innerHeight

    const w = Math.min(width, Math.max(vw - EDGE * 2, 0))
    // `align` is only a preference. The clamp wins, and that is what keeps the
    // menu on screen when its trigger sits near an edge.
    const preferred = align === 'start' ? r.left : r.right - w
    const left = Math.min(Math.max(preferred, EDGE), Math.max(vw - EDGE - w, EDGE))

    const roomAbove = r.top - GAP - EDGE
    const roomBelow = vh - r.bottom - GAP - EDGE
    let above = side === 'top'
    // Flip only when the preferred side is genuinely cramped *and* the other
    // side is better, so the menu does not swap sides over a few pixels.
    if (above && roomAbove < MIN_ROOM && roomBelow > roomAbove) above = false
    else if (!above && roomBelow < MIN_ROOM && roomAbove > roomBelow) above = true

    const room = Math.max(above ? roomAbove : roomBelow, 0)

    menu.style.left = `${Math.round(left)}px`
    menu.style.width = `${Math.round(w)}px`
    menu.style.maxHeight = `${Math.round(maxHeight ? Math.min(maxHeight, room) : room)}px`
    if (above) {
      menu.style.bottom = `${Math.round(vh - r.top + GAP)}px`
      menu.style.top = 'auto'
    } else {
      menu.style.top = `${Math.round(r.bottom + GAP)}px`
      menu.style.bottom = 'auto'
    }
    // `.menu-floating` starts hidden so the menu is never painted at the
    // unplaced position, even if a browser skips the pre-paint layout effect.
    menu.style.visibility = 'visible'
  }, [width, side, align, maxHeight])

  // Layout effect, not effect: this runs once the menu is in the DOM but before
  // paint, so the first frame the user sees is already positioned.
  useLayoutEffect(() => {
    if (open) place()
  }, [open, place])

  useEffect(() => {
    if (!open) return
    const onChange = () => place()
    window.addEventListener('resize', onChange)
    // Capture phase: a trigger can live inside a scrolling list, and scroll
    // does not bubble.
    window.addEventListener('scroll', onChange, true)
    return () => {
      window.removeEventListener('resize', onChange)
      window.removeEventListener('scroll', onChange, true)
    }
  }, [open, place])

  return { anchorRef, menuRef }
}
