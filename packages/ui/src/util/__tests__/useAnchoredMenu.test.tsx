import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { useState } from 'react'
import { render, screen, act } from '@testing-library/react'
import { useAnchoredMenu, type AnchoredMenuOptions } from '../useAnchoredMenu'

// A Word task pane is roughly this wide, which is the whole point: a 224px menu
// does not fit to the left of a trigger sitting in the middle of the strip.
const PANE_WIDTH = 320
const PANE_HEIGHT = 700

function setViewport(w: number, h: number) {
  Object.defineProperty(window, 'innerWidth', { value: w, configurable: true, writable: true })
  Object.defineProperty(window, 'innerHeight', { value: h, configurable: true, writable: true })
}

function Harness(opts: AnchoredMenuOptions) {
  const [open, setOpen] = useState(false)
  const { anchorRef, menuRef } = useAnchoredMenu<HTMLButtonElement>(open, opts)
  return (
    <div>
      <button ref={anchorRef} onClick={() => setOpen(true)}>
        Trigger
      </button>
      {open && (
        <div ref={menuRef} className="menu-floating" role="menu">
          menu
        </div>
      )}
    </div>
  )
}

/** Renders the harness, gives the trigger a box (jsdom does no layout), opens
 *  the menu and hands back the placed menu element. */
function openAt(rect: Partial<DOMRect>, opts: AnchoredMenuOptions): HTMLElement {
  render(<Harness {...opts} />)
  const trigger = screen.getByRole('button', { name: 'Trigger' })
  const full = { x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, ...rect }
  trigger.getBoundingClientRect = () => ({ ...full, toJSON: () => full }) as DOMRect
  act(() => trigger.click())
  return screen.getByRole('menu')
}

const px = (v: string) => parseFloat(v)

describe('useAnchoredMenu', () => {
  const realWidth = window.innerWidth
  const realHeight = window.innerHeight

  beforeEach(() => setViewport(PANE_WIDTH, PANE_HEIGHT))
  afterEach(() => setViewport(realWidth, realHeight))

  it('keeps a right-aligned menu inside the pane instead of running off the left edge', () => {
    // The mode button sits mid-strip. Right-aligning a 224px menu to its right
    // edge would put the menu's left at 190 - 224 = -34, i.e. off screen.
    const menu = openAt({ top: 600, bottom: 628, left: 130, right: 190 }, { width: 224 })
    expect(px(menu.style.left)).toBeGreaterThanOrEqual(0)
    expect(px(menu.style.left) + px(menu.style.width)).toBeLessThanOrEqual(PANE_WIDTH)
  })

  it('keeps a left-aligned menu from running off the right edge', () => {
    const menu = openAt(
      { top: 600, bottom: 628, left: 240, right: 300 },
      { width: 290, align: 'start' },
    )
    expect(px(menu.style.left) + px(menu.style.width)).toBeLessThanOrEqual(PANE_WIDTH)
  })

  it('narrows the menu when the pane is narrower than the requested width', () => {
    setViewport(200, PANE_HEIGHT)
    const menu = openAt({ top: 600, bottom: 628, left: 60, right: 120 }, { width: 290 })
    expect(px(menu.style.width)).toBeLessThanOrEqual(200)
    expect(px(menu.style.left)).toBeGreaterThanOrEqual(0)
  })

  it('opens above the trigger and caps its height to the room there', () => {
    const menu = openAt({ top: 600, bottom: 628, left: 130, right: 190 }, { width: 224 })
    expect(menu.style.top).toBe('auto')
    // Bottom edge sits just above the trigger: 700 - 600 + 4.
    expect(px(menu.style.bottom)).toBe(104)
    expect(px(menu.style.maxHeight)).toBeLessThanOrEqual(600)
  })

  it('flips below when the preferred side is too cramped', () => {
    // Trigger near the top of the pane: 8px of usable room above, plenty below.
    const menu = openAt({ top: 20, bottom: 48, left: 130, right: 190 }, { width: 224 })
    expect(menu.style.bottom).toBe('auto')
    expect(px(menu.style.top)).toBe(52)
  })

  it('honours an explicit maxHeight', () => {
    const menu = openAt(
      { top: 600, bottom: 628, left: 130, right: 190 },
      { width: 224, maxHeight: 256 },
    )
    expect(px(menu.style.maxHeight)).toBe(256)
  })

  it('is not visible until it has been placed', () => {
    // `.menu-floating` hides it; the hook reveals it once measured, so the
    // inline value is what proves placement ran before paint.
    expect(
      openAt({ top: 600, bottom: 628, left: 130, right: 190 }, { width: 224 }).style.visibility,
    ).toBe('visible')
  })

  it('repositions when the pane is resized', () => {
    const menu = openAt(
      { top: 600, bottom: 628, left: 240, right: 300 },
      { width: 290, align: 'start' },
    )
    const before = px(menu.style.width)
    act(() => {
      setViewport(180, PANE_HEIGHT)
      window.dispatchEvent(new Event('resize'))
    })
    expect(px(menu.style.width)).toBeLessThan(before)
    expect(px(menu.style.left) + px(menu.style.width)).toBeLessThanOrEqual(180)
  })
})
