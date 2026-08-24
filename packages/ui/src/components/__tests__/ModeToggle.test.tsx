import { describe, it, expect, beforeEach } from 'vitest'
import { render, screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { ModeToggle } from '../ModeToggle'
import { useSettingsStore } from '../../store/settingsStore'
import { DEFAULT_SETTINGS } from '@openofficellm/shared'
import { installFakeShell } from '../../test/fakeShell'

installFakeShell({ host: 'word' })

function resetSettingsStore() {
  useSettingsStore.setState({
    settings: { ...DEFAULT_SETTINGS },
    loading: false,
    saving: false,
    error: null,
  })
}

describe('ModeToggle', () => {
  beforeEach(() => {
    resetSettingsStore()
    localStorage.clear()
  })

  it('renders a compact mode button with Propose selected by default', () => {
    render(<ModeToggle />)
    expect(screen.getByRole('button', { name: 'Edit mode' })).toHaveTextContent('Propose')
  })

  it('opens a menu and switches to Direct without confirmation', async () => {
    const user = userEvent.setup()
    render(<ModeToggle />)
    await user.click(screen.getByRole('button', { name: 'Edit mode' }))
    await user.click(screen.getByRole('menuitem', { name: /Direct/i }))
    expect(useSettingsStore.getState().settings.defaultMode.word).toBe('direct')
    expect(screen.queryByRole('dialog')).not.toBeInTheDocument()
  })

  it('shows a one-time confirmation before enabling Agentic', async () => {
    const user = userEvent.setup()
    render(<ModeToggle />)
    await user.click(screen.getByRole('button', { name: 'Edit mode' }))
    await user.click(screen.getByRole('menuitem', { name: /Agentic/i }))
    expect(screen.getByRole('dialog')).toBeInTheDocument()
    expect(screen.getByRole('heading', { name: 'Enable Agentic mode?' })).toBeInTheDocument()
  })

  it('does not enable Agentic if the confirmation is cancelled', async () => {
    const user = userEvent.setup()
    render(<ModeToggle />)
    await user.click(screen.getByRole('button', { name: 'Edit mode' }))
    await user.click(screen.getByRole('menuitem', { name: /Agentic/i }))
    await user.click(screen.getByRole('button', { name: 'Cancel' }))
    expect(useSettingsStore.getState().settings.defaultMode.word).toBe('propose')
    expect(screen.queryByRole('dialog')).not.toBeInTheDocument()
  })

  it('enables Agentic after confirmation and skips the dialog next time', async () => {
    const user = userEvent.setup()
    const { unmount } = render(<ModeToggle />)
    await user.click(screen.getByRole('button', { name: 'Edit mode' }))
    await user.click(screen.getByRole('menuitem', { name: /Agentic/i }))
    await user.click(screen.getByRole('button', { name: 'Enable Agentic' }))
    expect(useSettingsStore.getState().settings.defaultMode.word).toBe('agentic')
    unmount()

    // Switch back to Propose, then to Agentic again — no dialog this time.
    useSettingsStore.setState({
      settings: {
        ...DEFAULT_SETTINGS,
        defaultMode: { ...DEFAULT_SETTINGS.defaultMode, word: 'propose' },
      },
    })
    render(<ModeToggle />)
    await user.click(screen.getByRole('button', { name: 'Edit mode' }))
    await user.click(screen.getByRole('menuitem', { name: /Agentic/i }))
    expect(useSettingsStore.getState().settings.defaultMode.word).toBe('agentic')
    expect(screen.queryByRole('dialog')).not.toBeInTheDocument()
  })
})
