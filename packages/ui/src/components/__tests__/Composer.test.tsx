import { describe, it, expect, beforeEach } from 'vitest'
import { render, screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { Composer } from '../Composer'
import { useChatStore } from '../../store/chatStore'
import { installFakeShell } from '../../test/fakeShell'

installFakeShell({ host: 'word' })

function resetChatStore() {
  useChatStore.setState({
    conversations: [],
    activeId: null,
    streaming: false,
    streamingRequestId: null,
    usage: { promptTokens: 0, completionTokens: 0, totalTokens: 0, estimatedCostUsd: 0 },
    lastError: null,
    reconnecting: null,
  })
}

describe('Composer', () => {
  beforeEach(() => {
    resetChatStore()
  })

  it('renders a textarea and send button', () => {
    render(<Composer />)
    expect(screen.getByRole('textbox', { name: 'Message composer' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Send message' })).toBeInTheDocument()
  })

  it('disables send when empty', () => {
    render(<Composer />)
    expect(screen.getByRole('button', { name: 'Send message' })).toBeDisabled()
  })

  it('sends on Enter and clears the input', async () => {
    const user = userEvent.setup()
    render(<Composer />)
    const ta = screen.getByRole('textbox', { name: 'Message composer' }) as HTMLTextAreaElement
    await user.type(ta, 'hello world')
    expect(ta.value).toBe('hello world')
    await user.keyboard('{Enter}')
    expect(ta.value).toBe('')
  })

  it('inserts a newline on Shift+Enter and does not send', async () => {
    const user = userEvent.setup()
    render(<Composer />)
    const ta = screen.getByRole('textbox', { name: 'Message composer' }) as HTMLTextAreaElement
    await user.type(ta, 'line one')
    await user.keyboard('{Shift>}{Enter}{/Shift}')
    await user.type(ta, 'line two')
    expect(ta.value).toBe('line one\nline two')
  })

  it('shows the Stop button while streaming', async () => {
    const user = userEvent.setup()
    render(<Composer />)
    const ta = screen.getByRole('textbox', { name: 'Message composer' })
    await user.type(ta, 'go')
    await user.keyboard('{Enter}')
    // send() will fail because no model is selected; streaming stays false.
    // Force streaming true to assert the stop button swap.
    useChatStore.setState({ streaming: true })
    // findByRole retries until the re-render lands.
    expect(await screen.findByRole('button', { name: 'Stop generating' })).toBeInTheDocument()
    expect(screen.queryByRole('button', { name: 'Send message' })).not.toBeInTheDocument()
  })

  it('shows a rough token estimate', async () => {
    const user = userEvent.setup()
    render(<Composer />)
    const ta = screen.getByRole('textbox', { name: 'Message composer' })
    await user.type(ta, 'abcdefgh') // 8 chars -> 2 tokens
    expect(screen.getByText(/≈ 2 tokens/)).toBeInTheDocument()
  })

  it('renders a paperclip button with an accessible Attach files label', () => {
    render(<Composer />)
    expect(screen.getByLabelText('Attach files')).toBeInTheDocument()
  })

  it('has a hidden file input behind the paperclip button', () => {
    render(<Composer />)
    const input = screen.getByLabelText('Attach files').querySelector('input[type="file"]')
    expect(input).not.toBeNull()
    expect(input).toHaveAttribute('type', 'file')
    // Multiple is set so the user can attach more than one file at a time.
    expect(input).toHaveAttribute('multiple')
  })
})
