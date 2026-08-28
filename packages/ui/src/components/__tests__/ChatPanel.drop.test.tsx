// Drag-and-drop file attachment on the ChatPanel.
//
// ChatPanel is the drop target: a dragenter with `dataTransfer.types`
// containing 'Files' shows an overlay; dragleave hides it; drop hands the
// files to handleFileUpload (mocked here so no upload happens). The overlay
// is local component state, so these tests drive it through real drag events
// rather than poking the store. jsdom's DataTransfer is a stub, so each event
// carries a hand-built dataTransfer with the `types` and `files` the handlers
// read. The chatStore is left with no active conversation so the panel renders
// its EmptyState branch, which carries the same drop handlers as the populated
// view.

import { describe, it, expect, beforeEach, vi } from 'vitest'
import { render, screen, fireEvent } from '@testing-library/react'
import { ChatPanel } from '../ChatPanel'
import { useChatStore } from '../../store/chatStore'
import { useSettingsStore } from '../../store/settingsStore'
import { installFakeShell } from '../../test/fakeShell'

installFakeShell({ host: 'word' })

// handleFileUpload uploads to the host; mock the util so drop never hits the
// network. Only handleFileUpload is used by ChatPanel.
vi.mock('../../util/attachments', () => ({
  handleFileUpload: vi.fn(),
}))

import { handleFileUpload } from '../../util/attachments'

function resetChatStore() {
  useChatStore.setState({
    conversations: [],
    activeId: null,
    streaming: false,
    streamingRequestId: null,
    stepLimit: null,
    lastError: null,
    reconnecting: null,
  })
}

/** A minimal DataTransfer stand-in: the ChatPanel handlers read only
 *  `types` (an array-like of strings) and `files` (an array-like of File). */
function fakeDataTransfer(types: string[] = [], files: File[] = []): unknown {
  return {
    types,
    files,
    dropEffect: 'none',
    effectAllowed: 'none',
    setData: () => {},
    getData: () => '',
    clearData: () => {},
    setDragImage: () => {},
  }
}

describe('ChatPanel drop zone', () => {
  beforeEach(() => {
    resetChatStore()
    useSettingsStore.setState({ settings: useSettingsStore.getState().settings, loading: false, saving: false, error: null })
    vi.clearAllMocks()
  })

  it('shows the drop overlay when a file drag enters', () => {
    render(<ChatPanel />)
    // No overlay before the drag.
    expect(screen.queryByLabelText('Drop files to attach')).not.toBeInTheDocument()

    const container = document.querySelector('.relative') as HTMLElement
    fireEvent.dragEnter(container, { dataTransfer: fakeDataTransfer(['Files'], []) })

    expect(screen.getByLabelText('Drop files to attach')).toBeInTheDocument()
  })

  it('ignores a drag that does not carry Files', () => {
    render(<ChatPanel />)
    const container = document.querySelector('.relative') as HTMLElement
    // A text drag (types: ['text/plain']) must not trigger the overlay.
    fireEvent.dragEnter(container, { dataTransfer: fakeDataTransfer(['text/plain'], []) })
    expect(screen.queryByLabelText('Drop files to attach')).not.toBeInTheDocument()
  })

  it('hides the overlay when the drag leaves', () => {
    render(<ChatPanel />)
    const container = document.querySelector('.relative') as HTMLElement
    fireEvent.dragEnter(container, { dataTransfer: fakeDataTransfer(['Files'], []) })
    expect(screen.getByLabelText('Drop files to attach')).toBeInTheDocument()

    fireEvent.dragLeave(container, { dataTransfer: fakeDataTransfer(['Files'], []) })
    expect(screen.queryByLabelText('Drop files to attach')).not.toBeInTheDocument()
  })

  it('uploads each dropped file via handleFileUpload', () => {
    render(<ChatPanel />)
    const container = document.querySelector('.relative') as HTMLElement
    const fileA = new File(['a'], 'a.txt', { type: 'text/plain' })
    const fileB = new File(['b'], 'b.txt', { type: 'text/plain' })

    fireEvent.drop(container, { dataTransfer: fakeDataTransfer(['Files'], [fileA, fileB]) })

    expect(handleFileUpload).toHaveBeenCalledTimes(2)
    expect((handleFileUpload as unknown as { mock: { calls: File[][] } }).mock.calls.map((c) => c[0])).toEqual([
      fileA,
      fileB,
    ])
    // Drop clears the overlay.
    expect(screen.queryByLabelText('Drop files to attach')).not.toBeInTheDocument()
  })

  it('does not upload when the drop carries no files', () => {
    render(<ChatPanel />)
    const container = document.querySelector('.relative') as HTMLElement
    fireEvent.drop(container, { dataTransfer: fakeDataTransfer(['Files'], []) })
    expect(handleFileUpload).not.toHaveBeenCalled()
  })
})