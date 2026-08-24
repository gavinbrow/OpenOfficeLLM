// OpenOfficeLLM task pane entrypoint.
//
// The shell's whole job: detect Office, register the Office implementation of
// the UI's Shell contract, then mount. Everything after registerShell is the
// same code the Chrome extension runs.
import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import '@openofficellm/ui/index.css'
import { App, ErrorBoundary, registerShell, useChatStore, getDocumentKey } from '@openofficellm/ui'
import { bootstrap } from './office/bootstrap'
import { officeShell } from './office/shell'

void bootstrap().then(() => {
  registerShell(officeShell)

  // Before the first render, so the pane never flashes the previous document's
  // conversation. localStorage is shared across every document the add-in runs
  // in; the document key is what separates them.
  useChatStore.getState().adoptDocument(getDocumentKey())

  const rootEl = document.getElementById('root')
  if (!rootEl) throw new Error('root element missing')
  createRoot(rootEl).render(
    <StrictMode>
      <ErrorBoundary>
        <App />
      </ErrorBoundary>
    </StrictMode>,
  )
})
