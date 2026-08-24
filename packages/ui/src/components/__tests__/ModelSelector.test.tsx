import { describe, it, expect, beforeEach } from 'vitest'
import { render, screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { ModelSelector } from '../ModelSelector'
import { useModelsStore } from '../../store/modelsStore'
import { installFakeShell } from '../../test/fakeShell'

installFakeShell({ host: 'word' })

const MODELS = [
  {
    id: 'ollama/glm-5.2:cloud',
    name: 'glm-5.2:cloud',
    providerId: 'ollama',
    providerName: 'Ollama',
    kind: 'local' as const,
    capabilities: { tools: true, vision: false, streaming: true },
    contextWindow: 8192,
  },
  {
    id: 'anthropic/claude-3-5-sonnet',
    name: 'claude-3-5-sonnet',
    providerId: 'anthropic',
    providerName: 'Anthropic',
    kind: 'cloud' as const,
    capabilities: { tools: true, vision: true, streaming: true },
  },
]

const PROVIDERS = [
  {
    id: 'ollama',
    name: 'Ollama',
    kind: 'local' as const,
    reachable: true,
    configured: true,
    capabilities: { tools: true, vision: false, streaming: true },
  },
  {
    id: 'anthropic',
    name: 'Anthropic',
    kind: 'cloud' as const,
    reachable: true,
    configured: true,
    capabilities: { tools: true, vision: true, streaming: true },
  },
]

function resetModelsStore() {
  useModelsStore.setState({
    models: MODELS,
    providers: PROVIDERS,
    selectedModelId: 'ollama/glm-5.2:cloud',
    hidden: [],
    favorites: [],
    loading: false,
    error: null,
    lastFetch: Date.now(),
  })
}

describe('ModelSelector', () => {
  beforeEach(() => {
    resetModelsStore()
    localStorage.clear()
  })

  it('shows the selected model name', () => {
    render(<ModelSelector />)
    expect(screen.getByText('glm-5.2:cloud')).toBeInTheDocument()
  })

  it('opens the dropdown and lists models grouped by provider', async () => {
    const user = userEvent.setup()
    render(<ModelSelector />)
    await user.click(screen.getByRole('button', { name: 'Select model' }))
    const listbox = screen.getByRole('listbox', { name: 'Models' })
    expect(listbox).toBeInTheDocument()
    // Provider names appear as section headers inside the listbox.
    expect(listbox).toHaveTextContent('Ollama')
    // Anthropic is configured in the mock providers, so its model appears.
    expect(listbox).toHaveTextContent('Anthropic')
  })

  it('filters models by search query', async () => {
    const user = userEvent.setup()
    render(<ModelSelector />)
    await user.click(screen.getByRole('button', { name: 'Select model' }))
    await user.type(screen.getByRole('textbox', { name: 'Search models' }), 'glm')
    const listbox = screen.getByRole('listbox', { name: 'Models' })
    expect(listbox).toHaveTextContent(/glm-5.2/i)
  })

  it('selects a model on click', async () => {
    const user = userEvent.setup()
    render(<ModelSelector />)
    await user.click(screen.getByRole('button', { name: 'Select model' }))
    const option = screen.getByRole('option', { name: /glm-5.2:cloud/i })
    await user.click(option)
    expect(useModelsStore.getState().selectedModelId).toBe('ollama/glm-5.2:cloud')
  })

  it('badges local models with "Local · no cost"', async () => {
    const user = userEvent.setup()
    render(<ModelSelector />)
    await user.click(screen.getByRole('button', { name: 'Select model' }))
    expect(screen.getByText('Local · no cost')).toBeInTheDocument()
  })

  it('collapses providers other than the one holding the selected model', async () => {
    const user = userEvent.setup()
    render(<ModelSelector />)
    await user.click(screen.getByRole('button', { name: 'Select model' }))
    // Ollama holds the selection, so its model is listed; Anthropic's is not
    // until its group is opened.
    expect(screen.getByRole('option', { name: /glm-5.2:cloud/i })).toBeInTheDocument()
    expect(screen.queryByRole('option', { name: /claude-3-5-sonnet/i })).not.toBeInTheDocument()
    expect(screen.getByRole('button', { name: /Anthropic/ })).toHaveAttribute(
      'aria-expanded',
      'false',
    )
  })

  it('expands a provider group when its header is clicked', async () => {
    const user = userEvent.setup()
    render(<ModelSelector />)
    await user.click(screen.getByRole('button', { name: 'Select model' }))
    await user.click(screen.getByRole('button', { name: /Anthropic/ }))
    expect(screen.getByRole('option', { name: /claude-3-5-sonnet/i })).toBeInTheDocument()
  })

  it('collapses the group holding the selected model when asked to', async () => {
    const user = userEvent.setup()
    render(<ModelSelector />)
    await user.click(screen.getByRole('button', { name: 'Select model' }))
    await user.click(screen.getByRole('button', { name: /Ollama/ }))
    expect(screen.queryByRole('option', { name: /glm-5.2:cloud/i })).not.toBeInTheDocument()
  })

  it('opens every matching group while a search is active', async () => {
    const user = userEvent.setup()
    render(<ModelSelector />)
    await user.click(screen.getByRole('button', { name: 'Select model' }))
    await user.type(screen.getByRole('textbox', { name: 'Search models' }), 'claude')
    // Anthropic is collapsed by default, but a match has to be reachable.
    expect(screen.getByRole('option', { name: /claude-3-5-sonnet/i })).toBeInTheDocument()
  })

  it('omits hidden models entirely', async () => {
    useModelsStore.setState({ hidden: ['anthropic/claude-3-5-sonnet'] })
    const user = userEvent.setup()
    render(<ModelSelector />)
    await user.click(screen.getByRole('button', { name: 'Select model' }))
    expect(screen.queryByRole('button', { name: /Anthropic/ })).not.toBeInTheDocument()
  })

  it('lists favourites in their own section above the providers', async () => {
    useModelsStore.setState({ favorites: ['anthropic/claude-3-5-sonnet'] })
    const user = userEvent.setup()
    render(<ModelSelector />)
    await user.click(screen.getByRole('button', { name: 'Select model' }))
    const favorites = screen.getByRole('group', { name: 'Favorites' })
    expect(favorites).toHaveTextContent('claude-3-5-sonnet')
  })

  it('favourites a model from the star button', async () => {
    const user = userEvent.setup()
    render(<ModelSelector />)
    await user.click(screen.getByRole('button', { name: 'Select model' }))
    await user.click(screen.getByRole('button', { name: 'Favorite glm-5.2:cloud' }))
    expect(useModelsStore.getState().favorites).toEqual(['ollama/glm-5.2:cloud'])
  })
})
