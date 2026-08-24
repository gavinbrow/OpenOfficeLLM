import { describe, it, expect, vi, beforeEach } from 'vitest'
import { useModelsStore } from '../../store/modelsStore'
import { getModels, getProviders } from '../../api/client'

vi.mock('../../api/client', () => ({
  getModels: vi.fn(),
  getProviders: vi.fn(),
}))

const MODELS = [
  {
    id: 'ollama/llama3',
    name: 'llama3',
    providerId: 'ollama',
    providerName: 'Ollama',
    kind: 'local' as const,
    capabilities: { tools: true, vision: false, streaming: true },
  },
  {
    id: 'anthropic/claude',
    name: 'claude',
    providerId: 'anthropic',
    providerName: 'Anthropic',
    kind: 'cloud' as const,
    capabilities: { tools: true, vision: true, streaming: true },
  },
]

const BOTH_REACHABLE = [
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

function resetStore() {
  useModelsStore.setState({
    models: [],
    providers: [],
    selectedModelId: null,
    hidden: [],
    favorites: [],
    loading: false,
    error: null,
    lastFetch: 0,
  })
}

describe('modelsStore filtering', () => {
  beforeEach(() => {
    resetStore()
    localStorage.clear()
    vi.clearAllMocks()
  })

  it('shows local models when reachable', async () => {
    vi.mocked(getModels).mockResolvedValue({ models: MODELS })
    vi.mocked(getProviders).mockResolvedValue({
      providers: [
        {
          id: 'ollama',
          name: 'Ollama',
          kind: 'local',
          reachable: true,
          configured: true,
          capabilities: { tools: true, vision: false, streaming: true },
        },
        {
          id: 'anthropic',
          name: 'Anthropic',
          kind: 'cloud',
          reachable: true,
          configured: false,
          capabilities: { tools: true, vision: true, streaming: true },
        },
      ],
    })
    await useModelsStore.getState().load()
    expect(useModelsStore.getState().models).toHaveLength(1)
    expect(useModelsStore.getState().models[0].id).toBe('ollama/llama3')
  })

  it('shows cloud models only when configured and reachable', async () => {
    vi.mocked(getModels).mockResolvedValue({ models: MODELS })
    vi.mocked(getProviders).mockResolvedValue({
      providers: [
        {
          id: 'ollama',
          name: 'Ollama',
          kind: 'local',
          reachable: true,
          configured: true,
          capabilities: { tools: true, vision: false, streaming: true },
        },
        {
          id: 'anthropic',
          name: 'Anthropic',
          kind: 'cloud',
          reachable: true,
          configured: true,
          capabilities: { tools: true, vision: true, streaming: true },
        },
      ],
    })
    await useModelsStore.getState().load()
    expect(useModelsStore.getState().models).toHaveLength(2)
  })

  it('surfaces a provider-list failure instead of reporting no models', async () => {
    // Every model is filtered out when provider info is missing, so a swallowed
    // provider error is indistinguishable from an empty catalogue.
    vi.mocked(getModels).mockResolvedValue({ models: MODELS })
    vi.mocked(getProviders).mockRejectedValue(new Error('forbidden'))
    await useModelsStore.getState().load()
    expect(useModelsStore.getState().error).toBe('forbidden')
    expect(useModelsStore.getState().loading).toBe(false)
  })

  it('clears selection when the selected model becomes unavailable', async () => {
    useModelsStore.setState({ selectedModelId: 'anthropic/claude' })
    vi.mocked(getModels).mockResolvedValue({ models: MODELS })
    vi.mocked(getProviders).mockResolvedValue({
      providers: [
        {
          id: 'ollama',
          name: 'Ollama',
          kind: 'local',
          reachable: true,
          configured: true,
          capabilities: { tools: true, vision: false, streaming: true },
        },
        {
          id: 'anthropic',
          name: 'Anthropic',
          kind: 'cloud',
          reachable: true,
          configured: false,
          capabilities: { tools: true, vision: true, streaming: true },
        },
      ],
    })
    await useModelsStore.getState().load()
    expect(useModelsStore.getState().selectedModelId).toBe('ollama/llama3')
  })
})

describe('modelsStore visibility and favourites', () => {
  beforeEach(() => {
    resetStore()
    localStorage.clear()
    vi.clearAllMocks()
    useModelsStore.setState({ models: MODELS, providers: BOTH_REACHABLE })
  })

  it('drops hidden models from the picker list but keeps them in the catalogue', () => {
    useModelsStore.getState().setHidden(['anthropic/claude'], true)
    expect(
      useModelsStore
        .getState()
        .visibleModels()
        .map((m) => m.id),
    ).toEqual(['ollama/llama3'])
    expect(useModelsStore.getState().models).toHaveLength(2)
  })

  it('hides and shows a whole provider in one call', () => {
    const ids = MODELS.map((m) => m.id)
    useModelsStore.getState().setHidden(ids, true)
    expect(useModelsStore.getState().visibleModels()).toHaveLength(0)
    useModelsStore.getState().setHidden(ids, false)
    expect(useModelsStore.getState().visibleModels()).toHaveLength(2)
  })

  it('does not record a model as hidden twice', () => {
    useModelsStore.getState().setHidden(['ollama/llama3'], true)
    useModelsStore.getState().setHidden(['ollama/llama3'], true)
    expect(useModelsStore.getState().hidden).toEqual(['ollama/llama3'])
  })

  it('moves the selection off a model the user just hid', () => {
    useModelsStore.getState().selectModel('ollama/llama3')
    useModelsStore.getState().setHidden(['ollama/llama3'], true)
    expect(useModelsStore.getState().selectedModelId).toBe('anthropic/claude')
  })

  it('leaves the selection alone when an unrelated model is hidden', () => {
    useModelsStore.getState().selectModel('ollama/llama3')
    useModelsStore.getState().setHidden(['anthropic/claude'], true)
    expect(useModelsStore.getState().selectedModelId).toBe('ollama/llama3')
  })

  it('clears the selection when every model is hidden', () => {
    useModelsStore.getState().selectModel('ollama/llama3')
    useModelsStore.getState().setHidden(
      MODELS.map((m) => m.id),
      true,
    )
    expect(useModelsStore.getState().selectedModelId).toBeNull()
  })

  it('prefers a favourite when it has to pick a replacement', () => {
    useModelsStore.getState().toggleFavorite('anthropic/claude')
    useModelsStore.getState().selectModel('ollama/llama3')
    useModelsStore.getState().setHidden(['ollama/llama3'], true)
    expect(useModelsStore.getState().selectedModelId).toBe('anthropic/claude')
  })

  it('toggles a favourite on and off and persists it', () => {
    useModelsStore.getState().toggleFavorite('ollama/llama3')
    expect(useModelsStore.getState().favorites).toEqual(['ollama/llama3'])
    expect(localStorage.getItem('openofficellm:models.favorites')).toBe('["ollama/llama3"]')
    useModelsStore.getState().toggleFavorite('ollama/llama3')
    expect(useModelsStore.getState().favorites).toEqual([])
  })

  it('never selects a hidden model on refresh', async () => {
    useModelsStore.getState().setHidden(['ollama/llama3'], true)
    useModelsStore.setState({ selectedModelId: null })
    vi.mocked(getModels).mockResolvedValue({ models: MODELS })
    vi.mocked(getProviders).mockResolvedValue({ providers: BOTH_REACHABLE })
    await useModelsStore.getState().load()
    expect(useModelsStore.getState().selectedModelId).toBe('anthropic/claude')
  })
})
