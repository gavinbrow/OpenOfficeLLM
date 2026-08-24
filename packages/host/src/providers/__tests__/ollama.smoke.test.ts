import { describe, it, expect } from 'vitest'
import { OllamaAdapter } from '../ollama.js'

const SMOKE = process.env.OLLAMA_SMOKE === '1'
const OLLAMA_URL = process.env.OLLAMA_URL ?? 'http://127.0.0.1:11434'

describe.skipIf(!SMOKE)('ollama smoke (live)', () => {
  it('lists models from a running ollama', async () => {
    const adapter = new OllamaAdapter({ baseUrl: OLLAMA_URL })
    const reachable = await adapter.isReachable()
    expect(reachable).toBe(true)
    const models = await adapter.listModels()
    expect(models.length).toBeGreaterThan(0)
    expect(models[0].id.startsWith('ollama/')).toBe(true)
  })

  it('streams a chat response', async () => {
    const adapter = new OllamaAdapter({ baseUrl: OLLAMA_URL })
    const models = await adapter.listModels()
    if (models.length === 0) {
      console.warn('no models available for smoke test')
      return
    }
    const modelId = models[0].id
    const ctrl = new AbortController()
    const out: string[] = []
    for await (const ev of adapter.stream(
      {
        messages: [{ role: 'user', content: 'Say the word "ok" and nothing else.' }],
        model: modelId,
        mode: 'propose',
        maxTokens: 10,
      },
      ctrl.signal,
    )) {
      if (ev.type === 'delta') out.push(ev.text)
      if (ev.type === 'error') throw new Error(`stream error: ${ev.message}`)
    }
    expect(out.join('').length).toBeGreaterThan(0)
  })

  it('aborts mid-stream', async () => {
    const adapter = new OllamaAdapter({ baseUrl: OLLAMA_URL })
    const models = await adapter.listModels()
    if (models.length === 0) return
    const modelId = models[0].id
    const ctrl = new AbortController()
    let count = 0
    for await (const ev of adapter.stream(
      {
        messages: [{ role: 'user', content: 'Count from 1 to 100 slowly.' }],
        model: modelId,
        mode: 'propose',
      },
      ctrl.signal,
    )) {
      count += 1
      if (count === 2) ctrl.abort()
      if (ev.type === 'error') break
    }
    expect(count).toBeLessThan(50)
  })
})
