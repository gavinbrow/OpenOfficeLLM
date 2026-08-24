import type { ModelInfo } from '@openofficellm/shared'

interface PriceEntry {
  inputPer1k: number
  outputPer1k: number
}

const PRICE_TABLE: Record<string, PriceEntry> = {
  // Anthropic list prices per 1k tokens (published rates are per 1M).
  'anthropic/claude-opus-5': { inputPer1k: 0.005, outputPer1k: 0.025 },
  'anthropic/claude-opus-4-8': { inputPer1k: 0.005, outputPer1k: 0.025 },
  'anthropic/claude-sonnet-5': { inputPer1k: 0.003, outputPer1k: 0.015 },
  'anthropic/claude-sonnet-4-6': { inputPer1k: 0.003, outputPer1k: 0.015 },
  'anthropic/claude-haiku-4-5': { inputPer1k: 0.001, outputPer1k: 0.005 },
  'openai/gpt-4o': { inputPer1k: 0.0025, outputPer1k: 0.01 },
  'openai/gpt-4o-mini': { inputPer1k: 0.00015, outputPer1k: 0.0006 },
  'openai/gpt-4-turbo': { inputPer1k: 0.01, outputPer1k: 0.03 },
  'google/gemini-1.5-pro': { inputPer1k: 0.00125, outputPer1k: 0.005 },
  'google/gemini-1.5-flash': { inputPer1k: 0.000075, outputPer1k: 0.0003 },
  'groq/llama-3.3-70b': { inputPer1k: 0.00059, outputPer1k: 0.00079 },
  'deepseek/deepseek-chat': { inputPer1k: 0.00014, outputPer1k: 0.00028 },
  'together/meta-llama-3.1-70b': { inputPer1k: 0.00088, outputPer1k: 0.0011 },
  'openrouter/auto': { inputPer1k: 0.001, outputPer1k: 0.002 },
  'xai/grok-2': { inputPer1k: 0.002, outputPer1k: 0.01 },
}

export function estimateCost(
  modelId: string,
  promptTokens: number,
  completionTokens: number,
): number {
  const entry = lookupPrice(modelId)
  if (!entry) return 0
  return (promptTokens / 1000) * entry.inputPer1k + (completionTokens / 1000) * entry.outputPer1k
}

export function applyPricing(model: ModelInfo): ModelInfo {
  const entry = lookupPrice(model.id)
  if (!entry) return model
  return {
    ...model,
    inputPricePer1k: entry.inputPer1k,
    outputPricePer1k: entry.outputPer1k,
  }
}

/** Look up a price entry by model id. Tries an exact match first, then a
 *  prefix match so dated model variants (e.g. `anthropic/claude-3-5-sonnet-20241022`)
 *  match their undated base entry (`anthropic/claude-3-5-sonnet`). Picks the
 *  longest matching prefix so `anthropic/claude-3-5-sonnet-v2` would match
 *  `anthropic/claude-3-5-sonnet` over `anthropic/claude-3` if both existed. */
function lookupPrice(modelId: string): PriceEntry | undefined {
  const key = modelId.toLowerCase()
  if (PRICE_TABLE[key]) return PRICE_TABLE[key]
  let best: string | null = null
  for (const k of Object.keys(PRICE_TABLE)) {
    if (key.startsWith(k) && (best === null || k.length > best.length)) best = k
  }
  return best ? PRICE_TABLE[best] : undefined
}

export function isLocalModel(modelId: string): boolean {
  const providerId = modelId.includes('/') ? modelId.split('/')[0] : modelId
  return ['ollama', 'lm-studio', 'llama-cpp', 'vllm', 'localai', 'opencode'].includes(providerId)
}
