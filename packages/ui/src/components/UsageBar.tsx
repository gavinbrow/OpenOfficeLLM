import { useChatStore } from '../store/chatStore'

export function UsageBar() {
  const usage = useChatStore((s) => s.usage)
  if (usage.totalTokens === 0 && usage.estimatedCostUsd === 0) return null
  const cost =
    usage.estimatedCostUsd > 0 ? `$${usage.estimatedCostUsd.toFixed(4)}` : 'local · no cost'
  return (
    <div
      className="flex items-center justify-between border-t border-surface-border px-2 py-1 text-[0.65rem] text-faint"
      role="status"
      aria-label="Session usage"
    >
      <span>
        ↑ {usage.promptTokens.toLocaleString()} · ↓ {usage.completionTokens.toLocaleString()} ·{' '}
        {usage.totalTokens.toLocaleString()} total
      </span>
      <span>{cost}</span>
    </div>
  )
}
