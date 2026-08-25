import { useUiStore } from '../store/uiStore'
import { DownloadIcon } from './icons'

export function UpdateBadge() {
  const updateAvailable = useUiStore((s) => s.updateAvailable)
  const latestVersion = useUiStore((s) => s.latestVersion)
  const updateCountdown = useUiStore((s) => s.updateCountdown)
  const startUpdateCountdown = useUiStore((s) => s.startUpdateCountdown)

  if (!updateAvailable || updateCountdown !== null) return null

  return (
    <button
      className="btn btn-ghost h-8 w-8 text-accent"
      onClick={startUpdateCountdown}
      aria-label={`Update to v${latestVersion ?? ''}`}
      title={`Update to v${latestVersion ?? ''}`}
    >
      <DownloadIcon size={16} />
    </button>
  )
}
