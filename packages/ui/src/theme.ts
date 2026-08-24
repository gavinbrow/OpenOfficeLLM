// Theme application: reconcile the user's theme choice with whatever the
// surrounding application reports. 'office' = follow the shell's detection —
// Office's own theme in the task pane, the browser's colour scheme in the
// extension.

import { isDark as shellIsDark } from './host/bridge'
import type { Theme } from './store/uiStore'

export function applyTheme(theme: Theme): void {
  const isDark = theme === 'office' ? shellIsDark() : theme === 'dark'
  document.documentElement.classList.toggle('dark', isDark)
}
