/** @type {import('tailwindcss').Config} */
//
// Same tokens as the task pane, from the shared preset. The content globs have
// to reach into packages/ui because that is where nearly every class lives —
// this package contributes only the side-panel chrome.
import preset from '@openofficellm/ui/tailwind-preset'

export default {
  presets: [preset],
  content: ['./sidepanel.html', './src/**/*.{ts,tsx}', '../ui/src/**/*.{ts,tsx}'],
}
