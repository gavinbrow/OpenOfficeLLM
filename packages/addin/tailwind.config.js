/** @type {import('tailwindcss').Config} */
//
// The design tokens live in the shared preset (packages/ui/tailwind-preset.js).
// This file exists to say where the classes are: both this shell's own markup
// and every component in the UI package, which is where nearly all of them are.
import preset from '@openofficellm/ui/tailwind-preset'

export default {
  presets: [preset],
  content: ['./index.html', './src/**/*.{ts,tsx}', '../ui/src/**/*.{ts,tsx}'],
}
