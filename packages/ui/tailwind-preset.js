/** @type {import('tailwindcss').Config} */
//
// A Tailwind *preset*, shared by every shell. It deliberately carries no
// `content` globs: those are per-shell, and a preset that guessed them would
// silently drop the classes used by whichever shell it did not know about.
//
// Every semantic colour resolves through a CSS variable defined in index.css,
// with the dark values swapped on `.dark`. The alternative — light values here
// plus a `dark:` override at every call site — is what produced white-on-white
// dropdowns and light grey panels floating in a dark pane: one missed override
// is an unreadable control, and there is no way to notice the miss except by
// looking at it. With variables, a component that forgets dark mode cannot
// exist.
export default {
  darkMode: 'class',
  theme: {
    extend: {
      colors: {
        surface: {
          // Page background.
          DEFAULT: 'rgb(var(--surface) / <alpha-value>)',
          // Floating surfaces (menus, dialogs) that must read as *above* the page.
          raised: 'rgb(var(--surface-raised) / <alpha-value>)',
          // Recessed fills: inputs, chips, code blocks.
          muted: 'rgb(var(--surface-muted) / <alpha-value>)',
          hover: 'rgb(var(--surface-hover) / <alpha-value>)',
          border: 'rgb(var(--surface-border) / <alpha-value>)',
          // Deliberately inverted chrome (toasts).
          inverse: 'rgb(var(--surface-inverse) / <alpha-value>)',
        },
        // Foreground/text ramp.
        fg: {
          DEFAULT: 'rgb(var(--fg) / <alpha-value>)',
          inverse: 'rgb(var(--fg-inverse) / <alpha-value>)',
        },
        muted: 'rgb(var(--fg-muted) / <alpha-value>)',
        faint: 'rgb(var(--fg-faint) / <alpha-value>)',
        accent: {
          DEFAULT: 'rgb(var(--accent) / <alpha-value>)',
          hover: 'rgb(var(--accent-hover) / <alpha-value>)',
          // Accent used as *text or an icon*, which needs more contrast against
          // a dark background than the same blue used as a button fill.
          fg: 'rgb(var(--accent-fg) / <alpha-value>)',
          soft: 'var(--accent-soft)',
          line: 'var(--accent-line)',
        },
        ok: 'rgb(var(--ok) / <alpha-value>)',
        warn: 'rgb(var(--warn) / <alpha-value>)',
        danger: 'rgb(var(--danger) / <alpha-value>)',
      },
      fontFamily: {
        sans: [
          'Segoe UI Variable',
          'Segoe UI',
          '-apple-system',
          'BlinkMacSystemFont',
          'Roboto',
          'Helvetica',
          'Arial',
          'sans-serif',
        ],
        mono: ['Cascadia Code', 'Consolas', 'ui-monospace', 'Menlo', 'monospace'],
      },
      spacing: {
        pane: '320px',
      },
      boxShadow: {
        pane: '0 2px 8px rgba(0,0,0,0.08)',
        pop: '0 8px 24px rgba(0,0,0,0.28)',
      },
      borderRadius: {
        xl2: '10px',
      },
    },
  },
  plugins: [],
}
