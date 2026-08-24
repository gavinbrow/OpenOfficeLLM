# Contributing

OpenOfficeLLM is a small project with a specific architectural commitment
(the loopback-origin escape hatch described in the
[README](README.md) and [Docs/PLAN.md](Docs/PLAN.md)). Most contributions
that fit that architecture are welcome.

## Before you start

Read the [README](README.md) "Why this exists" section and
[Docs/PLAN.md](Docs/PLAN.md). The loopback-origin design is load-bearing — a
change that moves the task pane off `https://127.0.0.1` breaks local model
access on Chromium 142+, which is the entire reason the project exists. If
your change touches that assumption, open an issue first.

The phased roadmap and current status live in [Docs/TODO.md](Docs/TODO.md).
Tasks marked `[~]` are in progress; check the file before starting work to
avoid duplicating effort.

## Development setup

Requirements: Node 22, npm 11, Windows 11 (the host service and Office
integration are Windows-specific; the Chrome extension builds on macOS and
Linux too).

```bash
npm install        # workspaces resolve everything
npm run typecheck
npm run lint
npm test
npm run build:all  # host + task pane + extension
```

The CI workflow (`.github/workflows/ci.yml`) runs `typecheck`, `lint`, and
`test` on every push and PR. All three must pass.

## Code style

- **Prettier** is authoritative for formatting (`npm run format:check`). Run
  `npm run format` before committing.
- **ESLint** with `--max-warnings 0`. No new warnings.
- **TypeScript** strict mode; no `any` without a comment explaining why.
- No comments unless they explain _why_ — the code already says _what_.

## Tests

Vitest, with workspace-aware config at the repo root. Tests live alongside
source in `__tests__/` directories. The Ollama smoke test
(`packages/host/src/providers/__tests__/ollama.smoke.test.ts`) is gated
behind an env var and skips in CI — it needs a running Ollama.

When fixing a bug, add a test that fails before the fix and passes after.
When adding a feature, test the public surface.

## Commit messages

Short, imperative, lowercase or sentence case — match the existing log. The
repo does not use conventional-commits formatting or signed-off-by lines.

## Pull requests

1. Open an issue for anything beyond a small fix — the architecture is
   opinionated and a quick conversation saves wasted work.
2. Keep PRs focused. A PR that does one thing well reviews faster than a PR
   that refactors half the tree on the way.
3. Make sure `typecheck`, `lint`, `test`, and `format:check` all pass
   locally before pushing.
4. If your change affects the trust model (origins, tokens, secrets, MCP
   consent), describe the security implications in the PR description.

## Things that will not be accepted without a prior conversation

- Moving the task pane off the loopback origin (see above).
- Adding `<all_urls>` to the Chrome extension's permissions.
- Storing API keys anywhere other than DPAPI / the OS keychain.
- Enabling MCP tools by default. The opt-in consent model is a deliberate
  security property.
- Adding a dependency on a closed-source or telemetry-bearing package.

## Office add-in specifics

The add-in manifest is **generated, not checked in** — every URL in it is
absolute and carries the host's port, which is picked at startup. After
changing the ribbon or add-in identity in `packages/host/src/manifest.ts`,
rebuild and restart the host, then restart Word — Office re-reads a
developer-registered manifest when its mtime changes. See
[Docs/SPIKE-LNA.md](Docs/SPIKE-LNA.md) for why the once-per-session insert
behaviour is a property of sideloading, not a bug.
