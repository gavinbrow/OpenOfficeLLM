# OpenOfficeLLM — Phased Task List

> Companion to [PLAN.md](PLAN.md). Task IDs are stable references (`P2.4`),
> so they can be cited in commits and issues.
>
> This file lists **remaining work only**. Completed tasks are removed; the
> history lives in git. Phases 0–5 are implemented and verified end to end
> (Word and Excel, the Chrome extension, skills, MCP, opencode import). The
> installer (Phase 6) is built and produces a working `.exe`; what is left is
> preflight checks, release hardening, and a handful of polish items.

**Legend** — `[ ]` not started · `[~]` in progress · `[x]` done · `[!]` blocked

---

## Phase 4 leftovers — Office integration

- [ ] **P4.9** Tracked-changes awareness — detect if tracking is on; surface
      it; optionally apply edits as tracked revisions
- [~] **P4.10** Comments — writing works (`add_comment` → `insertComment`);
  reading existing comments as context is not implemented
- [~] **P4.18** **Propose** — the review card exists (Apply / Apply one /
  Discard), but the preview is plain text; a word-level diff for Word and a
  cell-level table diff for Excel are not implemented
- [~] **P4.21** **Agentic** — snapshots are captured and restorable, but
  there is no "Revert run" button in the UI, and no warning when the user
  edited the document manually during the run
- [~] **P4.23** Token budgeting — per-item estimates exist; no warn/trim
  flow or hard-stop past the model's limit

## Phase 5 leftovers — Skills, MCP, opencode import

- [~] **P5.12** Per-call MCP approval — consent is per-tool enable in
  settings; the `McpConsentDialog` (show server, tool, and full arguments
  before a call; Allow once / Allow always / Deny) is not implemented
- [ ] **P5.24** Detect the opencode **desktop app** install
      (`@opencode-aidesktop`) and its running server; offer the opencode agent
      backend when reachable

---

## Phase 6 — Installer

> The installer is built: `npm run build:installer` produces
> `dist/OpenOfficeLLM-Setup-<version>.exe` via Node SEA + Inno Setup. It
> installs per-user (no admin), runs `host.exe --install` to provision, and
> upgrades in place (stable AppId, idempotent `--install`). See
> `installer/installer.iss` and `scripts/build-installer.mjs`.

- [ ] **P6.6** Validate the generated manifest with
      `office-addin-manifest validate` in CI, **failing the build on any error**.
      Errors that read like AppSource store policy can be hard load blockers —
      the sub-1.0 `<Version>` in P0.2 was reported only by this tool and was the
      reason the add-in would not load at all
- [ ] **P6.11** Preflight checks — Office 16.0 present, WebView2 runtime
      present (offer the MS bootstrapper if missing), warn if Office is running
- [ ] **P6.18** Post-install page — "Open Word to get started", link to
      docs, note about SmartScreen for anyone who shares the installer
- [~] **P6.21** `--diagnose` — prints port, manifest path, CA thumbprint +
  trust state, registry state, autostart command, and live health. Still
  missing: port ownership by PID and image path, and detected providers
- [ ] **P6.22** Document the SmartScreen path ("More info → Run anyway") in
      README and on the download page, since v1 ships unsigned

> **GATE 6** — clean Windows VM with **no Node installed**: installer runs
> without admin, Word shows the add-in, a local model answers, uninstall
> leaves nothing behind.

---

## Phase 7 — Verification, docs, release

### P7.A — End-to-end verification

Run the full script on a clean VM:

- [ ] **P7.1** Install with no Node, no admin rights
- [ ] **P7.2** Word → Home → Add-ins → OpenOfficeLLM. Pane docks right; model
      selector lists discovered Ollama models with zero configuration
- [ ] **P7.3** Select a paragraph → "rewrite this more formally" in
      **Propose** → diff → Apply → document updates correctly
- [ ] **P7.4** Switch to **Direct** → repeat → one Ctrl+Z reverts cleanly
- [ ] **P7.5** Switch to **Agentic** → multi-step edit → "Revert run"
      restores the original
- [ ] **P7.6** Excel → select a range → ask for a formula → applies to the
      correct cells
- [ ] **P7.7** Excel → 50k-row sheet → confirm sampling keeps the request in
      budget and the answer stays useful
- [ ] **P7.8** Add an Anthropic key → model list expands → **confirm via
      devtools that no key appears in any network response or storage**
- [ ] **P7.9** With opencode installed → import → models, skills, and MCP
      servers appear
- [ ] **P7.10** Kill the host service → pane shows the recovery screen, not a
      blank frame → restart from the pane works
- [ ] **P7.11** Disconnect the network → local models still work end to end,
      including Office.js loading from cache
- [ ] **P7.12** Reboot → service autostarts → pane works with no intervention
- [ ] **P7.13** Uninstall → add-in gone from Office, CA removed from the cert
      store, no stray registry keys

### P7.B — Hardening

- [ ] **P7.14** Security review — token handling, origin enforcement, DPAPI
      usage, MCP consent, markdown sanitization, path traversal on the static
      server
- [ ] **P7.15** Performance — pane cold start < 2 s, first token < 1 s on a
      local model, no jank while streaming into a long conversation
- [ ] **P7.16** Long-conversation behavior — 100+ messages: virtualization
      holds, memory stays flat, context trimming is sane
- [ ] **P7.17** Crash resilience — kill the service mid-stream; the pane
      recovers without losing history
- [ ] **P7.18** Office cache staleness — document the clear-cache procedure;
      verify updates land after reinstall

### P7.C — Documentation

- [~] **P7.19** `README.md` — text is done; screenshots missing
- [ ] **P7.20** `Docs/PROVIDERS.md` — setup for each provider, local and
      cloud
- [ ] **P7.21** `Docs/SKILLS.md` — authoring skills, frontmatter reference,
      examples
- [ ] **P7.22** `Docs/MCP.md` — adding servers, the consent model, security
      notes
- [ ] **P7.23** `Docs/TROUBLESHOOTING.md` — service won't start, port
      conflicts (`--repair`), cert warnings, add-in missing from the ribbon,
      Office cache, SmartScreen
- [ ] **P7.24** `Docs/ARCHITECTURE.md` — expanded from PLAN.md for
      contributors
- [ ] **P7.25** Issue templates

### P7.D — Release

- [ ] **P7.26** Version + changelog
- [x] **P7.27** GitHub Actions release job — build installer, attach to the
      release (`.github/workflows/release.yml`)
- [ ] **P7.28** Tag `v0.1.0`, publish, **clearly marked unsigned** with the
      SmartScreen note
- [ ] **P7.29** Open a tracking issue for Authenticode signing (no code
      change required)
- [ ] **P7.30** Open a tracking issue watching
      [office-js#6281](https://github.com/OfficeDev/office-js/issues/6281) — if
      Microsoft ships the `local-network-access` permissions policy, a
      CDN-hosted variant becomes possible

---

## CI

- [ ] **P7.31** Add `format:check` and a build step to
      `.github/workflows/ci.yml` (currently typecheck + lint + test only)

---

## Deferred to v2+

| Item                     | Notes                                                                                                                                  |
| ------------------------ | -------------------------------------------------------------------------------------------------------------------------------------- |
| Developer ID signing     | Apple's sibling of Authenticode: ~$99/yr, clears Gatekeeper; pair with notarization for the DMG                                        |
| Keychain-backed secrets  | macOS currently uses the AES-256-GCM fallback; `security`-backed key storage would match DPAPI strength                                |
| PowerPoint               | Weakest Office.js surface of the three                                                                                                 |
| Outlook                  | Most aggressive localhost blocking reported ([#6426](https://github.com/OfficeDev/office-js/issues/6426)); needs its own investigation |
| Office on the web        | Requires a hosted build and cloud-only models                                                                                          |
| Authenticode signing     | ~$100–400/yr; EV clears SmartScreen immediately                                                                                        |
| AppSource listing        | Incompatible with loopback hosting as designed                                                                                         |
| RAG over local documents | Natural extension once the provider layer is stable                                                                                    |
| Multi-document context   | Read from other open workbooks/documents                                                                                               |
