# OpenOfficeLLM — Architecture Plan

> Status: Phases 0–5 implemented and verified end to end; Phase 6 (installer)
> and Phase 7 (release hardening) pending. See [TODO.md](TODO.md).
> Scope: v1 — Windows desktop, Word + Excel (Chrome extension also working;
> macOS host supported for the extension only).
> Companion documents: [TODO.md](TODO.md) · [SPIKE-LNA.md](SPIKE-LNA.md) ·
> [EXTENSION.md](EXTENSION.md)

> **Phase 0 update (2026-08-16) — GATE 0 PASSED FOR WORD.** The loopback escape
> hatch in §2 is **confirmed on real Office**: the task pane, served from
> `https://127.0.0.1:7317`, fetched `http://127.0.0.1:11434/api/tags` in **5 ms**
> inside Word on Office WebView2 151, while the permissions policy reported
> `local-network-access: false`. The architecture below stands.
>
> Refinement to §2: on Windows desktop the pane is a **top-level WebView2
> document, not an iframe** (`inIframe: false`). The Permissions-Policy
> delegation failure in office-js#6281 is an _Office-on-the-web_ problem and does
> not apply to the desktop host.
>
> Excel confirmed 2026-08-17. The once-per-session insert behaviour is
> documented in [SPIKE-LNA.md](SPIKE-LNA.md) and ADR-004.

## 1. Context

Build a Claude-style AI assistant for Microsoft Office: a chat panel docked to the right of the document, a model selector, and document-aware editing — for **Word and Excel on Windows desktop**, with **first-class local model support** (Ollama, LM Studio, llama.cpp) alongside all major cloud providers, plus opencode integration. A Chrome side-panel extension shares the same UI and host.

This plan was written at the start of the project and is retained as the
architecture reference. The phased status lives in [TODO.md](TODO.md); some
structure has evolved since the original layout below (notably the
`packages/ui` split and the `packages/extension` shell).

## 2. The constraint that determines the architecture

Chromium 142+ ships **Local Network Access (LNA)** restrictions. An HTTPS page served from a public origin can no longer call `http://127.0.0.1`; the request fails with:

```
Access to fetch at 'http://localhost:11434' from origin 'https://example.com' has been blocked
by CORS policy: Permission was denied for this request to access the `unknown` address space.
```

This breaks the standard Office add-in architecture for local models. Office embeds add-ins in iframes **without** the `local-network-access` permissions policy, so Chrome cannot even show the user a permission prompt. It affects desktop Office (WebView2, Windows + Mac) and Office on the web. Microsoft has it in backlog with no timeline:

- [office-js#6281 — Production Office Add-ins Cannot Connect to Desktop Applications in Chrome 142+](https://github.com/OfficeDev/office-js/issues/6281)
- [office-js#6366 — Network Calls Blocked by WebView2 PNA Policy](https://github.com/OfficeDev/office-js/issues/6366)
- [office-js#6426 — Outlook blocking addin request to localhost webserver](https://github.com/OfficeDev/office-js/issues/6426)

The only workaround is an enterprise Chromium policy (`LocalNetworkAccessRestrictionsTemporaryOptOut`) — unacceptable for a consumer install.

### The escape hatch

Chromium enforces LNA only for _public→local_ requests. **Loopback→loopback is exempt.** If the task pane HTML is itself served from `https://127.0.0.1`, its calls to local model servers are not restricted.

This inverts the usual design: instead of hosting the add-in on a CDN, a **local host service serves the add-in UI and brokers every model call**. That service also solves CORS, mixed content, and API-key storage as a side effect — keys live in the OS credential store and never enter the webview.

> **This is the load-bearing assumption of the entire project.** It is verified by a spike in Phase 0 before any other code is written.

## 3. Why not fork an existing project

| Project                                                                                                                                                                          | License | Stack             | Hosts      | Transport                | Verdict                                             |
| -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------- | ----------------- | ---------- | ------------------------ | --------------------------------------------------- |
| [word-GPT-Plus](https://github.com/Kuingsmile/word-GPT-Plus) (1.3k★)                                                                                                             | MIT     | Vue/TS/Vite       | Word       | pane → provider directly | Broken by LNA; keys in localStorage                 |
| [wordollama-community](https://github.com/ByronLeeeee/wordollama-community) (18★)                                                                                                | GPL-3.0 | VSTO/.NET 4.8 COM | Word + WPS | in-process COM           | Different architecture; no installer; viral license |
| [OllamaWord](https://github.com/arthuc01/OllamaWord), [Word-Add-in-Ollama](https://github.com/ascii-phoenix/Word-Add-in-Ollama), [WordLLMs](https://github.com/kauttoj/WordLLMs) | mixed   | various           | Word       | direct / ad-hoc helper   | Demo-scale                                          |

None solve LNA, none support Excel, and none store secrets outside the browser. word-GPT-Plus is a good reference for Office.js edit patterns and its agent-tool catalog (MIT, safe to learn from), but its transport layer — the bulk of our work — is exactly what's broken.

**Decision: build fresh**, in React, matching the author's established React
stack (React 18 + Vite + TypeScript + Tailwind + Zustand).

## 4. Decisions

| Decision             | Choice                                                                   | Rationale                                                                                                                           |
| -------------------- | ------------------------------------------------------------------------ | ----------------------------------------------------------------------------------------------------------------------------------- |
| Platform             | Windows desktop first; Mac v2                                            | Same architecture, different installer + cert store                                                                                 |
| Office hosts         | Word + Excel                                                             | One XML manifest declares both                                                                                                      |
| opencode             | Import config/models/skills/MCP; optional agent backend                  | Reuses existing user setup without a hard dependency                                                                                |
| Edit authority       | User-selectable: **Propose** / **Direct** / **Agentic**                  | User's explicit requirement                                                                                                         |
| Codebase             | Fresh — React 18 + Vite + TS + Tailwind + Zustand                        | The author's established React stack                                                                                                |
| Installer            | Host service + add-in (~30–60 MB); detects Ollama / LM Studio / opencode | Fast install, no duplicated runtimes                                                                                                |
| Extensibility        | Skills **and** MCP tools                                                 | Both in v1                                                                                                                          |
| Distribution         | Open source (MIT), unsigned installer initially                          | Signing is a later release-stage task, no code change                                                                               |
| Manifest format      | XML, not unified JSON                                                    | Unified is GA for Word/Excel as of July 2026 but **not supported for local/shared-folder deployment** — our exact distribution path |
| Web framework (host) | Hono                                                                     | Tiny, first-class SSE, same framework opencode uses                                                                                 |

## 5. Architecture

```
                Word / Excel (WebView2)
                          │
          task pane iframe │ SourceLocation = https://127.0.0.1:7317/taskpane.html
                          ▼
   ┌──────────────────────────────────────────────┐
   │  Local Host Service  (Node, loopback, TLS)   │  ← serves the UI *and* the API
   │  • static task pane bundle                   │     so the pane is a loopback origin
   │  • /api/chat  (SSE)                          │
   │  • secrets in Windows DPAPI                  │
   │  • MCP client · skills · opencode import     │
   └───┬─────────────┬──────────────┬─────────────┘
       │ loopback    │ loopback     │ outbound HTTPS
       ▼             ▼              ▼
   Ollama:11434  LM Studio:1234   Anthropic / OpenAI / Google /
   llama.cpp     opencode:4096    Azure / OpenRouter / Groq / …
```

The task pane only ever talks to `https://127.0.0.1:7317` — **same-origin**. No CORS, no LNA, no mixed content, and no API key ever reaches the browser context.

### Trust boundaries

| Boundary                 | Control                                                                               |
| ------------------------ | ------------------------------------------------------------------------------------- |
| Webview → host service   | Origin check + per-launch bearer token injected into `taskpane.html`                  |
| Host service → secrets   | Windows DPAPI, `%APPDATA%\OpenOfficeLLM\secrets.dat`; never serialized to the webview |
| Host service → MCP tools | Disabled by default; per-server enable; optional per-call approval                    |
| Host service → document  | Only via edits the pane requests; the service never touches Office directly           |

### Repo layout

```
OpenOfficeLLM/
├─ Docs/
│  ├─ PLAN.md              this document
│  ├─ TODO.md              phased task list
│  ├─ SPIKE-LNA.md         Phase 0 evidence
│  ├─ EXTENSION.md         the Chrome side panel
│  └─ DECISIONS.md         ADR log
├─ packages/
│  ├─ host/                Node service: TLS, providers, MCP, skills, secrets
│  ├─ shared/              Protocol types shared by every package
│  ├─ ui/                  The assistant UI — host-agnostic; shells supply a Shell
│  ├─ addin/               Office shell: Office.js, Word/Excel adapters
│  └─ extension/           Chrome shell: MV3 side panel, page adapter
├─ scripts/                Shared build-time scripts (icon generation)
└─ spike/                  Phase 0 throwaway proof-of-concept (kept as regression)
```

The installer (Inno Setup, Phase 6) is not yet built; the `npm run setup`
provisioning path is what ships today.

## 6. Component design

### 6.1 Host service (`packages/host`)

Node 22 + TypeScript + Hono.

- **TLS bootstrap** — on first run generate a local CA + leaf cert (SANs: `localhost`, `127.0.0.1`) with `node-forge`; install the CA into `Cert:\CurrentUser\Root`. Per-user, **no admin required**.
- **Port selection** — default `7317`, scan upward if taken. The chosen port is templated into both the service config and the manifest at install time.
- **Static hosting** — serves the built `packages/addin` bundle at `/`.
- **Local auth** — reject requests whose `Origin` isn't the service's own; issue a per-launch token injected into `taskpane.html` and required on `/api/*`.
- **Secrets** — Windows DPAPI is invoked at runtime through Windows PowerShell using .NET `System.Security.Cryptography.ProtectedData`, with an AES-256-GCM fallback when DPAPI is unavailable (always the case on macOS). `GET /api/providers` returns configured-or-not booleans only.

**API surface**

| Method  | Path                   | Purpose                                             |
| ------- | ---------------------- | --------------------------------------------------- |
| GET     | `/api/health`          | Liveness + version + port                           |
| GET     | `/api/providers`       | Discovered + configured providers                   |
| GET     | `/api/models`          | Normalized model list across providers              |
| POST    | `/api/chat`            | SSE: `delta`, `tool_call`, `usage`, `done`, `error` |
| POST    | `/api/chat/cancel`     | Abort an in-flight stream                           |
| GET/PUT | `/api/settings`        | Read/write user settings                            |
| GET     | `/api/skills`          | Available skills                                    |
| GET     | `/api/mcp/servers`     | Configured MCP servers + tools                      |
| POST    | `/api/mcp/consent`     | Grant/revoke tool permission                        |
| POST    | `/api/import/opencode` | Import opencode config                              |

### 6.2 Provider layer (`packages/host/src/providers/`)

```ts
interface ProviderAdapter {
  id: string
  listModels(): Promise<ModelInfo[]>
  stream(req: ChatRequest, signal: AbortSignal): AsyncIterable<StreamEvent>
  capabilities: { tools: boolean; vision: boolean; streaming: boolean }
}
```

| Adapter             | Covers                                                                                                             |
| ------------------- | ------------------------------------------------------------------------------------------------------------------ |
| `openai-compatible` | **LM Studio** (1234), llama.cpp (8080), vLLM, LocalAI, OpenRouter, Groq, DeepSeek, Together, xAI, custom base URLs |
| `ollama`            | Native `/api/chat` + `/api/tags` (11434)                                                                           |
| `anthropic`         | Messages API, streaming + tool use                                                                                 |
| `google`            | Gemini                                                                                                             |
| `azure-openai`      | Azure deployments                                                                                                  |
| `opencode`          | `POST /session`, `POST /session/:id/prompt_async`, SSE `GET /global/event`; honors `OPENCODE_SERVER_PASSWORD`      |

`discovery.ts` probes 11434 / 1234 / 4096 / 8080 on startup and auto-registers whatever answers. This is what makes first run feel zero-config.

### 6.3 Task pane (`packages/addin`)

React 18 + Vite + TS + Tailwind + Zustand.

| Component           | Role                                                                 |
| ------------------- | -------------------------------------------------------------------- |
| `ChatPanel`         | Message list, streaming markdown, stop button, token/cost readout    |
| `ModelSelector`     | Grouped by provider; local models badged; search; last-used pinning  |
| `ModeToggle`        | **Propose / Direct / Agentic** segmented control, persisted per host |
| `ContextChips`      | Shows exactly what's being sent; user can drop any of it             |
| `DiffPreview`       | Proposed-change preview with Apply / Discard                         |
| `ServiceDownScreen` | Recovery state when the host service isn't reachable                 |

Office.js loads from `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`; the host service caches a copy and rewrites the tag when offline.

### 6.4 Office integration (`packages/addin/src/office/`)

```ts
interface HostAdapter {
  getContext(scope: ContextScope): Promise<DocumentContext>
  applyEdits(edits: Edit[]): Promise<void>
  snapshot(): Promise<Snapshot>
  restore(s: Snapshot): Promise<void>
}
```

- **`word.ts`** — selection, `body.getRange()`, paragraphs, `body.search()`, content controls, comments, tracked changes. Snapshot via `body.getOoxml()`, restore via `insertOoxml(..., 'Replace')`.
- **`excel.ts`** — selected range, used range **with sampling** (head/tail rows + schema inference; never dump a 100k-row sheet), sheet names, formulas, named ranges, tables. Snapshot captures values + formulas of touched sheets.

**Edit modes**

| Mode                  | Behaviour                                                                                      |
| --------------------- | ---------------------------------------------------------------------------------------------- |
| **Propose** (default) | Render a diff; user clicks Apply                                                               |
| **Direct**            | Applies in a single `context.sync()` batch so one Ctrl+Z reverts it                            |
| **Agentic**           | Multi-step tool loop; snapshot before the run, "Revert run" after; step cap + per-step display |

### 6.5 Skills, MCP, opencode import

- **Skills** — markdown + YAML frontmatter (`name`, `description`, `hosts`, `mode`, `model?`). Loaded from built-ins, `%APPDATA%\OpenOfficeLLM\skills\`, and `~/.config/opencode/skills/`. Surface as one-click buttons above the composer.
- **MCP** — host service acts as MCP client (stdio + streamable HTTP). Servers imported from opencode's `mcp` block. **Tools disabled by default.**
- **opencode import** — strictly read-only:
  - `%USERPROFILE%\.config\opencode\opencode.json(c)` → providers, models, `mcp`, `agent`, `instructions`
  - `.config\opencode\{skills,agents,commands}\`
  - `~/.local/share/opencode/auth.json` → **opt-in, explicit confirmation**, re-encrypted into DPAPI rather than copied in plaintext

  Must tolerate a bare config. On the current development machine `opencode.jsonc` contains only `$schema`, `auth.json` holds only `ollama-cloud`, and the rest of the setup lives in the opencode desktop app's 139 MB session DB.

### 6.6 Manifest + installer

**Manifest** — one XML manifest covering both hosts:

```xml
<Hosts><Host Name="Document"/><Host Name="Workbook"/></Hosts>
<DefaultSettings>
  <SourceLocation DefaultValue="https://127.0.0.1:7317/taskpane.html"/>
</DefaultSettings>
```

**Installer** — Inno Setup, **per-user, no admin**, into `%LOCALAPPDATA%\Programs\OpenOfficeLLM`:

1. Copy host service (Node SEA single executable) + built add-in bundle
2. Pick a free port; template it into the manifest and service config
3. Generate + trust the local CA in `Cert:\CurrentUser\Root`
4. Register: `HKCU\Software\Microsoft\Office\16.0\WEF\Developer\<manifestPath>` = manifest path — no network share, no admin, the same mechanism `npm start` uses
   - Fallback for locked-down machines: `HKCU\...\WEF\TrustedCatalogs\{GUID}` with `Url` = `\\localhost\<share>`, `Flags`=1
5. Autostart via `HKCU\...\Run`; start the service
6. Uninstall reverses every step, **including removing the CA**

## 7. Risks

| Risk                                                            | Mitigation                                                                                                                                       |
| --------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| **LNA exemption doesn't hold in Office's WebView2**             | **Phase 0 spike proves it before anything else is built.** If it fails, the project pivots to a cloud-only add-in and local support is re-scoped |
| **Manifest hardcodes a port**; a later conflict breaks the pane | Installer picks the port and templates it; ship a `repair` command that re-provisions manifest + config                                          |
| Pane is dead if the service isn't running                       | Autostart at logon + a friendly "start service" recovery screen. A real UX cost — the honest trade for local models working at all               |
| `WEF\Developer` registration is technically a dev channel       | Documented, admin-free, and what every Office dev tool uses; trusted-catalog fallback for locked-down environments                               |
| Unsigned installer trips SmartScreen                            | Accepted for now; document "More info → Run anyway". Authenticode later needs no code change                                                     |
| Microsoft fixes LNA and CDN hosting becomes viable              | Provider layer is transport-agnostic; a hosted build reuses everything except TLS bootstrap and discovery                                        |
| Large Excel sheets blow the token budget                        | Sampling + schema inference in `excel.ts`; context chips show exactly what's sent                                                                |
| Office.js CDN dependency conflicts with offline-first           | Host service caches a copy and serves it when offline                                                                                            |
| Office cache serves a stale task pane during development        | Documented cache-clear step; cache-busting query param in dev builds                                                                             |

## 8. Out of scope for v1

Mac · Office on the web · iPad · Outlook · PowerPoint · AppSource / Microsoft Marketplace · multi-user or shared-server deployment · code signing.
