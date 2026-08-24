# OpenOfficeLLM

A Claude-style AI assistant for Microsoft **Word** and **Excel** on Windows desktop and for **Chrome**, with first-class local model support (Ollama, LM Studio, llama.cpp) alongside every major cloud provider, plus opencode integration.

One assistant, three surfaces: the Office task pane edits your document, the Chrome side panel reads the page you are on. Both run the same chat UI against the same local host service, the same providers, and the same keys.

> **Status:** pre-v1 / in development. See [`Docs/PLAN.md`](Docs/PLAN.md) and [`Docs/TODO.md`](Docs/TODO.md) for the roadmap.

## Why this exists — the Local Network Access problem

Chromium 142+ enforces **Local Network Access (LNA)** restrictions: an HTTPS page served from a public origin can no longer call `http://127.0.0.1`. Office embeds add-ins in iframes _without_ the `local-network-access` permissions policy, so Chrome cannot even prompt the user — the request just fails. This breaks the standard Office add-in architecture for local models, and Microsoft has it in backlog with no timeline ([office-js#6281](https://github.com/OfficeDev/office-js/issues/6281)).

### The escape hatch

Chromium only enforces LNA for _public → local_ requests. **Loopback → loopback is exempt.** If the task pane HTML is itself served from `https://127.0.0.1`, its calls to local model servers are not restricted.

This inverts the usual Office add-in design: instead of hosting the add-in on a CDN, a **local host service serves the add-in UI and brokers every model call**. That service also solves CORS, mixed content, and API-key storage as a side effect — keys live in the OS credential store (Windows DPAPI) and never enter the webview.

This is the load-bearing assumption of the entire project. It is verified by the Phase 0 spike — see [`Docs/SPIKE-LNA.md`](Docs/SPIKE-LNA.md).

## Architecture

```
            Word / Excel (WebView2)
                      │
             task pane │ SourceLocation = https://127.0.0.1:7317/index.html
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

## Repo layout

```
OpenOfficeLLM/
├─ Docs/
│  ├─ PLAN.md              architecture
│  ├─ EXTENSION.md         the Chrome side panel
│  ├─ SPIKE-LNA.md         Phase 0 evidence
│  ├─ DECISIONS.md         ADR log
│  └─ TODO.md              phased task list
├─ installer/
│  └─ installer.iss        Inno Setup script (per-user, no admin)
├─ packages/
│  ├─ host/                Node service: TLS, providers, MCP, skills, secrets
│  ├─ shared/              Protocol types shared by every package
│  ├─ ui/                  The assistant itself — chat, settings, stores.
│  │                       Host-agnostic; shells supply a Shell (host/bridge.ts)
│  ├─ addin/               Office shell: Office.js, Word/Excel adapters
│  └─ extension/           Chrome shell: MV3 side panel, page adapter
├─ scripts/                Build-time scripts (icon generation, SEA, installer)
└─ spike/                  Phase 0 throwaway proof-of-concept (kept as a
                            regression check for the LNA escape hatch)
```

The installer (Inno Setup) is a Phase 6 item and not yet built; `npm run
setup` is the current provisioning path.

**`packages/ui` is where nearly all the code is.** It knows nothing about
Office.js or `chrome.*`; it asks a registered `Shell` for the document adapter,
the tool catalog, and the identity of whatever is being worked on. Each shell is
a thin file that builds one — compare
[`packages/addin/src/office/shell.ts`](packages/addin/src/office/shell.ts) with
[`packages/extension/src/browser/shell.ts`](packages/extension/src/browser/shell.ts):
same shape, twenty lines each. Adding a third surface means writing a third one.

The add-in manifest is **generated, not checked in**. Every URL in it is absolute
and carries the port, and the host picks its port at startup, so a static file
would silently point at the wrong origin the moment anything else held 7317.
The host writes it to `%APPDATA%\OpenOfficeLLM\manifest\openofficellm.xml` from
the port it actually bound — see [`packages/host/src/manifest.ts`](packages/host/src/manifest.ts).

## Getting started

### For end users

Download `OpenOfficeLLM-Setup-<version>.exe` from the [releases
page](https://github.com/openofficellm/OpenOfficeLLM/releases) and run it. No
Node, no admin rights, no terminal. The installer:

1. Copies the host service and web bundle to `%LOCALAPPDATA%\Programs\OpenOfficeLLM`.
2. Runs `host.exe --install`, which trusts the local CA, writes the manifest,
   registers the add-in with Office, and sets up autostart.

> **SmartScreen:** Windows will warn "Windows protected your PC" because the
> installer is unsigned. Click **More info → Run anyway**. Future versions
> will be Authenticode-signed.

Then in **Word or Excel** — the steps are identical in both:

1. **Home** tab → **Add-ins**
2. Under **Developer Add-ins**, click **OpenOfficeLLM**
3. The **OpenOfficeLLM** ribbon group and its **AI Assistant** button appear on
   the Home tab, and the pane opens

> **You have to do this once per Word/Excel session,** and until you do, there
> is nothing on the ribbon to find. A manifest registered under `WEF\Developer`
> is scoped to the developer path — Office records it with `UniqueId: developer`
> and never writes the `AppStates` entry that store-installed add-ins get, which
> is the record that survives a restart. So the ribbon group is gone again next
> time you launch, and the add-in has to be opened from the Add-ins menu again.
> This is a property of sideloading, not a bug in the add-in — see
> [`Docs/SPIKE-LNA.md`](Docs/SPIKE-LNA.md) for the evidence.

### For developers

Requirements: Node 22, npm 11, Windows 11, and Word or Excel from Microsoft 365
(Office 16.0 with the WebView2 runtime).

```bash
npm install
npm run setup
```

`npm run setup` builds everything, then provisions the current user — no admin
rights at any point. It:

1. **Trusts the local certificate authority.** Windows shows a _Security
   Warning_ dialog asking you to confirm. This is expected and unavoidable:
   Windows prompts before any addition to the root store, even a per-user
   self-generated one, and the prompt happens below the API level so no
   tool can suppress it. Choose **Yes**.
2. Writes the manifest for the configured port.
3. Registers it with Office under `HKCU\…\Office\16.0\WEF\Developer`.
4. Adds a logon entry so the host service starts with Windows.

Then start the service and open Word:

```bash
npm start
```

Then, in **Word or Excel** — the steps are identical in both:

1. **Home** tab → **Add-ins**
2. Under **Developer Add-ins**, click **OpenOfficeLLM**
3. The **OpenOfficeLLM** ribbon group and its **AI Assistant** button appear on
   the Home tab, and the pane opens

> **You have to do this once per Word/Excel session,** and until you do, there
> is nothing on the ribbon to find. A manifest registered under `WEF\Developer`
> is scoped to the developer path — Office records it with `UniqueId: developer`
> and never writes the `AppStates` entry that store-installed add-ins get, which
> is the record that survives a restart. So the ribbon group is gone again next
> time you launch, and the add-in has to be opened from the Add-ins menu again.
> This is a property of sideloading, not a bug in the add-in — see
> [`Docs/SPIKE-LNA.md`](Docs/SPIKE-LNA.md) for the evidence and the options for
> a durable install.

If Office says **"Sorry, we can't load the add-in"**, the host service is not
running. Office loads the pane from it, so there is nothing to show. Run
`npm run diagnose` to confirm, then `npm start`.

The host service must be running whenever you use the add-in — Office loads the
pane from it, so a stopped service means a pane that will not load. Autostart
takes care of this from your next logon onward.

To check the state of an install:

```bash
npm run diagnose
```

To reverse it — unregisters the add-in, removes the logon entry, and removes
the CA by thumbprint, leaving your config and secrets in place:

```bash
npm run uninstall
```

## The Chrome extension

The side panel is the same assistant reading the active tab instead of a
document. It is **read-only** by design — a web page belongs to whoever served
it, so the browser adapter deliberately implements no write path at all.

```bash
npm run build:extension
```

Then load `packages/extension/dist` as an unpacked extension at
`chrome://extensions` with Developer mode on, and pair it with the host once:

```bash
node packages/host/dist/index.js --pair <extension-id>
```

The extension is served from its own origin, so unlike the task pane it is
cross-origin to the host and gets no injected token. Pairing is what lets it
read one — an unpaired extension is refused, whatever it presents. Full
walkthrough, security model and macOS notes:
[`Docs/EXTENSION.md`](Docs/EXTENSION.md).

## Development

```bash
npm run typecheck
npm run lint
npm test
npm run build:all     # host + task pane + extension
```

To build the installer (requires [Inno Setup 6](https://jrsoftware.org/isdl.php)):

```bash
npm run build:installer    # → dist/OpenOfficeLLM-Setup-<version>.exe
```

See `packages/host`, `packages/addin` and `packages/extension` for per-package
scripts. After
changing the ribbon or add-in identity in `packages/host/src/manifest.ts`,
rebuild and restart the host, then restart Word — Office re-reads a
developer-registered manifest when its mtime changes.

## License

MIT. See [`LICENSE`](LICENSE).
