# The Chrome extension

The same assistant as the Word and Excel task pane, in a browser side panel,
reading the tab you are looking at.

It is not a separate product and not a separate codebase. Both shells mount the
same React app from `packages/ui` against the same local host service, so a
model you configured for Word is already available in Chrome, and a fix to the
composer lands in both.

---

## What it can and cannot do

**It reads.** The main readable content of the page, your selection, the
heading outline, page metadata, links, and a search across the page — five
tools, all read-only.

**It does not write.** This is a design decision, not an unfinished feature.
A Word document belongs to the user, so the Word adapter edits it. A web page
belongs to whoever served it and is rendered for the user to read; an assistant
that rewrites it in place produces something that looks authoritative,
screenshots convincingly, and corresponds to nothing anyone published. The
browser adapter therefore omits `applyEdits`, `snapshot` and `restore`
entirely — the `HostAdapter` contract marks them optional so a read-only host
can say so in the type system rather than by throwing when someone tries.

The propose/direct/agentic mode toggle still appears, because the shared UI is
shared, but with no write tools in the catalog there is nothing for it to gate.

---

## Setup

### 1. Build it

```bash
npm run build:extension
```

Output lands in `packages/extension/dist`, ready to load unpacked.

### 2. Load it in Chrome

Go to `chrome://extensions`, turn on **Developer mode**, choose **Load
unpacked**, and pick `packages/extension/dist`. Chrome assigns the extension a
32-character id; you need it for the next step, and it is shown on that card.

### 3. Pair it with the host

```bash
node packages/host/dist/index.js --pair <extension-id>
```

Until you do this the panel refuses to connect, and says so with the exact
command to run. To see or undo pairings:

```bash
node packages/host/dist/index.js --list-pairs
node packages/host/dist/index.js --unpair <extension-id>
```

Pairing takes effect immediately — the allowlist is read per request, so there
is no host restart.

### 4. Grant sites as you go

The extension declares **no** site permissions up front. When you open the panel
on a page it has not been granted, a strip offers to allow that one origin.
Chrome's own permission prompt follows.

This is deliberate. An extension that asks for `<all_urls>` at install time is
asking to read your banking session, and "it only reads when you press send" is
not a claim you can verify. Per-site, on a click you initiated, is.

---

## Security model

The task pane gets two things for free that the extension cannot have:

|                | Task pane                                 | Extension                 |
| -------------- | ----------------------------------------- | ------------------------- |
| Origin         | `https://127.0.0.1:7317` — the host's own | `chrome-extension://<id>` |
| Token delivery | Injected into the served HTML             | Must be requested         |

So the extension needs an explicit trust decision, and that is the entire
purpose of the allowlist:

- **`GET /pair`** returns the bearer token, and answers **only** to an origin on
  the allowlist. It sits outside `/api/*` because it is the one request an
  extension can make before it has a token.
- **Every `/api/*` route** accepts the pane's own origin or a paired extension
  origin, and nothing else — _including_ a request that presents a valid token.
  A leaked token from a web page is still refused.
- **Matching is exact.** No wildcards, no prefixes, no case folding. See
  [`packages/host/src/pairing.ts`](../packages/host/src/pairing.ts) and its
  tests, which are mostly about what must _not_ be admitted.

### The `X-OpenOfficeLLM-Extension` header

Chrome does not reliably attach an `Origin` to a fetch made from an extension
page that holds host permission for the target — it treats the request as
privileged rather than cross-origin. The extension therefore also names itself
in a custom header.

That is a sound CSRF signal in its own right: a web page cannot set a custom
header on a cross-origin request without triggering a preflight, and the host's
`OPTIONS` handler refuses every origin that is not already paired. When both
`Origin` and the header are present they must agree, so a paired extension's id
cannot simply be borrowed by anything that learned it.

### TLS

The host serves loopback HTTPS with a locally generated CA. Chrome reads the OS
trust store, so `--install` (or `--trust-cert`) is what makes the extension able
to reach it at all. A certificate error in the panel's connection screen is
almost always this.

---

## macOS

The extension is inherently cross-platform; the host is what needed work.

- **Config, secrets, certs and history** now live at
  `~/Library/Application Support/OpenOfficeLLM` on macOS and under `$XDG_CONFIG_HOME`
  on Linux. Windows keeps `%APPDATA%\OpenOfficeLLM` unchanged — moving it would
  strand existing installs.
- **CA trust** is implemented via `security add-trusted-cert` into the login
  keychain, scoped to SSL and to `trustAsRoot` rather than `trustRoot`. It opens
  a Touch ID / password prompt, which is correct — you are authorising a
  certificate authority — but means it must be run from a terminal you are
  looking at, never from a background start.
- **Office add-in registration** (`--install`) is now platform-dispatched. On
  macOS the manifest is copied into each Office app's sideload folder
  (`~/Library/Containers/com.microsoft.{Word,Excel}/Data/Documents/wef/`),
  which Office re-reads at launch, and autostart is a per-user LaunchAgent
  (`~/Library/LaunchAgents/com.openofficellm.host.plist`) instead of the
  registry `Run` key. Windows keeps `HKCU\...\WEF\Developer` unchanged.

---

## How it fits together

```
                Chrome
   ┌──────────────────────────────────┐
   │  Side panel (chrome-extension://)│
   │    packages/ui  ← the assistant  │
   │    browserShell ← Shell impl     │
   └───┬──────────────────────────┬───┘
       │ chrome.tabs.sendMessage  │ https + bearer token
       ▼                          ▼
   content script            Local Host Service
   (read-only, IIFE)         https://127.0.0.1:7317
   in the active tab         providers · MCP · skills · secrets
```

| File                                                                             | Job                                                                 |
| -------------------------------------------------------------------------------- | ------------------------------------------------------------------- |
| [`src/sidepanel.tsx`](../packages/extension/src/sidepanel.tsx)                   | Entrypoint. Pair, attach to a tab, register the shell, mount.       |
| [`src/browser/shell.ts`](../packages/extension/src/browser/shell.ts)             | The `Shell` the UI talks to. Twenty lines.                          |
| [`src/browser/bootstrap.ts`](../packages/extension/src/browser/bootstrap.ts)     | Which tab, whether it is readable, per-tab chat identity.           |
| [`src/browser/pageAdapter.ts`](../packages/extension/src/browser/pageAdapter.ts) | `HostAdapter` over `chrome.tabs.sendMessage`, with injection retry. |
| [`src/browser/tools.ts`](../packages/extension/src/browser/tools.ts)             | The five read tools and their dispatcher.                           |
| [`src/browser/pairing.ts`](../packages/extension/src/browser/pairing.ts)         | Find the host's port, fetch a token, configure the API client.      |
| [`src/content/page.ts`](../packages/extension/src/content/page.ts)               | The only code that runs in someone else's page. Extraction.         |
| [`src/background.ts`](../packages/extension/src/background.ts)                   | Almost empty, on purpose — MV3 workers are killed at will.          |

### Notes for the next person

- **The content script is built separately** (`vite.content.config.ts`) as an
  IIFE. MV3 content scripts are classic scripts: no `import`, no code splitting.
  Its filename is referenced by `CONTENT_SCRIPT` in `pageAdapter.ts`.
- **Chats are keyed per tab**, not per URL. In Word each document gets a chat;
  the browser analogue is the tab, because that is what the user thinks of as
  "where I was". Keying on URL would start a fresh chat on every navigation,
  including a link the user followed because of what the assistant just said.
  Tab ids are reissued after a browser restart, so a restart lands on a fresh
  chat — the same failure direction the pane takes for an unsaved document.
- **Nothing stateful belongs in the service worker.** The agent loop, the
  conversation and the host connection all live in the panel, which stays alive
  while it is open.

---

## Known gaps

- **Pairing is CLI-only.** There is no UI for it in Settings yet; the panel
  prints the exact command instead.
- **Port discovery is a short scan** of 7317–7326, with the last good port
  remembered in `chrome.storage.local`. A host that lands outside that range
  will not be found.
- **Firefox is not supported.** The id format check in `pairing.ts` accepts
  Chromium ids only, and rejects Firefox's brace-wrapped UUIDs rather than
  silently admitting an origin shape nobody has thought about.
- **No PDF or `chrome://` support**, and there cannot be — no extension may
  script those. The panel says so rather than failing opaquely.
