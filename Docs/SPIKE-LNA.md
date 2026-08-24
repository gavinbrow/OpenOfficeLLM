# Phase 0 Spike — Local Network Access

> **GATE 0 PASSED FOR WORD AND EXCEL.** The load-bearing assumption of the
> entire architecture is confirmed on real Office. Excel passed 2026-08-17;
> the durable-install question is resolved in Result 4.
>
> Spike code: [`/spike`](../spike). It is throwaway, but keep it — it is the
> regression test for the day Chromium or Office changes this behaviour.

## Verdict

| Question                              | Answer                                         |
| ------------------------------------- | ---------------------------------------------- |
| Does loopback→loopback bypass LNA?    | **Yes** — Chromium 148 and Office WebView2 151 |
| Does it work inside a Word task pane? | **Yes** — 200 in 5 ms                          |
| Does it work in Excel?                | **Yes** — confirmed 2026-08-17                 |
| Can the add-in be installed durably?  | **No — session-scoped by design** (Result 4)   |

The architecture stands. Serving the task pane from `https://127.0.0.1` makes
local model calls same-origin and LNA never engages. What is _not_ solved is
getting the add-in to install and stay installed — see Result 4.

## Environment

|                  |                                                          |
| ---------------- | -------------------------------------------------------- |
| Office           | 16.0.20228.20190, O365 ProPlus Retail, x64               |
| WebView2 runtime | 151.0.4129.78                                            |
| Control browser  | Chromium 148 (Electron 42)                               |
| Ollama           | 0.32.9 on `127.0.0.1:11434`, one model (`glm-5.2:cloud`) |
| Node             | v22.12.0                                                 |
| OS               | Windows 11 Pro 26200                                     |
| Elevation        | none — every step ran non-elevated                       |

LNA shipped in Chromium 142. Both engines are well past it, so neither result is
a false pass from a stale engine.

## Result 1 — GATE 0 PASSED in Word

Verbatim from the pane, posted to the spike server's `/report` endpoint at
`2026-08-16T22:54:39Z`:

```json
{
  "gate": true,
  "href": "https://127.0.0.1:7317/taskpane.html?_host_Info=Word$Win32$16.01$en-US$$$$0",
  "userAgent": "...Chrome/151.0.0.0 Safari/537.36 Edg/151.0.0.0",
  "brands": ["Chromium 151", "Microsoft Edge WebView2 151", "Microsoft Edge 151"],
  "officeHost": "Word",
  "officePlatform": "PC",
  "officeDiagnostics": { "host": "Word", "version": "16.0.20228.20190", "platform": "PC" }
}
```

| Test                                             | Word / WebView2 151 | Chromium 148   |
| ------------------------------------------------ | ------------------- | -------------- |
| Same-origin control                              | 200 (9 ms)          | 200 (46 ms)    |
| **`fetch('http://127.0.0.1:11434/api/tags')`**   | **200 (5 ms)**      | **200 (5 ms)** |
| `fetch('http://localhost:11434/api/tags')`       | 200 (312 ms)        | 200 (310 ms)   |
| `fetch(..., { targetAddressSpace: 'loopback' })` | 200 (3 ms)          | 200 (2 ms)     |

Meanwhile the permissions policy reports the feature as **denied**:

```
local-network-access: false
private-network-access: false
network-ish features known to engine: local-network, loopback-network
```

That combination is the whole point. The engine knows about LNA, the page is not
granted the permission, and the fetch succeeds anyway — because loopback→loopback
never triggers the check.

Two incidental findings worth keeping:

- **`inIframe: false`.** On Windows desktop the task pane is a _top-level_
  WebView2 document, not an iframe. The Permissions-Policy delegation failure
  described in office-js#6281 is an _Office-on-the-web_ problem; it does not
  apply to the desktop host. Good news for this architecture, and it means a
  future CDN-hosted variant would still be blocked on the web but not here.
- `targetAddressSpace: 'loopback'` is accepted but unnecessary. Harmless to keep
  as a belt-and-braces hint.

`localhost` resolves ~60× slower than `127.0.0.1` (312 ms vs 5 ms) — almost
certainly IPv6 `::1` being tried first. **Always dial `127.0.0.1` literally.**

## Result 2 — a local CA cannot be trusted silently

The plan (P1.8, P6.13) specified:

```powershell
Import-Certificate -FilePath ca.crt -CertStoreLocation Cert:\CurrentUser\Root
```

That is wrong twice over.

1. `Import-Certificate` routes through `CryptUIWizImport` and **fails outright**
   non-interactively: `Import-Certificate : UI is not allowed in this operation.`
   Use the `X509Store` API — see [`spike/register.ps1`](../spike/register.ps1).

2. Switching API does **not** remove the prompt. Windows guards the user's Root
   store ("protected root") at the CryptoAPI level, so adding a CA always raises
   a **"Security Warning"** dialog listing the thumbprint. `mkcert` hits the same
   wall. There is no supported silent path.

**Consequence:** "no admin required" survives — the whole spike ran unelevated.
"Silent/unattended install" does not. The installer must show its own explanation
screen immediately _before_ triggering the dialog. Remove by **thumbprint**,
never subject name.

Cert requirements that actually matter to Chromium: SAN is mandatory (CN is
ignored), and `127.0.0.1` must be an **IP SAN (type 7)**, not a DNS SAN — a DNS
SAN of `127.0.0.1` is silently ignored. See
[`spike/gen-certs.mjs`](../spike/gen-certs.mjs).

## Result 3 — `<Version>` below 1.0 is silently fatal

The first manifest used `<Version>0.0.1</Version>`. Office **discarded it with no
diagnostic** — nothing in the ribbon, nothing in the add-ins menu, nothing in
Office's own runtime log. The only tool that caught it:

```
office-addin-manifest validate
  Error #1: Manifest Version Too Low: The manifest has unsupported version number less than 1.0.
```

I had written that off as an AppSource store nit. It is a hard load requirement.
`1.0.0.0` loads; `0.0.1` does not. **Run the validator and treat every error as
load-blocking**, however much it sounds like store policy.

## Result 4 — RESOLVED: `WEF\Developer` is session-scoped by design (P0.17)

**The add-in was never failing to load. It loads exactly as designed — and the
design is session-scoped.**

A manifest registered under `WEF\Developer` does not get a ribbon entry at
startup. It gets an entry in **Add-ins → Developer Add-ins**, and the
`VersionOverrides` ribbon customization appears only once the add-in is
_inserted_ from that list. Insert it and the group and button appear
immediately; restart Word and they are gone again, because the insertion is not
recorded anywhere that survives the session.

The registry shows why. Every add-in gets a provider bucket under
`HKCU\…\WEF\Providers\<hash>`:

|                    | AppSource add-in                       | our developer sideload |
| ------------------ | -------------------------------------- | ---------------------- |
| `UniqueId`         | `0115236f-…_ADAL`                      | **`developer`**        |
| `AppStates` subkey | present, one entry per add-in + expiry | **absent**             |

`AppStates` is the record that makes an add-in come back after a restart.
Office never writes one for the developer path, so there is nothing to restore
from. This is a property of the mechanism, not a defect in the manifest, the
certificate, the host, or the Office build.

That fully explains the original P0.17 observation — "it ran twice, then
stopped appearing". Those two runs were sessions in which the add-in had been
inserted; every later run was a fresh session where it had not.

### What this means for the installer (P6.14)

Ranked by how well each survives a restart:

1. **AppSource** — the only mechanism that yields a real `AppStates` entry for
   an individual user. Costs a store submission.
2. **Centralized deployment** via the M365 admin center — durable, but requires
   a tenant admin, so it is useless for individual users.
3. **Trusted catalog** — **the UI path for this is gone on Office
   16.0.20228.20190.** A catalog registered at
   `WEF\TrustedCatalogs\{GUID}` with `Flags=1` and a working UNC `Url`
   (`\\localhost\C$\…`, which needs no `New-SmbShare` and no elevation)
   produced no "Shared Folder" tab: this build replaced the classic add-ins
   dialog with the new **Apps** store UI, which does not surface trusted
   catalogs at all. **The admin/share question that was blocking P6.14 is moot
   — the whole path is.**
4. **`WEF\Developer`** — what ships today. One click per session from
   Home → Add-ins → OpenOfficeLLM.

**Open product decision:** accept the once-per-session click, or take the
AppSource route. Nothing in between works for a non-admin individual user on
this Office build.

### Ruled out

- Spike server down at Word start — verified `200` immediately before and every
  15 s during a cold start
- Port contention — real and it did confound two earlier tests (see Result 5),
  but eliminated and the symptom persists
- Registry key deleted or malformed — verified intact after every failure
- Manifest invalid — `office-addin-manifest validate` reports "The manifest is
  valid" in the current state
- Crash-resiliency disabling it — `Word\Resiliency\DisabledItems` holds only
  "Microsoft Word previewer"; no `CrashingAddinList` key exists
- Stale WEF cache — cleared `AddinInfo`, `AppCommands`, `AggregatedCache`
- A bogus trusted catalog poisoning the load — registered a catalog with a local
  folder path, later removed it entirely; no change either way
- **Office add-ins broken globally** — the AppSource Claude add-in loads normally
  in the same Word session, confirmed via UI Automation
- Force-killing Word breaking Office state — plausible, and `Stop-Process -Force`
  was used many times, but the store add-in survives it intact

### Also learned

- **Word's Start screen initialises no add-ins.** A launch with no document
  produces zero runtime-log lines. Always test with a real document:
  `WINWORD.EXE <file.docx>`.
- **COM-automation Word loads no web add-ins.** A Word created via
  `New-Object -ComObject Word.Application` never parses a manifest. It also
  silently creates a _second_ Word instance rather than attaching.
- Office does persist the manifest — it lands at
  `Wef\{DD9E46E2-…}\vai+V16+KDygsa7v22t_6Q==\Manifests\<Id>_<Version>` and is
  rewritten on **every** launch — but never generates a matching
  `Wef\AppCommands` entry, while both AppSource add-ins on this machine have
  one. That asymmetry is the fingerprint of the session-scoped path described
  above, not evidence of a parse failure.
- `Wef\AppCommands\18.0\Word.RibbonCache.en-US` is a real ribbon cache, but it
  is **not** the cause: deleting it changes nothing. It contains only the
  AppSource add-in ids (`WA200010453`), and developer add-ins never appear in
  it.
- Correlation is not mechanism: the add-in appeared on one restart where a
  second manifest happened to be registered alongside it, which looked causal.
  Re-registering that second manifest did **not** reproduce it. The real
  variable was whether the add-in had been inserted in that session.

## Result 5 — two services on one port fail invisibly

Several tests were confounded by orphaned `node` processes fighting over 7317.
The loser died with `EADDRINUSE` while a stale instance kept the port, so the
server appeared "up" while serving nothing usable, and TLS validation failed with
a misleading `Could not establish trust relationship`.

**This validates P1.6 (single-instance lock) as essential, not nice-to-have.**
The host service must detect an existing instance and refuse to start silently,
and `--diagnose` (P6.21) must report who owns the port. Check
`Get-NetTCPConnection -LocalPort <p> -State Listen` before trusting a health
probe, and never treat "something is listening" as "my service is running".

## Tooling that paid off

**UI Automation is the only reliable oracle for "is the add-in actually there".**
Network traffic is not — Office caches icons, so absence of icon fetches does not
mean the add-in failed to load. This inverted two of my conclusions:

```powershell
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$w = Get-Process WINWORD | Select-Object -First 1
$root = [System.Windows.Automation.AutomationElement]::RootElement
$cond = New-Object System.Windows.Automation.PropertyCondition(
  [System.Windows.Automation.AutomationElement]::ProcessIdProperty, $w.Id)
$win = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $cond)
$all = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants,
  (New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::IsControlElementProperty, $true)))
$all | ForEach-Object { $_.Current.Name } | Where-Object { $_ -match 'LNA Spike' }
```

**Office runtime logging** — the only view into manifest handling:

```
HKCU\Software\Microsoft\Office\16.0\WEF\Developer\RuntimeLogging
  (Default) = <repo>\spike\office-runtime.log
```

It logs `Add-in manifest parsing starting.` and never logs completion or error,
so **absence of output is the signal**, not its content.

**WebView2 devtools** — set the env var in the launching process only, never
user-wide (it would expose a debug port for every WebView2 app including Office
and the Claude desktop app):

```powershell
$env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = '--remote-debugging-port=9222'
Start-Process 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE' -ArgumentList '"...docx"'
# targets at http://127.0.0.1:9222/json/list — only once a pane is actually open
```

**Self-reporting page.** The spike page POSTs its findings to `/report`, which
the server prints to stdout. This is how the Gate 0 result was captured — no
devtools attached, nothing for a human to read back. Worth keeping in the real
product as a diagnostics dump.

## Reproducing

```bash
cd spike && npm install && node gen-icons.mjs && node gen-certs.mjs
```

Then `.\register.ps1` (answer **Yes** to the Security Warning), start the server
**detached** so it outlives the shell, and open `spike-test.docx` in Word:

```bash
cd spike && node server.mjs
```

Undo with `.\unregister.ps1`.

## Plan changes this forces

> Task IDs refer to the original plan; the current remaining-work list lives
> in [TODO.md](TODO.md).

| Task        | Change                                                                                              |
| ----------- | --------------------------------------------------------------------------------------------------- |
| P1.6        | Single-instance lock is essential — port contention fails invisibly and misleadingly                |
| P1.8, P6.13 | `Import-Certificate` → `X509Store`; the trust dialog is unavoidable and must be explained first     |
| P2.5, P2.10 | Dial `127.0.0.1` literally, never `localhost` — ~60× slower via IPv6                                |
| P6.1        | `<Version>` must be ≥ 1.0.0.0                                                                       |
| P6.6        | Validator errors are load-blocking; run in CI and fail the build                                    |
| P6.14       | Trusted catalog ruled out — the UI path is gone on this Office build; `WEF\Developer` ships (P6.15) |
| P6.21       | `--diagnose` must report who owns the port, not just whether it responds                            |
| P7.23       | Add: add-in vanishes after restart; Start screen loads no add-ins; never force-kill Office          |
