# Architecture Decision Records

A lightweight log of the decisions that shape OpenOfficeLLM and the rationale
behind them. Numbered ADRs; one section per decision. Append-only — supersede
don't delete.

---

## ADR-001 — Loopback origin as the add-in host (the LNA escape hatch)

**Date:** 2026-08-16
**Status:** Accepted. Verified by the Phase 0 spike — GATE 0 passed for Word
**and** Excel on Office WebView2 151 (2026-08-17); see
[`SPIKE-LNA.md`](SPIKE-LNA.md). The architecture stands on the
loopback→loopback exemption; if it ever fails inside a future WebView2, the
project pivots to a cloud-only add-in.

**Context.** Chromium 142+ enforces Local Network Access (LNA): an HTTPS page
served from a public origin cannot call `http://127.0.0.1`. Office embeds
add-ins in iframes without the `local-network-access` permissions policy, so
Chrome cannot even prompt — the request just fails. There is no consumer-
viable opt-out (the enterprise
`LocalNetworkAccessRestrictionsTemporaryOptOut` policy is unacceptable for a
consumer install). Microsoft has it in backlog with no timeline
([#6281](https://github.com/OfficeDev/office-js/issues/6281),
[#6366](https://github.com/OfficeDev/office-js/issues/6366)).

**Decision.** Serve the task pane HTML from `https://127.0.0.1:7317` via a
local host service. Loopback→loopback requests are exempt from LNA, so the
pane's calls to Ollama/LM Studio/llama.cpp succeed without any opt-out. As a
side effect, the host service becomes the natural broker for every model
call — solving CORS, mixed content, and API-key storage. Keys live in
Windows DPAPI and never enter the webview.

**Consequences.**

- A local service must be running for the add-in to work at all → autostart at
  logon, plus a recovery screen in the pane (`ServiceDownScreen`).
- The manifest's `<SourceLocation>` is hardcoded to a loopback origin;
  port conflicts handled by templating the port at install time.
- A self-signed CA must be trusted in `Cert:\CurrentUser\Root` — no admin
  required, but Windows always shows an unavoidable Security Warning dialog.
- Distribution is **not** AppSource — it's an open-source installer.

---

## ADR-002 — XML manifest, not unified JSON manifest

**Date:** 2026-08-16
**Status:** Accepted.

**Context.** The unified (JSON) manifest format reached GA for Word/Excel in
July 2026, but the unified format is **not supported for local / shared-folder
deployment** — the exact distribution path for a sideloaded, locally-hosted
add-in. The XML format is supported for both AppSource and sideloading.

**Decision.** Ship an XML manifest, **generated at runtime** by
`packages/host/src/manifest.ts` and written to
`%APPDATA%\OpenOfficeLLM\manifest\openofficellm.xml`. One manifest
covers both Word (`<Host Name="Document"/>`) and Excel
(`<Host Name="Workbook"/>`).

**Consequences.**

- Verbose XML to maintain, but tooling (`office-addin-manifest validate`) is
  mature.
- Validator errors must be treated as load-blocking — the Phase 0 spike
  found that a `<Version>` below `1.0.0.0` is silently discarded by Office
  with no diagnostic anywhere; only the validator catches it.

---

## ADR-003 — Trusted catalog as the primary registration path

**Date:** 2026-08-16
**Status:** **Superseded by ADR-004** (2026-08-17). The trusted-catalog path
was ruled out on Office 16.0.20228.20190 — see
[`SPIKE-LNA.md`](SPIKE-LNA.md), Result 4. This ADR is
retained for provenance.

**Context.** The original plan called for `HKCU\...\WEF\Developer\<path>`
registration as primary, with the trusted-catalog approach as a fallback. The
spike showed `WEF\Developer` registration is **not durable across Word
restarts**: Office reads the manifest once (when the registry value or
manifest mtime changes), caches it, but never generates a `Wef\AppCommands`
entry, so the ribbon button disappears on the next launch.

**Decision (original).** Make the trusted catalog the primary registration:
`HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{GUID}` with
`Id`, `Url`, `Flags=dword:00000001`. Demote `WEF\Developer` to a
developer-convenience fallback.

**Why it was superseded.** A correctly registered catalog against a working
`\\localhost\C$` UNC (no `New-SmbShare`, no elevation) produced no "Shared
Folder" tab: Office 16.0.20228.20190 replaced the classic add-ins dialog with
the new **Apps** store UI, which does not surface trusted catalogs at all.
The admin/share question that was blocking this path is moot — the whole path
is. ADR-004 records what ships instead.

---

## ADR-004 — `WEF\Developer` registration, one click per session

**Date:** 2026-08-17
**Status:** Accepted. This is what ships today, by elimination rather than
preference.

**Context.** ADR-003's trusted-catalog path is gone on the current Office
build (see SPIKE-LNA Result 4). The remaining options for a non-admin
individual user are `WEF\Developer` sideloading or an AppSource submission.

**Decision.** Ship `WEF\Developer` registration
(`HKCU\Software\Microsoft\Office\16.0\WEF\Developer\<manifestPath>`,
implemented in [`packages/host/src/register.ts`](../packages/host/src/register.ts)).
It is session-scoped by design: Office records the developer path with
`Providers\<hash>\UniqueId = developer` and never writes the `AppStates`
subkey that store-installed add-ins get, so the add-in must be opened from
**Add-ins → Developer Add-ins** once per Word/Excel session.

**Consequences.**

- One click per session from Home → Add-ins → OpenOfficeLLM. Zero install
  cost, ships today, but is not the "seamless" bar the project originally set.
- `--install` is idempotent and doubles as `--repair`.
- The durable-install product decision is **resolved: accept one click per
  session.** AppSource is incompatible with the loopback-origin architecture
  (see the deferred-to-v2+ table in [TODO.md](TODO.md)); the installer wraps
  `WEF\Developer` sideloading and re-provisions on every upgrade via the
  idempotent `--install` command. P0.19 is closed.
