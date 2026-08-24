// Generates the Office add-in manifest.
//
// The manifest is generated rather than shipped as a static file because every
// URL inside it is absolute and carries the port. The host picks its port at
// startup (config, --port, or a scan when the preferred one is taken), so a
// checked-in manifest would silently point at the wrong origin the first time
// anything else held 7317 — the pane would fail to load with no clue why.
// Writing it from the port we actually bound makes that class of bug
// impossible. See P6.5/P6.12/P6.20.

import fs from 'node:fs'
import path from 'node:path'
import { resolveManifestDir, ensureDirs, HOST_INTERFACE } from './paths.js'
import { logger } from './logging.js'

/** Stable across versions and installs. Office keys its per-add-in state off
 *  this GUID, so changing it orphans the previous registration and reappears
 *  as a second entry in the add-ins list. Never reuse the spike's GUID
 *  (7e3d0a41-…) — the two must be able to coexist during development. */
export const ADDIN_ID = 'b7f5a2c1-3d84-4e6f-9a12-5c8e0d4b7f93'

/** Office silently declines to surface a manifest whose version is below 1.0 —
 *  no error in the ribbon, the add-ins dialog, or Office's own runtime log.
 *  This cost a day in P0.2; the only tool that reported it was
 *  `office-addin-manifest validate`. Keep the leading component >= 1. */
export const MANIFEST_VERSION = '1.0.0.0'

export const MANIFEST_FILENAME = 'openofficellm.xml'

function xmlEscape(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

/**
 * Ribbon markup for one host. Word and Excel need separate <Host> blocks that
 * differ only in xsi:type, but every id inside is shared so the two surfaces
 * stay identical.
 */
function hostBlock(xsiType: 'Document' | 'Workbook'): string {
  return `      <Host xsi:type="${xsiType}">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="Taskpane.Url"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="OpenOfficeLLMGroup">
                <Label resid="Group.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="OpenOfficeLLMTaskpaneButton">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>OpenOfficeLLMTaskpane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>`
}

export interface ManifestOptions {
  port: number
  /** Override the origin host. Only the loopback literal is supported in
   *  practice — see the AppDomains note below. */
  hostname?: string
}

export function renderManifest(opts: ManifestOptions): string {
  const hostname = opts.hostname ?? HOST_INTERFACE
  const origin = xmlEscape(`https://${hostname}:${opts.port}`)

  // Element order is load-bearing: OfficeApp's schema uses xsd:sequence, so a
  // reordered element is a hard validation failure, not a warning.
  return `<?xml version="1.0" encoding="UTF-8"?>
<!--
  GENERATED FILE — do not edit.

  Written by the host service on every startup from the port it actually bound
  (see packages/host/src/manifest.ts). Local edits are overwritten. To change
  the ribbon or the add-in identity, edit manifest.ts and restart the host.
-->
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">

  <Id>${ADDIN_ID}</Id>
  <Version>${MANIFEST_VERSION}</Version>
  <ProviderName>OpenOfficeLLM</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="OpenOfficeLLM"/>
  <Description DefaultValue="AI assistant for Word and Excel, with first-class local model support."/>
  <IconUrl DefaultValue="${origin}/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="${origin}/icon-80.png"/>
  <SupportUrl DefaultValue="${origin}/index.html"/>

  <!-- Only the loopback origin. Any host the pane may navigate to must be
       listed, and we deliberately never navigate off-origin: keeping the pane
       on the loopback origin is what puts it in the same address space as
       Ollama and LM Studio, which is the entire reason Local Network Access
       does not block us. -->
  <AppDomains>
    <AppDomain>${origin}</AppDomain>
  </AppDomains>

  <Hosts>
    <Host Name="Document"/>
    <Host Name="Workbook"/>
  </Hosts>

  <!-- No top-level <Requirements>. One manifest covers both hosts, so
       demanding WordApi here would make Office refuse to load the add-in in
       Excel (and vice versa). Capability checks belong in the pane at
       runtime. -->

  <DefaultSettings>
    <SourceLocation DefaultValue="${origin}/index.html"/>
  </DefaultSettings>

  <Permissions>ReadWriteDocument</Permissions>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
${hostBlock('Document')}
${hostBlock('Workbook')}
    </Hosts>

    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="${origin}/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="${origin}/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="${origin}/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="${origin}/index.html"/>
        <bt:Url id="Commands.Url" DefaultValue="${origin}/commands.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Group.Label" DefaultValue="OpenOfficeLLM"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="AI Assistant"/>
        <bt:String id="GetStarted.Title" DefaultValue="OpenOfficeLLM is ready"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Open the OpenOfficeLLM chat pane."/>
        <bt:String id="GetStarted.Description" DefaultValue="Open the AI Assistant pane from the Home tab to start chatting with local or cloud models."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`
}

export function manifestPath(): string {
  return path.join(resolveManifestDir(), MANIFEST_FILENAME)
}

export interface WriteManifestResult {
  path: string
  changed: boolean
}

/**
 * Write the manifest for `port`, returning whether the content actually
 * changed.
 *
 * The no-op-write check matters more than it looks. Office re-reads a
 * developer-registered manifest when the file's mtime moves, and a re-read
 * makes it tear down and rebuild the ribbon entry. Rewriting identical bytes
 * on every host restart would mean Office re-parsing the manifest constantly
 * for no reason.
 */
export function writeManifest(opts: ManifestOptions): WriteManifestResult {
  ensureDirs()
  const target = manifestPath()
  const next = renderManifest(opts)
  let current: string | null = null
  try {
    current = fs.readFileSync(target, 'utf8')
  } catch {
    current = null
  }
  if (current === next) {
    return { path: target, changed: false }
  }
  fs.writeFileSync(target, next, 'utf8')
  logger.info({
    msg: current === null ? 'manifest written' : 'manifest updated (port or content changed)',
    path: target,
    port: opts.port,
  })
  return { path: target, changed: true }
}
