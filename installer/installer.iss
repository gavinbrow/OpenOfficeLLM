; OpenOfficeLLM installer — Inno Setup script.
;
; Per-user, no admin. Installs to %LOCALAPPDATA%\Programs\OpenOfficeLLM and
; runs `host.exe --install` at the end, which trusts the local CA, writes the
; add-in manifest, registers it with Office, and sets up autostart.
;
; Build with:
;   ISCC /DAPP_VERSION=0.1.0 installer/installer.iss
;
; Or via the orchestrator:
;   node scripts/build-installer.mjs
;
; AppId is a stable GUID — never change it. It tells Inno Setup that a new
; version is an upgrade, not a side-by-side install. Changing it would orphan
; every existing install.

#define AppId "{{B7F5A2C1-3D84-4E6F-9A12-5C8E0D4B7F93}"

; APP_VERSION is passed via /D on the command line (see build-installer.mjs).
; Fallback to a hardcoded value so the script can be opened in ISCC's GUI.
#ifndef APP_VERSION
  #define APP_VERSION "0.1.0"
#endif

[Setup]
AppId={#AppId}
AppName=OpenOfficeLLM
AppVersion={#APP_VERSION}
AppPublisher=OpenOfficeLLM
AppPublisherURL=https://github.com/openofficellm/OpenOfficeLLM
AppSupportURL=https://github.com/openofficellm/OpenOfficeLLM/issues
DefaultDirName={localappdata}\Programs\OpenOfficeLLM
DefaultGroupName=OpenOfficeLLM
; Per-user install — no UAC prompt, no admin rights at any point.
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
; Compression
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
DisableProgramGroupPage=yes
DisableDirPage=no
; On upgrade: stop the running host, replace files, then re-provision.
CloseApplications=force
RestartApplications=no
; Uninstaller reverses the provisioning (host.exe --uninstall) before deleting
; files — see [UninstallRun] below.
Uninstallable=yes
; Versioning
OutputBaseFilename=OpenOfficeLLM-Setup-{#APP_VERSION}
OutputDir=..\dist
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Files]
; The SEA binary — self-contained, no Node needed.
Source: "staging\host.exe"; DestDir: "{app}"; Flags: ignoreversion
; The add-in web bundle — served by the host at https://127.0.0.1:7317/
Source: "staging\web\*"; DestDir: "{app}\web"; Flags: recursesubdirs ignoreversion
; Version file — read by /api/health and the update-check flow.
Source: "staging\version.txt"; DestDir: "{app}"; Flags: ignoreversion

[Run]
; Provisioning: trust CA, write manifest, register with Office, set autostart.
; Idempotent, so this is the same command for fresh install and upgrade — the
; host detects what's already done and skips it.
Filename: "{app}\host.exe"; Parameters: "--install"; Flags: runhidden; StatusMsg: "Setting up the add-in..."
; Start the host now. --install only *enables* autostart, which does not fire
; until the next logon; until something is actually listening on 7317 Office
; has nowhere to fetch the pane from and Word reports that it cannot load the
; add-in. Installing and then opening Word — the obvious thing to do — hit
; exactly that window. acquireLock() makes a redundant instance exit cleanly,
; so this is safe even if one is somehow already running.
Filename: "{app}\host.exe"; Flags: nowait runhidden; StatusMsg: "Starting OpenOfficeLLM..."

[Icons]
; Start menu shortcut — launches the host (which serves the pane and brokers calls).
Name: "{group}\OpenOfficeLLM"; Filename: "{app}\host.exe"
Name: "{group}\Uninstall OpenOfficeLLM"; Filename: "{uninstallexe}"

[UninstallRun]
; Reverse the provisioning before removing files: unregister from Office,
; remove the autostart entry, untrust the CA. Config, secrets and chat history
; in %APPDATA%\OpenOfficeLLM are left in place (the host's --uninstall prints
; where they are so the user can delete them manually if desired).
Filename: "{app}\host.exe"; Parameters: "--uninstall"; Flags: runhidden; RunOnceId: "UninstallHost"

[UninstallDelete]
; Clean up the install directory entirely (Inno Setup only removes files it
; tracked, so this catches anything the host wrote at runtime like logs).
Type: filesandordirs; Name: "{app}"