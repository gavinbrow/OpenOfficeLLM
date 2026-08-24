# Phase 0 spike registration. Everything here is per-user by design — if any
# step asks for admin, that is itself a finding worth recording, because the
# real installer has the same constraint.

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot

Write-Host "== 1. Trust the spike CA (CurrentUser\Root, no admin) =="
$caPath = Join-Path $root 'certs\ca.crt'
if (-not (Test-Path $caPath)) { throw "Missing $caPath - run: node gen-certs.mjs" }

# NOTE: Import-Certificate is the obvious cmdlet here and it is the wrong one.
# It routes through CryptUIWizImport, which pops the "You are about to install a
# certificate from a certification authority claiming to represent..." dialog
# for the Root store, and fails outright ("UI is not allowed in this operation")
# when no interactive desktop is available. The X509Store API writes the same
# store through CertAddCertificateContextToStore with no UI at all.
$thumbprint = (Get-Content (Join-Path $root 'certs\ca.thumbprint.txt')).Trim()
$existing = Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Thumbprint -eq $thumbprint }
if ($existing) {
  Write-Host "   already trusted: $thumbprint"
} else {
  $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $caPath
  $store = New-Object System.Security.Cryptography.X509Certificates.X509Store 'Root', 'CurrentUser'
  $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
  try { $store.Add($cert) } finally { $store.Close() }
  Write-Host "   imported: $($cert.Thumbprint)"
  if ($cert.Thumbprint -ne $thumbprint) {
    Write-Warning "   thumbprint mismatch (expected $thumbprint)"
  }
}

Write-Host "== 2. Register the manifest (HKCU WEF\Developer, no admin) =="
$manifest = Join-Path $root 'manifest.xml'
if (-not (Test-Path $manifest)) { throw "Missing $manifest" }

$key = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
if (-not (Test-Path $key)) {
  New-Item -Path $key -Force | Out-Null
  Write-Host "   created $key"
}
New-ItemProperty -Path $key -Name $manifest -Value $manifest -PropertyType String -Force | Out-Null
Write-Host "   registered $manifest"

Write-Host ""
Write-Host "Done. Next:"
Write-Host "  1. node server.mjs      (leave running)"
Write-Host "  2. Open Word, then Excel"
Write-Host "  3. Home > Add-ins > (More Add-ins) > Developer Add-ins > OpenOfficeLLM LNA Spike"
