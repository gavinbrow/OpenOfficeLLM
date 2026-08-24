# Reverses register.ps1. Removes the CA by thumbprint, never by subject name —
# matching on subject would happily delete somebody else's cert.

$ErrorActionPreference = 'Continue'
$root = $PSScriptRoot

$key = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
$manifest = Join-Path $root 'manifest.xml'
if (Test-Path $key) {
  Remove-ItemProperty -Path $key -Name $manifest -ErrorAction SilentlyContinue
  Write-Host "unregistered $manifest"
}

$tpFile = Join-Path $root 'certs\ca.thumbprint.txt'
if (Test-Path $tpFile) {
  $thumbprint = (Get-Content $tpFile).Trim()
  Get-ChildItem Cert:\CurrentUser\Root |
    Where-Object { $_.Thumbprint -eq $thumbprint } |
    ForEach-Object {
      Remove-Item -Path $_.PSPath -Force
      Write-Host "removed CA $($_.Thumbprint)"
    }
}

Write-Host "Also clear the Office web cache if the pane still appears:"
Write-Host "  %LOCALAPPDATA%\Microsoft\Office\16.0\Wef"
