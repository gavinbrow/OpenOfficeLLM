# Alternative registration: trusted catalog instead of the Developer key.
#
# WEF\Developer proved non-durable — Office reads the manifest once after the
# registry value changes, then drops the add-in on the next restart (P0.17).
# A trusted catalog is the mechanism Microsoft actually documents for Windows
# desktop sideloading.
#
# Microsoft documents the catalog Url as a network share. Creating a share needs
# admin, which collides with the no-admin requirement, so this tries a plain
# local folder path first. If Office ignores the catalog, that answers the
# question and the share route becomes the fallback.

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot

# Fixed GUID so re-running replaces rather than accumulates catalogs.
$catalogId = '{B2F4A7C1-9E38-4D52-A6B7-3C1E8F0D5A94}'
$catalogDir = Join-Path $root 'catalog'

New-Item -ItemType Directory -Path $catalogDir -Force | Out-Null
Copy-Item (Join-Path $root 'manifest.xml') (Join-Path $catalogDir 'manifest.xml') -Force
Write-Host "catalog folder: $catalogDir"

$base = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs'
if (-not (Test-Path $base)) { New-Item -Path $base -Force | Out-Null }

$key = Join-Path $base $catalogId
if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }

New-ItemProperty -Path $key -Name 'Id'    -Value $catalogId  -PropertyType String -Force | Out-Null
New-ItemProperty -Path $key -Name 'Url'   -Value $catalogDir -PropertyType String -Force | Out-Null
# 1 = enabled and shown in the add-ins menu
New-ItemProperty -Path $key -Name 'Flags' -Value 1           -PropertyType DWord  -Force | Out-Null

Write-Host "registered trusted catalog $catalogId -> $catalogDir"
Get-ItemProperty -Path $key | Format-List Id, Url, Flags

Write-Host ""
Write-Host "Restart Word with a real document (NOT the Start screen), then look in"
Write-Host "  Home > Add-ins > More Add-ins > SHARED FOLDER"
