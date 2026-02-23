param(
    [string]$RevitYear = "2024"
)

$ErrorActionPreference = "Stop"

$addinRoot = Join-Path $env:APPDATA "Autodesk\Revit\Addins\$RevitYear"
$pluginDir = Join-Path $addinRoot "JeiAudit"
$addinFile = Join-Path $addinRoot "JeiAudit.addin"

if (Test-Path $addinFile) {
    Remove-Item $addinFile -Force
    Write-Host "Removed: $addinFile"
}
else {
    Write-Host "Not found: $addinFile"
}

if (Test-Path $pluginDir) {
    Remove-Item $pluginDir -Recurse -Force
    Write-Host "Removed: $pluginDir"
}
else {
    Write-Host "Not found: $pluginDir"
}

Write-Host "JeiAudit uninstall complete."
