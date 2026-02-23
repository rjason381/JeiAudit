param(
    [string]$Configuration = "Release",
    [string]$RevitYear = "2024",
    [string]$RevitApiDir = "C:\Program Files\Autodesk\Revit 2024"
)

$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectPath = Join-Path $scriptRoot "..\src\JeiAudit\JeiAudit.csproj"
$projectPath = (Resolve-Path $projectPath).Path
$projectDir = Split-Path -Parent $projectPath

Write-Host "Building JeiAudit..." -ForegroundColor Cyan
dotnet build $projectPath -c $Configuration -p:RevitApiDir="$RevitApiDir"
if ($LASTEXITCODE -ne 0) {
    throw "Build failed. Check compiler output above."
}

$buildOutputDir = Join-Path $projectDir "bin\$Configuration"
$dllPath = Join-Path $buildOutputDir "JeiAudit.dll"
if (-not (Test-Path $dllPath)) {
    throw "Build succeeded but DLL not found at: $dllPath"
}

$addinRoot = Join-Path $env:APPDATA "Autodesk\Revit\Addins\$RevitYear"
$pluginDir = Join-Path $addinRoot "JeiAudit"
$addinFile = Join-Path $addinRoot "JeiAudit.addin"

New-Item -ItemType Directory -Path $pluginDir -Force | Out-Null
$assemblyPath = Join-Path $pluginDir "JeiAudit.dll"
$copiedToAddinsFolder = $true
try {
    Get-ChildItem -Path $buildOutputDir -File | ForEach-Object {
        Copy-Item $_.FullName (Join-Path $pluginDir $_.Name) -Force
    }
}
catch {
    $copiedToAddinsFolder = $false
    if (-not (Test-Path $assemblyPath)) {
        throw "Could not copy binaries to $pluginDir and no previous add-in DLL exists. Close Revit and run again."
    }
    Write-Warning "Could not overwrite $assemblyPath (likely Revit is open)."
    Write-Warning "Manifest will remain pointed to the add-ins folder. Restart Revit and reinstall to apply latest DLL."
}

$addinContent = @"
<?xml version="1.0" encoding="utf-8" standalone="no"?>
<RevitAddIns>
  <AddIn Type="Application">
    <Name>JeiAudit</Name>
    <Assembly>$assemblyPath</Assembly>
    <AddInId>9D5D0A6E-5FEA-4E44-B9FD-B3093708D7B5</AddInId>
    <FullClassName>JeiAudit.App</FullClassName>
    <VendorId>JEIA</VendorId>
    <VendorDescription>JeiAudit - Revit parameter existence auditor</VendorDescription>
  </AddIn>
</RevitAddIns>
"@

Set-Content -Path $addinFile -Value $addinContent -Encoding UTF8

Write-Host ""
Write-Host "JeiAudit installed." -ForegroundColor Green
Write-Host "Add-in folder: $pluginDir"
Write-Host "Manifest: $addinFile"
if (-not $copiedToAddinsFolder) {
    Write-Host "Using existing DLL in add-ins folder: $assemblyPath"
}
Write-Host ""
Write-Host "Restart Revit 2024 to load the plugin."
