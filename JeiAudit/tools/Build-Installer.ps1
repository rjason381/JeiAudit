param(
    [string]$Configuration = "Release",
    [string]$RevitYear = "2024",
    [string]$RevitApiDir = "C:\Program Files\Autodesk\Revit 2024"
)

$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = (Resolve-Path (Join-Path $scriptRoot "..\..")).Path
$projectPath = (Resolve-Path (Join-Path $repoRoot "JeiAudit\src\JeiAudit\JeiAudit.csproj")).Path
$issPath = (Resolve-Path (Join-Path $repoRoot "JeiAudit\installer\JeiAudit-Installer.iss")).Path
$outDir = Join-Path $repoRoot "artifacts\installer"

$isccCandidates = @(
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    "C:\Program Files\Inno Setup 6\ISCC.exe"
)
$iscc = $isccCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) {
    throw "No se encontro ISCC.exe (Inno Setup 6). Instala Inno Setup para generar el .exe."
}

Write-Host "Building JeiAudit ($Configuration)..." -ForegroundColor Cyan
dotnet build $projectPath -c $Configuration -p:RevitApiDir="$RevitApiDir"
if ($LASTEXITCODE -ne 0) {
    throw "Build failed. Revisa el log de compilacion."
}

New-Item -Path $outDir -ItemType Directory -Force | Out-Null

Write-Host "Compiling installer with Inno Setup..." -ForegroundColor Cyan
& $iscc "/DMyConfiguration=$Configuration" "/DMyRevitYear=$RevitYear" $issPath
if ($LASTEXITCODE -ne 0) {
    throw "ISCC fallo al generar el instalador."
}

$setupPath = Join-Path $outDir "JeiAudit_Setup_R$RevitYear.exe"
if (-not (Test-Path $setupPath)) {
    throw "No se encontro el instalador generado en: $setupPath"
}

Write-Host ""
Write-Host "Instalador generado correctamente:" -ForegroundColor Green
Write-Host $setupPath

