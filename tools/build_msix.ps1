param(
    [string]$Executable = "dist\ScanItEasy_GUI.exe",
    [string]$Output = "dist\ScanItEasy-2.1.0-Store.msix"
)
$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$layout = Join-Path $repoRoot "build\msix-layout"
$sourceExe = Join-Path $repoRoot $Executable
$outputPackage = Join-Path $repoRoot $Output
$python = Join-Path $repoRoot "venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $sourceExe)) { throw "Executable not found: $sourceExe" }
if (-not (Test-Path -LiteralPath $python)) { throw "Virtual environment Python not found: $python" }
$resolvedRoot = [IO.Path]::GetFullPath($repoRoot)
$resolvedLayout = [IO.Path]::GetFullPath($layout)
if (-not $resolvedLayout.StartsWith($resolvedRoot, [StringComparison]::OrdinalIgnoreCase)) { throw "Refusing to clean layout outside repository" }
& $python (Join-Path $repoRoot "tools\generate_store_assets.py")
if ($LASTEXITCODE -ne 0) { throw "Store asset generation failed" }
if (Test-Path -LiteralPath $layout) { Remove-Item -LiteralPath $layout -Recurse -Force }
New-Item -ItemType Directory -Path (Join-Path $layout "Assets") -Force | Out-Null
Copy-Item -LiteralPath $sourceExe -Destination (Join-Path $layout "ScanItEasy.exe")
Copy-Item -LiteralPath (Join-Path $repoRoot "packaging\Package.appxmanifest") -Destination (Join-Path $layout "Package.appxmanifest")
Copy-Item -Path (Join-Path $repoRoot "packaging\Assets\*.png") -Destination (Join-Path $layout "Assets")
$winapp = Get-Command winapp.exe -ErrorAction SilentlyContinue
if ($null -eq $winapp) {
    $wingetLink = Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Links\winapp.exe"
    if (-not (Test-Path -LiteralPath $wingetLink)) { throw "winapp.exe not found" }
    $winapp = Get-Item -LiteralPath $wingetLink
}
& $winapp.Source pack $layout --manifest (Join-Path $layout "Package.appxmanifest") --output $outputPackage --skip-pri
if ($LASTEXITCODE -ne 0) { throw "MSIX packaging failed with exit code $LASTEXITCODE" }
Write-Host "Microsoft Store package created: $outputPackage"