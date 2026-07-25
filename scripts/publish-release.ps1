# Build release zips for GitHub Releases.
# self-contained: only PptxAvalonia.exe + PptxAvalonia.pdb + demo.pptx
# framework-dependent: single-file host + native deps + pdb + demo.pptx

param(
    [string]$Version = "v1.0.0"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

$artifacts = Join-Path $root "artifacts"
Remove-Item -Recurse -Force $artifacts -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path `
    "$artifacts\sc-raw", "$artifacts\fd-raw", "$artifacts\sc-pack", "$artifacts\fd-pack" | Out-Null

# --- Self-contained ---
dotnet publish PptxAvalonia.csproj -c Release -r win-x64 --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -p:DebugType=portable `
    -p:CopyOutputSymbolsToPublishDirectory=true `
    -o "$artifacts\sc-raw"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Copy-Item "$artifacts\sc-raw\PptxAvalonia.exe" "$artifacts\sc-pack\"
Copy-Item "$artifacts\sc-raw\PptxAvalonia.pdb" "$artifacts\sc-pack\"
Copy-Item "$root\Samples\demo.pptx" "$artifacts\sc-pack\demo.pptx"

$scZip = "$artifacts\PptxAvalonia-$Version-win-x64-self-contained.zip"
Compress-Archive -Path "$artifacts\sc-pack\*" -DestinationPath $scZip -Force

# --- Framework-dependent ---
dotnet publish PptxAvalonia.csproj -c Release -r win-x64 --self-contained false `
    -p:PublishSingleFile=true `
    -p:DebugType=portable `
    -p:CopyOutputSymbolsToPublishDirectory=true `
    -o "$artifacts\fd-raw"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Get-ChildItem "$artifacts\fd-raw" -File | Where-Object {
    $_.Extension -match '\.(exe|dll|pdb)$'
} | ForEach-Object { Copy-Item $_.FullName "$artifacts\fd-pack\" }
Copy-Item "$root\Samples\demo.pptx" "$artifacts\fd-pack\demo.pptx" -Force

$fdZip = "$artifacts\PptxAvalonia-$Version-win-x64-framework-dependent.zip"
Compress-Archive -Path "$artifacts\fd-pack\*" -DestinationPath $fdZip -Force

Write-Host "Self-contained pack:"
Get-ChildItem "$artifacts\sc-pack" | ForEach-Object { Write-Host "  $($_.Name)" }
Write-Host "Zips:"
Get-ChildItem "$artifacts\*.zip" | ForEach-Object { Write-Host "  $($_.Name) ($([math]::Round($_.Length/1MB,2)) MB)" }
