# Build release zips for GitHub Releases.
# self-contained: only PptxAvalonia.exe + PptxAvalonia.pdb + demo.pptx
# framework-dependent: single-file host + native deps + pdb + demo.pptx
#
# Optional Authenticode signing when certificate env vars are set:
#   CODE_SIGNING_PFX_PATH or CODE_SIGNING_PFX_BASE64 + CODE_SIGNING_PFX_PASSWORD

param(
    [string]$Version = "v1.0.0",
    [switch]$SkipSign
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

$artifacts = Join-Path $root "artifacts"
Remove-Item -Recurse -Force $artifacts -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path `
    (Join-Path $artifacts "sc-raw"),
    (Join-Path $artifacts "fd-raw"),
    (Join-Path $artifacts "sc-pack"),
    (Join-Path $artifacts "fd-pack") | Out-Null

function Invoke-OptionalSign {
    param(
        [string]$TargetPath,
        [switch]$Recurse
    )

    if ($SkipSign) {
        Write-Host "SkipSign: not signing $TargetPath"
        return
    }

    $hasPath = $env:CODE_SIGNING_PFX_PATH -and (Test-Path $env:CODE_SIGNING_PFX_PATH)
    $hasB64 = -not [string]::IsNullOrWhiteSpace($env:CODE_SIGNING_PFX_BASE64)
    if (-not ($hasPath -or $hasB64)) {
        Write-Host "No signing certificate configured; building unsigned packages."
        return
    }

    $signScript = Join-Path $root "scripts\sign.ps1"
    if ($Recurse) {
        & $signScript -Path $TargetPath -Recurse
    } else {
        & $signScript -Path $TargetPath
    }
}

# --- Self-contained ---
dotnet publish PptxAvalonia.csproj -c Release -r win-x64 --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -p:DebugType=portable `
    -p:CopyOutputSymbolsToPublishDirectory=true `
    -o (Join-Path $artifacts "sc-raw")
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Copy-Item (Join-Path $artifacts "sc-raw\PptxAvalonia.exe") (Join-Path $artifacts "sc-pack\")
Copy-Item (Join-Path $artifacts "sc-raw\PptxAvalonia.pdb") (Join-Path $artifacts "sc-pack\")
Copy-Item (Join-Path $root "Samples\demo.pptx") (Join-Path $artifacts "sc-pack\demo.pptx")
Copy-Item (Join-Path $root "scripts\閫?撠?.bat") (Join-Path $artifacts "sc-pack\")
Copy-Item (Join-Path $root "scripts\unblock-after-download.ps1") (Join-Path $artifacts "sc-pack\")
Invoke-OptionalSign -TargetPath (Join-Path $artifacts "sc-pack\PptxAvalonia.exe")

$scZip = Join-Path $artifacts "PptxAvalonia-$Version-win-x64-self-contained.zip"
if (Test-Path $scZip) { Remove-Item $scZip -Force }
Compress-Archive -Path (Join-Path $artifacts "sc-pack\*") -DestinationPath $scZip -Force

# --- Framework-dependent ---
dotnet publish PptxAvalonia.csproj -c Release -r win-x64 --self-contained false `
    -p:PublishSingleFile=true `
    -p:DebugType=portable `
    -p:CopyOutputSymbolsToPublishDirectory=true `
    -o (Join-Path $artifacts "fd-raw")
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Get-ChildItem (Join-Path $artifacts "fd-raw") -File | Where-Object {
    $_.Extension -match '\.(exe|dll|pdb)$'
} | ForEach-Object {
    Copy-Item $_.FullName (Join-Path $artifacts "fd-pack\")
}
Copy-Item (Join-Path $root "Samples\demo.pptx") (Join-Path $artifacts "fd-pack\demo.pptx") -Force
Copy-Item (Join-Path $root "scripts\閫?撠?.bat") (Join-Path $artifacts "fd-pack\")
Copy-Item (Join-Path $root "scripts\unblock-after-download.ps1") (Join-Path $artifacts "fd-pack\")
Invoke-OptionalSign -TargetPath (Join-Path $artifacts "fd-pack\PptxAvalonia.exe")

$fdZip = Join-Path $artifacts "PptxAvalonia-$Version-win-x64-framework-dependent.zip"
if (Test-Path $fdZip) { Remove-Item $fdZip -Force }
Compress-Archive -Path (Join-Path $artifacts "fd-pack\*") -DestinationPath $fdZip -Force

# Signature report
$signedNote = "unsigned"
$exe = Join-Path $artifacts "sc-pack\PptxAvalonia.exe"
if (Test-Path $exe) {
    try {
        $sig = Get-AuthenticodeSignature $exe
        if ($sig.Status -eq "Valid") {
            $signedNote = "signed ($($sig.SignerCertificate.Subject))"
        } else {
            $signedNote = "signature-status=$($sig.Status)"
        }
    } catch {
        $signedNote = "signature-check-failed"
    }
}

$builtUtc = [DateTime]::UtcNow.ToString("yyyy-MM-dd HH:mm:ss")
$reportLines = @(
    "# PptxAvalonia $Version package report",
    "",
    "Self-contained: $scZip",
    "Framework-dependent: $fdZip",
    "Authenticode: $signedNote",
    "Built: $builtUtc UTC"
)
Set-Content -Path (Join-Path $artifacts "PACKAGE_REPORT.md") -Value $reportLines -Encoding UTF8

Write-Host "Self-contained pack:"
Get-ChildItem (Join-Path $artifacts "sc-pack") | ForEach-Object { Write-Host ("  " + $_.Name) }
Write-Host "Zips:"
Get-ChildItem (Join-Path $artifacts "*.zip") | ForEach-Object {
    $mb = [math]::Round($_.Length / 1MB, 2)
    Write-Host ("  " + $_.Name + " (" + $mb + " MB)")
}
Write-Host ("Authenticode: " + $signedNote)

