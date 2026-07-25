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
    "$artifacts\sc-raw", "$artifacts\fd-raw", "$artifacts\sc-pack", "$artifacts\fd-pack" | Out-Null

function Invoke-OptionalSign([string]$TargetPath, [switch]$Recurse) {
    if ($SkipSign) {
        Write-Host "SkipSign: 不簽署 $TargetPath"
        return
    }

    $hasCert = ($env:CODE_SIGNING_PFX_PATH -and (Test-Path $env:CODE_SIGNING_PFX_PATH)) -or
               (-not [string]::IsNullOrWhiteSpace($env:CODE_SIGNING_PFX_BASE64))
    if (-not $hasCert) {
        Write-Host "未設定簽章憑證，產出未簽署套件（可能被智慧型應用程式控制封鎖）。"
        return
    }

    $args = @{ Path = $TargetPath }
    if ($Recurse) { $args.Recurse = $true }
    & "$root\scripts\sign.ps1" @args
    if ($LASTEXITCODE -ne 0 -and $LASTEXITCODE -ne $null) {
        # sign.ps1 exits 0 when skipping; non-zero is real failure
        if ($LASTEXITCODE -gt 0) { exit $LASTEXITCODE }
    }
}

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
Invoke-OptionalSign "$artifacts\sc-pack\PptxAvalonia.exe"

$scZip = "$artifacts\PptxAvalonia-$Version-win-x64-self-contained.zip"
if (Test-Path $scZip) { Remove-Item $scZip -Force }
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
Invoke-OptionalSign "$artifacts\fd-pack\PptxAvalonia.exe"

$fdZip = "$artifacts\PptxAvalonia-$Version-win-x64-framework-dependent.zip"
if (Test-Path $fdZip) { Remove-Item $fdZip -Force }
Compress-Archive -Path "$artifacts\fd-pack\*" -DestinationPath $fdZip -Force

# Signature report
$signedNote = "unsigned"
$exe = "$artifacts\sc-pack\PptxAvalonia.exe"
if (Test-Path $exe) {
    try {
        $sig = Get-AuthenticodeSignature $exe
        if ($sig.Status -eq "Valid") { $signedNote = "signed ($($sig.SignerCertificate.Subject))" }
        else { $signedNote = "signature-status=$($sig.Status)" }
    } catch {
        $signedNote = "signature-check-failed"
    }
}

@"
# PptxAvalonia $Version package report

- Self-contained: $scZip
- Framework-dependent: $fdZip
- Authenticode: $signedNote
- Built: $([DateTime]::UtcNow.ToString("yyyy-MM-dd HH:mm:ss")) UTC
"@ | Set-Content "$artifacts\PACKAGE_REPORT.md" -Encoding UTF8

Write-Host "Self-contained pack:"
Get-ChildItem "$artifacts\sc-pack" | ForEach-Object { Write-Host "  $($_.Name)" }
Write-Host "Zips:"
Get-ChildItem "$artifacts\*.zip" | ForEach-Object { Write-Host "  $($_.Name) ($([math]::Round($_.Length/1MB,2)) MB)" }
Write-Host "Authenticode: $signedNote"
