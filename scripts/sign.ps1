# Authenticode-sign binaries with a PFX certificate.
# Usage:
#   .\scripts\sign.ps1 -Path artifacts\sc-pack\PptxAvalonia.exe
#   .\scripts\sign.ps1 -Path artifacts\fd-pack -Recurse
#
# Certificate sources (first match wins):
#   1) -PfxPath + -PfxPassword
#   2) env CODE_SIGNING_PFX_PATH + CODE_SIGNING_PFX_PASSWORD
#   3) env CODE_SIGNING_PFX_BASE64 + CODE_SIGNING_PFX_PASSWORD  (CI)

param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [switch]$Recurse,

    [string]$PfxPath = $env:CODE_SIGNING_PFX_PATH,
    [string]$PfxPassword = $env:CODE_SIGNING_PFX_PASSWORD,
    [string]$PfxBase64 = $env:CODE_SIGNING_PFX_BASE64,

    [string]$TimestampUrl = "http://timestamp.digicert.com",
    [string]$Description = "PptxAvalonia"
)

$ErrorActionPreference = "Stop"

function Find-SignTool {
    $cmd = Get-Command signtool.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    $roots = @(
        "${env:ProgramFiles(x86)}\Windows Kits\10\bin",
        "${env:ProgramFiles}\Windows Kits\10\bin"
    )
    foreach ($root in $roots) {
        if (-not (Test-Path $root)) { continue }
        $found = Get-ChildItem -Path $root -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -match '\\x64\\signtool\.exe$' } |
            Select-Object -First 1
        if ($found) { return $found.FullName }
    }
    return $null
}

function Resolve-PfxFile {
    if ($PfxPath -and (Test-Path $PfxPath)) {
        return (Resolve-Path $PfxPath).Path
    }

    if (-not [string]::IsNullOrWhiteSpace($PfxBase64)) {
        $tmp = Join-Path $env:TEMP ("pptxavalonia-codesign-" + [guid]::NewGuid().ToString("N") + ".pfx")
        [IO.File]::WriteAllBytes($tmp, [Convert]::FromBase64String($PfxBase64))
        return $tmp
    }

    return $null
}

$signTool = Find-SignTool
if (-not $signTool) {
    Write-Error "找不到 signtool.exe。請安裝 Windows SDK（Signing Tools for Desktop Apps）。"
}

$pfx = Resolve-PfxFile
if (-not $pfx) {
    Write-Host "未提供簽章憑證，略過簽署。"
    Write-Host "請設定 CODE_SIGNING_PFX_PATH 或 CODE_SIGNING_PFX_BASE64 + CODE_SIGNING_PFX_PASSWORD。"
    exit 0
}

if ([string]::IsNullOrEmpty($PfxPassword)) {
    Write-Error "缺少 CODE_SIGNING_PFX_PASSWORD / -PfxPassword。"
}

$targets = @()
if (Test-Path $Path -PathType Container) {
    $filter = if ($Recurse) { Get-ChildItem $Path -Recurse -File } else { Get-ChildItem $Path -File }
    $targets = $filter | Where-Object { $_.Extension -match '\.(exe|dll)$' -and $_.Name -notmatch '^(libSkia|libHarfBuzz|av_libgles)' }
} elseif (Test-Path $Path -PathType Leaf) {
    $targets = @(Get-Item $Path)
} else {
    Write-Error "路徑不存在: $Path"
}

if ($targets.Count -eq 0) {
    Write-Host "沒有可簽署的檔案。"
    exit 0
}

$pass = $PfxPassword
foreach ($file in $targets) {
    Write-Host "Signing $($file.FullName) ..."
    & $signTool sign `
        /fd SHA256 `
        /td SHA256 `
        /tr $TimestampUrl `
        /f $pfx `
        /p $pass `
        /d $Description `
        $file.FullName
    if ($LASTEXITCODE -ne 0) {
        Write-Error "簽署失敗: $($file.FullName) (exit $LASTEXITCODE)"
    }

    & $signTool verify /pa $file.FullName
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "驗證簽章失敗: $($file.FullName)"
    }
}

# Clean temp pfx from base64
if ($PfxBase64 -and $pfx.StartsWith($env:TEMP)) {
    Remove-Item $pfx -Force -ErrorAction SilentlyContinue
}

Write-Host "簽署完成：$($targets.Count) 個檔案。"
