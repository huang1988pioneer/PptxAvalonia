# 解除 Windows「下載自網際網路」標記（Mark of the Web），
# 有助於減少 SmartScreen / 部分封鎖。
# 用法：在解壓後的資料夾執行  解除封鎖.bat  或：
#   powershell -ExecutionPolicy Bypass -File unblock-after-download.ps1

$ErrorActionPreference = "Continue"
$dir = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }

Write-Host "解除封鎖目錄: $dir"
$files = Get-ChildItem -Path $dir -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Extension -match '\.(exe|dll|ps1|bat)$' }

$count = 0
foreach ($f in $files) {
    try {
        Unblock-File -Path $f.FullName -ErrorAction Stop
        $count++
        Write-Host "  OK  $($f.Name)"
    } catch {
        Write-Host "  SKIP $($f.Name): $($_.Exception.Message)"
    }
}

Write-Host ""
Write-Host "已處理 $count 個檔案。"
Write-Host ""
Write-Host "若仍出現「智慧型應用程式控制已封鎖」："
Write-Host "  設定 > 隱私權與安全性 > Windows 安全性 > 應用程式及瀏覽器控制項"
Write-Host "  > 智慧型應用程式控制設定 > 改為「評估」或「關閉」"
Write-Host ""
Write-Host "注意：關閉後通常無法再改回「開啟」。"
Write-Host "按任意鍵結束..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
