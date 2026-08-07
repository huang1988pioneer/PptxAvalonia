@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo  正在解除 Windows 下載封鎖標記...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0unblock-after-download.ps1"
if errorlevel 1 (
  echo.
  echo 若 PowerShell 失敗，請手動：
  echo   右鍵 PptxAvalonia.exe - 內容 - 勾選「解除鎖定」- 套用
  pause
)
