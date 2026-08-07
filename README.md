# PptxAvalonia

以 **Avalonia UI** 打造的桌面程式，提供接近 Microsoft PowerPoint 的常用**檢視／放映**功能，可開啟並預覽簡報。

## 功能（對齊 PowerPoint 常用操作）

### 檔案
- 開啟簡報（選單 / `Ctrl+O` / 拖放 / 命令列參數）
  - **`.pptx`**：直接以 Open XML 載入
  - **`.ppt` / `.pps` / `.odp`**：透過本機 **LibreOffice**（或 OpenOffice）轉成 `.pptx` 後預覽
- 未安裝 LibreOffice 時，仍可正常開啟 `.pptx`；開啟 `.ppt` / `.odp` 會提示需安裝
- 最近開啟的檔案
- 匯出目前投影片 / 全部投影片為 **PNG**
- 關閉簡報

### 檢視
- **一般**：縮圖 + 預覽 + 頁面文字 + 備忘稿
- **投影片瀏覽**：網格縮圖（`Ctrl+2`）
- **大綱**：文字大綱與備忘稿（`Ctrl+3`）
- 符合視窗 / 縮放 / 顯示或隱藏縮圖與備忘稿
- **介面風格**（可即時切換，會記住設定）：
  - LibreOffice Impress（經典功能表 + 工具列）
  - Google 簡報（扁平頂列）
  - WPS Presentation（紅色 Ribbon）
  - FreeOffice Presentations（藍色 Ribbon）
  - Microsoft PowerPoint（橘色 Ribbon）

### 投影片放映
- **全螢幕放映**（`F5` 從頭、`Shift+F5` 從目前）
- 左鍵下一張、右鍵上一張
- `B` 黑屏、`W` 白屏、`Esc` 結束
- **自動放映**（可設間隔與循環）

### 導覽與尋找
- 上一張 / 下一張、第一張 / 最後一張（`Home` / `End`）
- 移至投影片（`Ctrl+G`）
- 尋找文字（`Ctrl+F`，含投影片與備忘稿）

## 系統需求

- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)（開發）
- Windows / Linux / macOS（Avalonia）
- Release 自包含包：Windows 10/11 x64（不必安裝 .NET）
- （選用）[LibreOffice](https://www.libreoffice.org/)：開啟 `.ppt` / `.pps` / `.odp` 時需要

## Windows：智慧型應用程式控制已封鎖？

與 [XlsxAvalonia](https://github.com/huang1988pioneer/XlsxAvalonia) 相同，Release **皆未做 Code Signing**。  
兩者打包方式一致（self-contained：`exe` + `pdb` + demo）。  
**智慧型應用程式控制（SAC）可依雲端信譽對「未簽章」程式做出不同判斷**——因此可能出現「Xlsx 可開、Pptx 被擋」的情況，不一定是專案設定錯誤。

### 使用者端

1. 右鍵 `PptxAvalonia.exe` → 內容 → **解除鎖定**  
2. 設定 → Windows 安全性 → 應用程式及瀏覽器控制項 → 智慧型應用程式控制 → **評估／關閉**  
3. 若同目錄有 `crash.log`，可回報內容（與 XlsxAvalonia 相同機制）  
4. 本機開發：`dotnet run -c Release -- Samples/demo.pptx`

### 發行端（根本解法）

| 項目 | 說明 |
|------|------|
| `scripts/sign.ps1` | `signtool` + PFX 簽署 |
| `scripts/publish-release.ps1` | 打包；有憑證則自動簽署 |
| `.github/workflows/release.yml` | Tag / 手動觸發建置與 Release |
| [docs/CODE_SIGNING.md](docs/CODE_SIGNING.md) | 憑證、Secrets、驗證完整說明 |

GitHub Secrets（選填）：`CODE_SIGNING_PFX_BASE64`、`CODE_SIGNING_PFX_PASSWORD`。  
自簽憑證**無法**通過 SAC。

## 建置與執行

```bash
dotnet restore
dotnet build -c Release
dotnet run -c Release -- Samples/demo.pptx
```

### 打包 Release（本機）

```powershell
# 未簽署
.\scripts\publish-release.ps1 -Version v1.0.0

# 有憑證時（會自動簽署 .exe）
$env:CODE_SIGNING_PFX_PATH = "C:\certs\codesign.pfx"
$env:CODE_SIGNING_PFX_PASSWORD = "****"
.\scripts\publish-release.ps1 -Version v1.0.0
```

## 快捷鍵

| 按鍵 | 功能 |
|------|------|
| `Ctrl+O` | 開啟 |
| `F5` / `Shift+F5` | 全螢幕放映（從頭 / 從目前） |
| `Esc` | 結束放映 / 停止自動放映 |
| `B` / `W` | 放映中黑屏 / 白屏 |
| `←` `→` | 上一張 / 下一張 |
| `Home` / `End` | 第一張 / 最後一張 |
| `Ctrl+G` | 移至投影片 |
| `Ctrl+F` / `F3` | 尋找 / 下一個 |
| `Ctrl+1` / `2` / `3` | 一般 / 瀏覽 / 大綱 |

## 架構

| 路徑 | 說明 |
|------|------|
| `Services/PptxLoader.cs` | Open XML 解析（形狀、文字、圖片、備忘稿、大綱） |
| `Services/PresentationFormatConverter.cs` | `.ppt` / `.pps` / `.odp` → `.pptx`（LibreOffice headless） |
| `Services/SlideRenderer.cs` | Avalonia 視覺樹預覽 |
| `Services/SlideExportService.cs` | 匯出 PNG |
| `Services/RecentFilesService.cs` | 最近檔案 |
| `Views/SlideShowWindow.axaml` | 全螢幕放映 |
| `ViewModels/MainViewModel.cs` | 導覽、檢視、放映、尋找、匯出 |

## 限制

輕量預覽，非完整 PowerPoint 編輯器。不支援（或僅近似）：圖表、表格、智慧圖形、動畫過場、漸層、影片、完整字型替代。

`.ppt` / `.odp` 為**轉換後預覽**（非原生解析），版面可能與原軟體略有差異；未安裝 LibreOffice 時無法開啟這兩種格式。

## 授權

示範專案，可自由修改使用。
