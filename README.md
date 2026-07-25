# PptxAvalonia

以 **Avalonia UI** 打造的桌面程式，提供接近 Microsoft PowerPoint 的常用**檢視／放映**功能，可開啟並預覽 `.pptx`。

## 功能（對齊 PowerPoint 常用操作）

### 檔案
- 開啟 `.pptx`（選單 / `Ctrl+O` / 拖放）
- 最近開啟的檔案
- 匯出目前投影片 / 全部投影片為 **PNG**
- 關閉簡報

### 檢視
- **一般**：縮圖 + 預覽 + 頁面文字 + 備忘稿
- **投影片瀏覽**：網格縮圖（`Ctrl+2`）
- **大綱**：文字大綱與備忘稿（`Ctrl+3`）
- 符合視窗 / 縮放 / 顯示或隱藏縮圖與備忘稿

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

## 建置與執行

```bash
dotnet restore
dotnet build -c Release
dotnet run -c Release -- Samples/demo.pptx
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
| `Services/SlideRenderer.cs` | Avalonia 視覺樹預覽 |
| `Services/SlideExportService.cs` | 匯出 PNG |
| `Services/RecentFilesService.cs` | 最近檔案 |
| `Views/SlideShowWindow.axaml` | 全螢幕放映 |
| `ViewModels/MainViewModel.cs` | 導覽、檢視、放映、尋找、匯出 |

## 限制

輕量預覽，非完整 PowerPoint 編輯器。不支援（或僅近似）：圖表、表格、智慧圖形、動畫過場、漸層、影片、完整字型替代。

## 授權

示範專案，可自由修改使用。
