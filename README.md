# PptxAvalonia

以 **Avalonia UI** 打造的桌面程式，可開啟並預覽 `.pptx`（Office Open XML）簡報。

## 功能

- 開啟本機 `.pptx`（工具列按鈕、`Ctrl+O`、拖放檔案到視窗）
- 左側投影片縮圖列表
- 中央預覽（形狀、文字、圖片、連線）
- 上一張 / 下一張（按鈕或方向鍵、`PageUp`/`PageDown`）
- **符合視窗**完整顯示整頁（預設；`Ctrl+1`）
- 縮放（`Ctrl` + `+` / `-` / `0`）
- **自動放映**（間隔可調、可循環；`F5` / `Shift+F5` / `Esc`）
- 命令列直接開啟：`PptxAvalonia.exe path\to\file.pptx`

## 系統需求

- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)
- Windows / Linux / macOS（Avalonia 跨平台）

## 建置與執行

```bash
dotnet restore
dotnet build -c Release
dotnet run -c Release -- Samples/demo.pptx
```

或執行編譯結果：

```bash
./bin/Release/net8.0/PptxAvalonia.exe Samples/demo.pptx
```

## 快捷鍵

| 按鍵 | 功能 |
|------|------|
| `Ctrl+O` | 開啟檔案 |
| `←` / `→` | 上一張 / 下一張 |
| `F5` | 開始 / 停止自動放映 |
| `Shift+F5` | 從頭自動放映 |
| `Esc` / `空白鍵` | 停止自動放映（空白鍵在非放映時為下一張） |
| `Ctrl+1` | 符合視窗 |
| `Ctrl+0` | 100% 原始大小 |
| `Ctrl++` / `Ctrl+-` | 放大 / 縮小 |

## 架構概要

| 路徑 | 說明 |
|------|------|
| `Services/PptxLoader.cs` | 以 DocumentFormat.OpenXml 解析投影片、主題色、形狀與圖片 |
| `Services/SlideRenderer.cs` | 將記憶體模型轉成 Avalonia 視覺樹 |
| `Models/SlideModels.cs` | 簡報／投影片／元素模型 |
| `ViewModels/MainViewModel.cs` | 開啟檔案、導覽、縮放、自動放映 |
| `Views/MainWindow.axaml` | 主介面（工具列、縮圖列、預覽區） |

座標單位：OOXML EMU → 96 DPI 像素（`÷ 9525`）。

## 預覽能力與限制

目前為**輕量預覽**，涵蓋常見元素：

- 實心填色、邊框、主題色（scheme color + tint/shade）
- 矩形、圓角矩形、橢圓、三角形、菱形
- 文字段落（字級、粗斜體、底線、對齊、垂直錨點）
- 內嵌圖片（png/jpg 等）
- 連線（connector）
- 母片／版面配置上的形狀（簡化合成）

**尚未完整支援**（會忽略或近似）：

- 智慧圖形、圖表、表格
- 複雜群組座標變換、3D、動畫、過場
- 漸層／圖案填滿、影片
- 完整字型替換與文字自動縮放

若需要像素級還原，可另行串接 LibreOffice headless 轉 PDF/影像再顯示。

## 示範檔

`Samples/demo.pptx`：三頁繁中示範簡報，可用於快速驗證預覽。

## 授權

示範專案，可自由修改使用。
