# 程式碼簽章（Code Signing）與智慧型應用程式控制

Windows 11 **智慧型應用程式控制（Smart App Control）** 會封鎖**未以受信任憑證簽署**的應用程式。  
本專案 Release 預設可選簽署；未設定憑證時產出**未簽章**套件。

## 建議方案

| 方案 | 說明 | 是否能過 Smart App Control |
|------|------|---------------------------|
| **公開 CA 的 Code Signing 憑證**（Sectigo、DigiCert、SSL.com…） | 購買後匯出 PFX | 是（新標準憑證建議有時間戳） |
| **Azure Trusted Signing** | Microsoft 雲端簽署服務 | 是（需 Azure 訂閱與身分驗證） |
| **自簽憑證** | 僅本機測試 | **否**（SAC 仍會擋） |
| **關閉／評估 SAC** | 使用者端權宜 | 暫時可用，非發行解法 |

## 本機簽署

### 1. 準備 PFX

將憑證匯出為 `.pfx`（含私鑰），並記住密碼。

### 2. 環境變數

```powershell
$env:CODE_SIGNING_PFX_PATH = "C:\certs\codesign.pfx"
$env:CODE_SIGNING_PFX_PASSWORD = "your-password"
```

### 3. 打包並簽署

```powershell
# 需已安裝 Windows SDK（內含 signtool.exe）
.\scripts\publish-release.ps1 -Version v1.0.0
```

或只簽署某個檔案：

```powershell
.\scripts\sign.ps1 -Path .\artifacts\sc-pack\PptxAvalonia.exe
```

### 4. 驗證

```powershell
Get-AuthenticodeSignature .\artifacts\sc-pack\PptxAvalonia.exe | Format-List *
```

`Status` 應為 `Valid`。

## GitHub Actions 自動簽署

Repository → **Settings → Secrets and variables → Actions** 新增：

| Secret | 內容 |
|--------|------|
| `CODE_SIGNING_PFX_BASE64` | PFX 檔的 Base64 字串 |
| `CODE_SIGNING_PFX_PASSWORD` | PFX 密碼 |

產生 Base64（PowerShell）：

```powershell
[Convert]::ToBase64String([IO.File]::ReadAllBytes("C:\certs\codesign.pfx")) | Set-Clipboard
```

然後：

- 推送 tag：`git tag v1.0.1 && git push origin v1.0.1`  
  或  
- Actions → **Release** → **Run workflow**

Workflow 會執行 `scripts/publish-release.ps1`；若偵測到 secrets 則自動 `signtool sign`。

## Azure Trusted Signing（進階）

若使用 [Azure Trusted Signing](https://learn.microsoft.com/en-us/azure/trusted-signing/)：

1. 在 Azure 建立 Trusted Signing 帳戶與憑證設定檔  
2. 在 CI 使用 `Azure/trusted-signing-action`（或官方 CLI）取代 PFX 步驟  
3. 可另開 workflow job，於 publish 後呼叫 Trusted Signing  

本 repo 目前以 **PFX + signtool** 為主路徑，方便一般 Code Signing 憑證接入。

## 沒有憑證時使用者怎麼開

見 README「智慧型應用程式控制已封鎖？」：

1. exe 內容 → **解除鎖定**  
2. 智慧型應用程式控制改為評估／關閉  
3. 企業環境請 IT 允許清單  

## 安全注意

- **勿**把 PFX 或密碼提交進 Git  
- 僅放在 GitHub Secrets 或本機安全存放  
- CI log 勿列印密碼與 Base64 全文  
