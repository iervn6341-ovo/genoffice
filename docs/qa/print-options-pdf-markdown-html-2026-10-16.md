# PDF／Markdown／HTML 列印選項評估（2026-10-16）

## 路徑總覽
- macOS：三個應用的「列印」一律直接叫出 macOS 原生列印面板（含預覽），不再開 Chromium 面板。
  - PDF：先 flush（存檔，編輯、旋轉、刪頁都寫入檔案），把磁碟上的實際 PDF 交給 `printPdfWithSystemPanel`（向量，不點陣化）。頁碼範圍由原生面板負責。IPC：`pdf:print-native`。
  - Markdown／HTML：renderer 組出自含式列印 HTML（作者原色，無編輯介面、無深色重映射），main 在隱藏、停用 script 的視窗以 `printToPDF`（紙張依地區 Letter/A4，邊界 0.6in，含背景）寫成暫存 PDF，再交給原生面板。IPC：`markdown:print-document`、`html:print-document`。取消不報錯；失敗以 toast／狀態列顯示。
- 非 macOS：PDF 沿用既有列印對話框（加選項）；Markdown／HTML 新增列印對話框（HTML 原本完全沒有列印處理），main 以隱藏視窗 `webContents.print` 送系統對話框。

## PDF（非 macOS）
| 項目 | 狀態 | 本次 |
|---|---|---|
| 頁碼範圍（全部／目前頁／自訂） | 已有 | — |
| 奇偶頁 | 缺少 | 已實作（依頁碼奇偶） |
| 反向順序 | 缺少 | 已實作 |
| 頁面調整：符合紙張／縮小超大頁面／實際大小／自訂縮放 | 缺少（只有符合） | 已實作（實際大小與自訂會置中並裁切超出部分） |
| 置中 | 部分 | 縮放模式皆置中 |
| 自動旋轉 | 缺少 | 不做：頁面方向由系統對話框的版面方向決定，無法在 renderer 可靠驗證 |
| 註解／標記是否列印 | 缺少 | 已實作（pdf.js annotationMode 關閉） |
| 每張紙頁數、雙面、逐份、份數、印表機 | 系統對話框提供 | 不適用 |

## Markdown／HTML（非 macOS）
| 項目 | 狀態 | 本次 |
|---|---|---|
| 紙張（A3/A4/A5/Letter/Legal/Tabloid） | 缺少 | 已實作 |
| 方向 | 缺少 | 已實作 |
| 邊界（預設／無／窄／寬） | 缺少 | 已實作（@page） |
| 縮放 10–200% | 缺少 | 已實作 |
| 頁首頁尾（標題＋頁碼） | 缺少 | 已實作 |
| 背景色與圖片 | 缺少 | 已實作 |
| 頁碼範圍 | 缺少 | 已實作（可寫 8- 表示到最後） |
| 份數、雙面、逐份、印表機 | 系統對話框提供 | 不適用 |

## 驗證
單元測試：平台分流、頁碼範圍解析、奇偶／反向、縮放樣式、選項對應 `webContents.print`、i18n 鍵集。macOS 原生面板以探針驗證（`e2e/_probe/print-others-mac.probe.spec.ts`）。Windows／Linux 未實測。
