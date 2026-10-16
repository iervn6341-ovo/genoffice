# Office Mac 細項驗收與並行執行紀錄（2026-10-14）

基準為 Microsoft 365 Mac 桌面版。三位 GPT-6 Sol／High 實作者分別處理 Docs D6、Sheets S4、Slides L5；GPT-6 Luna／Extra 檢查員獨立複驗來源。測試使用 Word、Excel、PowerPoint Mac 各自產生的原生 DOCX、XLSX、PPTX 夾檔，存於同層 `office-mac-field-evidence-2026-10-13`。

三位實作者在 09:35:57 UTC 同時啟動各自的 Playwright 程序。第二輪授權 GUI 執行的 Electron 啟動紀錄顯示三個不同 PID 於 09:36:00.695、09:36:00.754、09:36:00.767 UTC 開啟，使用三個不同的 `GENOFFICE_USER_DATA` 目錄；第一個直到 09:36:02.609 UTC 才關閉。因此三個應用實例確實同時執行，並非只同時排隊。原始記錄在 `/private/tmp/genoffice-office-e2e-escalated-20261014.vqdmXg/events.log` 與 `electron-events.jsonl`。預設沙盒下三個 Electron 均在 macOS GUI registration 前 SIGABRT；使用 `require_escalated` 後能正常開啟，此失敗不屬產品斷言。

| 模組 | 本輪範圍 | 實際應用結果 | 來源／聚焦檢查 |
| --- | --- | --- | --- |
| Docs D6 | 整份文件的註腳／章末註位置、編號格式／起始／重編，原生 XML 稀疏寫回 | `docs-note-settings-verification.spec.ts` 2/2；Word Mac 夾檔、Undo／Redo、存檔與新程序重開通過 | engine 31/31、Docs 19/19；Luna／Extra 無 OOXML 寫回 blocker |
| Sheets S4 | 逐邊／內側／對角邊框、線型與顏色；複合格式 XF 保留 | `sheets-border-effects-verification.spec.ts` 最終 2/2；Excel Mac 夾檔與 2×2 案、Undo／Redo、公式與新程序重開通過 | 聚焦 100/100、檢查員獨立 64/64；原先兩項資料保留 blocker 已修正 |
| Slides L5 | 三種原生 Quick Layout 圖例／數值標籤預設 | `slides-chart-native-format-verification.spec.ts` 最終 4/4；PowerPoint Mac 夾檔、Undo／Redo、Save As 與新程序重開通過 | 聚焦 12/12；Luna／Extra 無產品 blocker |

Docs 與 Slides 驗收結果來自三線同步執行及 Slides 的完整重跑；Sheets 的完整 2/2 成功紀錄來自後續單線前景重跑。macOS 的共用視窗焦點會讓背景 Electron 的下拉選單在並行互動時瞬間收起，因此前景敏感的選單互動採用單線驗收；三個應用同時開啟的能力已由 PID 與獨立資料目錄證實。各測試共用的 launcher 將 Electron 開關事件寫到 `GENOFFICE_E2E_CONCURRENCY_LOG` 指定檔案。

剩餘差異：D6 的跨節不同編號、逐頁重編與章末註節末位置仍有畫布預覽限制，選取文字／單節作用範圍、自訂符號、轉換與導覽待做。S4 的合併儲存格邊框明確停用，畫布會計底線仍以單／雙底線近似。L5 目前僅涵蓋此 PowerPoint Mac 圖表觀察到的十一個 Quick Layout 中三種，其他圖片／圖表進階工具仍待實作。以上任務維持 `in_progress`，不代表完整 Microsoft 365 Mac 功能對等。
