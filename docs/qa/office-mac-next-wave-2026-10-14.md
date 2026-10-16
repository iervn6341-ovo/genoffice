# Office Mac 細項下一階段驗收（2026-10-14）

比對基準為 Microsoft 365 Mac 桌面版。Docs、Sheets、Slides 由三位 GPT-6 Sol／High 實作者分線處理，GPT-6 Luna／Extra 檢查員獨立複驗。此階段接續[前一階段的並行執行紀錄](./office-mac-concurrent-e2e-2026-10-14.md)；先前已以三個不同 Electron PID 和獨立使用者資料目錄證實三個應用同時開啟。本階段各線可同時開發，對 macOS 視窗焦點敏感的 GUI 操作仍以各自前景驗收結果為準。

| 模組 | 本階段修正與 Mac 依據 | 實際應用驗收 |
| --- | --- | --- |
| Docs D6 | 分節註腳預覽讀取所屬節的編號格式與起始值，維持原始段落錨點；Word Mac 兩節原生 DOCX 與合成跨節案例 | Docs+Shell Electron `docs-note-section-preview-verification.spec.ts` 2/2：`III → h`、跨節 `4 → 4`、Undo／Redo、存檔、新程序重開再編輯、原生 reference／note part 保留。結果：`/private/tmp/genoffice-docs-note-section-fixed-results`。聚焦 19/19、相關 typecheck／lint 通過。 |
| Sheets S4 | 合併儲存格外框依可見邊界編輯；拆合併保留四角各自的原生 XF／邊框，並同步選取狀態。依 Excel Mac 原生 XLSX 與拆合併結果 | Sheets+Shell Electron `sheets-merged-border-verification.spec.ts` 2/2：編輯、Save→Undo→Save→Redo→Save、重開再編輯、四角 XF／邊框與 C2 `<f>1+2</f><v>3</v>` 公式快取保留，選取 `isMerged` 跟隨拆合併／Undo／Redo。結果：`/private/tmp/genoffice-sheets-merged-border-final9-results`。聚焦 9/9、build／typecheck／Prettier／scoped ESLint 通過。 |
| Slides L5 | 依 PowerPoint Mac 十一種原生 Quick Layout OOXML 快照，加入圖表版面 1、3；只編輯已支援的圖例、標籤、gapWidth／overlap 欄位 | Slides+Shell Electron `slides-chart-native-format-verification.spec.ts` 的 Layout 1／3 案 1/1：編輯、Undo／Redo、Save As、新程序重開再編輯；保留資料序列／快取、未知 XML、關聯、樣式／配色與內嵌活頁簿。結果：`/private/tmp/genoffice-slides-quick-layout-fixed-results`；重開夾檔見 `office-mac-field-evidence-2026-10-13/slides-chart-native/quick-layout-mac/genoffice-layout-3-reopened.pptx`。聚焦 16/16、相關 typecheck／lint 通過。 |

原生來源集中於 `office-mac-field-evidence-2026-10-13` 下的 `docs-native-notes/word-mac-two-section-notes-source.docx`、`sheets-merged-border/excel-mac-merged-border-source.xlsx` 與 `slides-chart-native/quick-layout-mac/`。Docs 的 Word Mac 夾檔兩個註腳 reference 都落在第二節，因此跨節預覽以另外的合成夾檔驗證；未把該原生夾檔誤當跨節 reference 證據。Slides 首次 Electron 測試曾載入過期 Shell bundle，重建 Slides+Shell 後上述結果通過。

三位 Luna／Extra 檢查員對各自模組完成獨立複驗，均無本階段 blocker。Sheets 合併格的新增邊框 render-cache 線段有單元測試，未另做整張畫布的像素比對；這不影響已驗證的原生 XF、公式快取與重開編輯行為。

剩餘差異：Docs 的逐頁註腳重編與節末章末註尚無對等畫布預覽，選取文字／單節作用範圍、自訂符號、註腳轉換與導覽待做。Sheets 對只選到合併範圍一部分的框線操作保守拒絕；單一合併範圍的內側框線依 Excel 停用，會計底線的畫布呈現仍待補。Slides 目前只覆蓋觀察到的十一種 Mac Quick Layout 中兩種原生版面；其餘版面及其他進階圖表／圖片工具待做。D6、S4、L5 因此仍為 `in_progress`，不是完整 Microsoft 365 Mac 對等。
