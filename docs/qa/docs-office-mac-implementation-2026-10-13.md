# Docs Office for Mac 第一波實作紀錄

日期：2026-10-13（Asia/Taipei）。本紀錄描述目前工作樹的暫定實作，供獨立 QA 複驗；原始 2026-10-12 稽核報告及證據未改動。

## 已改動

- Insert → Header/Footer 文字入口改為多行編輯，顯示並可插入 `{PAGE}` 與 `{NUMPAGES}` token；透過 `applyHfText` 保留未變動文字的 run 格式和表格列。頁碼位置圖庫改為在既有段落加入 PAGE 欄位，或調整已有 PAGE 段落對齊，不再把整個頁首頁尾替換成一段 `#`。
- 畫布頁首頁尾控制加入節切換及目前變體的「連結到前一節」狀態。Ribbon 的頁首頁尾編輯入口接上目前顯示的節／變體。非末節的 first/odd(default)/even 編輯由 `sectionHf` 儲存；非末節的「首頁不同」透過既有 section rewrite 儲存。PAGE 和 NUMPAGES 仍以引擎欄位語意保存。
- `saveDocx` 的 `sectionHf` 可指定變體與繼承來源。每次節內容編輯另建 part，以原 part 為合併基底，連同 part relationship 一起複製，保住表格與圖片及供跨存檔 Undo 使用的舊 part；明確 local ref 若與其他節共用 rId，也只替換目前節／變體的 ref。新增 ref 時可展開 self-closing `sectPr`，末節的 ref 寫入 trailing `sectPr`。`sectionHfLinks` 移除該節的 local reference，「恢復 local ref」則重用仍在檔案內的舊 rId／part。文字編輯會清除同一節／變體的 pending link，儲存引擎在兩種操作同時傳入時也讓內容編輯優先。layout snapshot 記錄各節 local ref，Undo 橫跨存檔時比較目前 ref 產生反向 link/ref 寫入；header-only 存檔的 ProseMirror 比較保留編輯器的 layout attr，避免誤判為本文改變而清空歷史。預覽會沿前節的 pending 文字修改解析繼承內容。
- 共用 `RibbonFoldGroup` 折疊時，內層開啟 `[data-rb-panel]`、`.menu-select-drop`、`.gcp-palette` 或 `aria-expanded=true` 的按鈕不關閉父群組；點完成命令、外側及 Escape 關閉。內層非完成步驟可用 `data-rb-keep-open` 標記。
- 三個頁首頁尾字串鍵已加入所有既有 ribbon locale shard；CSS 使用語意主題 token。

## 已執行檢查

- `./node_modules/.bin/vitest run packages/docx-engine/tests/section-hf-workflow.test.ts packages/docx-engine/tests/hf-variants.test.ts packages/docx-engine/tests/header-footer-rich.test.ts apps/docs/tests/doc-dirty.test.ts apps/docs/tests/hf-text.test.ts apps/docs/tests/layout-history.test.ts packages/ui/tests/ribbon-fold-group.test.ts`：7 files，91 tests 通過。新增案例覆蓋繼承的 first 變體另建 part 並保留 table/picture relationship、明確共用 rId 時只替換目標節 reference、self-closing `sectPr`、末節 first 變體與衝突 pending link 儲存、連續三次存檔的 local ref 移除／精確恢復、PAGE/NUMPAGES 往返、跨存檔 Undo ref 差異計算、header-only save 的 layout attr 比較、dirty reset，以及折疊群組的巢狀啟動／完成命令／外點／Escape。
- `./node_modules/.bin/tsc --noEmit -p apps/docs/tsconfig.json`：通過。
- `./node_modules/.bin/tsc --noEmit -p packages/docx-engine/tsconfig.json`：通過。
- `./node_modules/.bin/tsc --noEmit -p packages/ui/tsconfig.json`：通過。
- `./node_modules/.bin/eslint packages/ui/src/RibbonFoldGroup.tsx apps/docs/src/renderer/{App.tsx,components/ribbon-insert-tab.tsx,components/PaginationPreview.tsx,doc-dirty.ts,editor/hf-text.ts,editor/layout-history.ts,file-actions.ts} apps/docs/tests/{doc-dirty,hf-text,layout-history}.test.ts packages/docx-engine/src/patch.ts packages/docx-engine/tests/section-hf-workflow.test.ts`：0 errors；`App.tsx` 另有 1 個既有 `react-hooks/exhaustive-deps` warning。
- `git diff --check`：通過。`node tools/check-theme-colors.mjs`：無新增 raw renderer CSS 色值。
- 本實作者未 build，也未啟動 Electron 或原生 Office；窄／寬窗實際 UI 與儲存關閉重開仍待獨立檢查員。

## 已知未驗及需先修的風險

1. **未建模動態欄位**：頁首頁尾引擎可保留既有 DATE 等欄位的原始 bytes，然而編輯同一 text paragraph 後，該段仍會重新序列化，未建模欄位可能降為 cached 文字；目前沒有 DATE/圖片插入 UI。本輪只提供 PAGE/NUMPAGES 欄位。下一波應先做非 PAGE 欄位的結構保留測試，再增加日期與圖片工作流。
2. **跨存檔 Undo 範圍**：header-only 存檔且本文／docxIndex 對應不變時，layout history 可跨存檔還原舊 local ref 或重新連結前節；若存檔重建了本文節點或索引，既有 save 流程仍會重設編輯器歷史以免沿用舊錨點。既有畫布頁首頁尾文字雙擊編輯本身也尚未進 layout history，需由 QA 操作複驗。多次編輯會留下未被參照的舊 header/footer part，以換取精確 Undo；後續可在確定歷史無需它們後清理。
3. **即時預覽細節**：待儲存的「連結到前一節」文字會解析前節的 pending 修改，包含最後一節 first 變體；pagination preview 的圖片來源在儲存前仍可能沿用舊 local ref，儲存並重開後應由實際 refs 重新解析。節切換目前在單張 canvas 上操作，尚未逐節顯示獨立紙張幾何。
4. **共用折疊群組**：`data-rb-keep-open` 是明確的中間步驟標記；尚未盤點所有 app 的內層分類按鈕。需在 1280px 與寬窗實際操作 Docs Symbol/Equation、Sheets MenuSelect/ColorPicker、其他含分類的內層 palette，檢查完成命令、外點、Escape。

下一個優先批次是未建模欄位／日期／圖片、方程式結構與 Symbol dialog，以及 Font Advanced／Paragraph pagination 等原稽核 P1；不能把本次引擎測試當成完整 Office for Mac 相容驗收。
