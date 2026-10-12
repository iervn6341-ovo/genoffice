# 交接文件：功能區／下拉選單稽核的未完成項目（2026-10-11）

承接 [ribbon-dropdown-audit-2026-10-11.md](ribbon-dropdown-audit-2026-10-11.md)。該文件記錄已修的問題；本文件只列**還沒做完或沒驗證**的三件事，以及接手時怎麼做。

## 接手進度（2026-10-11，同日後續）

拉取 `b6401c9b` 後依本文件實作與測試；三點的結果如下，細節見 [ribbon-dropdown-audit-2026-10-11.md](ribbon-dropdown-audit-2026-10-11.md) 的「Sheets (fourth pass)」。

| 項目                     | 結果                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| ------------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 1. 與真正 Excel 並排比對 | **已做（Sheets）**：QA 機器（Windows）裝有 Excel 16，可用 PowerShell COM 自動化直接讀取行為與樣式（見記憶 `reference-excel-com-ground-truth`）。填滿／序列／選擇性貼上／清除超連結／轉為範圍的語意、60 種表格樣式的元素與色彩都取自 Excel 實測；Medium 3 ＝ 橙、Medium 4 ＝ 灰已確認；App 存出的 xlsx 也用 Excel 開啟核對。Docs 的框線／行號／分欄預設值**尚未**用 Word 比對（Word 也已安裝，可用同法）。⌘Z 一次一步（頁面設定、主題）未用 Excel 實測 |
| 2. 稽核最後一輪重跑      | **未重跑**：`e2e/_probe/audit-sheets.probe.spec.ts` 不在 repo（探針從未 commit），這台機器上也沒有。改以可提交的規格覆蓋：`sheets-ribbon-fit`（1280／1000／900 寬度不橫向捲動，新增的 Cells 下拉、Fill／Clear／Paste 選單通過）、`sheets-ribbon-fixes`（含主題復原）與本次新增的 5 支規格；全部 Sheets e2e 見下。**PivotTable 仍未稽核**。若要完整重跑，需重建探針                                                                                    |
| 3. Sheets 缺的功能       | **已實作並測試**：Fill 上／左／序列、Clear 註解／超連結、選擇性貼上（轉置、值＋數字格式、對話框含運算／略過空白）、Cells 群組的 Insert／Delete 下拉（含插入／刪除工作表）、Format as Table 60 種樣式圖庫、Table Design 分頁（名稱、樣式選項、轉為範圍）、主題／色彩／字型即時預覽圖庫。**未做**：New Table Style、Header／Total Row 切換、Resize Table、檔案原有表格的設計（唯讀）、選擇性貼上的註解／驗證                                            |

這次順手修掉的既有問題：Review → Delete note（檔案型活頁簿）與 AI 刪註解**不能 ⌘Z**（Univer 的 `deleteNote()` 不進復原堆疊）；`xlsx-sidecar-cancel.test.ts` 在 Windows／Node 24 上**通過**（本文件的 1 個失敗是 Mac／Node 26 的環境問題）。

驗證指令在 Windows 上要改：Sheets 的 `npm run build` 會卡在 `cargo`（沒裝 Rust），用 `cd apps/sheets && npx electron-vite build`（預編的 sidecar 在 `native/xlsx-engine/target`）；改過 `apps/*/src/main` 要另外 `npm run build -w @genoffice/shell`。

## 第二輪（2026-10-12）：把上表「未做／有限制」的項目補上

依使用者要求，把第一輪列為未做或有限制的項目逐項實作與測試；行為先用 Excel／Word 的 COM 讀出，再寫單元測試，再從 UI 實際操作，涉及存檔的再用 Excel 開啟存出的檔案。細節見 [ribbon-dropdown-audit-2026-10-11.md](ribbon-dropdown-audit-2026-10-11.md) 的「Docs (fourth pass)」與「Sheets (fifth pass)」。

| 項目                                                                | 結果                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| ------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| 稽核重跑（7 個分頁）                                                | **已重跑並通過**：`e2e/sheets-ribbon-audit.spec.ts`（新進 repo）7 個分頁共按 284 個選項／按鈕，頁面錯誤 0、console 錯誤 0、不可復原 0、復原不完整 0、失敗 0（17 項在略過清單：PivotTable／PivotChart、圖片、螢幕擷取、註解編輯、列印切換、保護活頁簿、AI 按鈕、New Rule／New Table Style）。查出並修好兩個真問題：**Size 下拉選項被 AI 面板的 6px 縮放把手擋住**（`.excel-header` z-index 10 → 25，加回歸測試）、**約 20 個對話框按 Esc 不會關**（共用 `use-dialog-escape.ts`，`sheets-dialog-escape` 規格）；另有一項是稽核腳本本身的狀態外洩（已修）。Watch Window 是非強制面板，Esc 不關是預期。PivotTable 仍只靠各自的規格 |
| Docs 與真 Word 比對                                                 | **已做（4 項）**：Home → Borders（`w:space` 1／4 pt，**原本是 0，已修**）、Layout → Line Numbers、Design → Page Borders ▸ Box、Columns ▸ Two ＋分隔線；其餘 Docs 項目仍只依文件與預設值，未並排比對                                                                                                                                                                                                                                                                                                                                                                                                                            |
| New Table Style／Header Row／Total Row／Filter Button／Resize Table | **已做**（session 表格）；存檔寫 `<tableStyle>`＋`<dxf>`、`headerRowCount="0"`、`totalsRowCount`，Excel 開啟核對                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| 檔案原有表格的 Table Design                                         | **部分**：樣式圖庫、Format as Table、四個樣式選項（Banded Rows／Columns、First／Last Column）可改且可 ⌘Z（含跨存檔），存檔改寫表格部件的 `tableStyleInfo`。名稱、Resize、轉為範圍、Header／Total Row、Filter Button 仍只限本次建立的表格（分頁中停用並附提示）                                                                                                                                                                                                                                                                                                                                                                 |
| 選擇性貼上的註解／驗證／來源佈景／合併條件式格式／合併儲存格        | **已做**；貼上於 Cut 之後會被拒絕（Excel 也失敗）                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| 內建佈景                                                            | **已做**：Excel 的 48 個內建佈景與其色彩／字型配置                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             |
| 鍵盤 Delete（超連結格）                                             | **已修**                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| 頁面設定、主題的 ⌘Z 跨存檔                                          | **已修**（undo 記錄「有效值」＝日誌疊在檔案上）                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| 偶發的貼上對話框失敗                                                | **原因查明**：測試期間另一個程序（Excel COM `Copy()`）覆蓋了系統剪貼簿；未受干擾的重跑 12／12 通過，規格改為自行寫剪貼簿                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| Fill Across Worksheets／Justify                                     | **已做**（Across 以對話框選工作表，因為分頁籤沒有群組選取）                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| 翻譯                                                                | **仍是機器翻譯**（17 個非英文語系），未經母語者審閱                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            |

順手修掉的既有問題：開啟含表格的活頁簿後**第一次 ⌘Z 沒有作用**（載入時註冊的 plain 主題被當成使用者編輯，`undo-carry.ts`）；表頭儲存格在建表後被改名，存檔時表格欄名與儲存格不一致（Excel 會要求修復，閘道現在於存檔時對齊）。

驗證（最終程式碼）：Sheets 單元 3009 通過（1 略過）、閘道 60；**Sheets e2e 全套 109 個全過**（含稽核以外的所有 `sheets-*` 規格）；Sheets／閘道 `tsc`、prettier、`check-theme-colors` 皆過。先前 Docs 規格使用 `Meta+z／Meta+a`（Mac 按鍵）在 Windows 上要把 `Meta` 換成 `Control` 才會過；Docs 的 `protect-dialog` 單元測試在整套同時跑時會逾時（雜湊密碼 10 秒），單獨跑通過（Docs 這輪沒有再改，未重跑）。跑 e2e 時若設 `QA_KEEP_DIR`，那個資料夾要事先存在，否則每個「留存存檔」的規格會在最後一步失敗（`copyfile ENOENT`），看起來像產品錯誤，其實不是。

仍然開放的限制：重新套樣式的檔案表格以「僅顯示用」的 range theme 繪製，列印／PDF 輸出可能不含新外觀（未檢查）；檔案自己定義的 `<tableStyle>` 無法在 app 內繪製（選項切換會被拒絕並提示先選圖庫樣式）；巨大表格（載入超過 50 萬格）拒絕重新套樣式；跨存檔的 ⌘Z 寫不回「自訂邊界」與「多個列印範圍」（日誌沒有對應欄位）；Justify 以 canvas 量測，換行可能與 Excel 差一個字；新字串的 17 個非英文語系為機器翻譯；PivotTable 稽核仍被略過。

## 現況（交接當時）

- Slides、Docs（缺口已補）、Sheets 已稽核並修正；修改全在工作目錄，**尚未 commit**（是否推到 `qa/local-fixes` 待使用者回覆）。
- 不要 commit：`.claude-flow/`、`apps/docs/.claude-flow/`、`1_Stock_Funds_Rose_______in_____.pdf`、`setup-genoffice-qa-skill.sh`、`e2e/_probe/`、`test/`、`第四次估咪簡報.pptx`。
- 最後一次驗證：Sheets e2e 全部 45 個通過；Sheets 單元 2819 通過、1 失敗（`apps/sheets/tests/xlsx-sidecar-cancel.test.ts`，主程序 sidecar 客戶端，本次未動，但**未在乾淨 HEAD 上驗證**是否本來就失敗）；Docs 單元 2529 通過、e2e 49 通過／6 略過；三個 app 的 `tsc`、eslint、prettier、`tools/check-theme-colors.mjs` 皆過。

## 1. 沒有跟真正的 Word／Excel 並排比對

**現況**：新增的 Docs 與 Sheets 行為是依 Office 文件與記憶寫的，沒有在本機的 Word／Excel 上逐項對照。

**特別需要核對的點**

| 項目                                                                      | 目前做法                                                                                                                   | 風險                                                                                                                                                                                       |
| ------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Sheets Format as Table 外觀（`apps/sheets/src/renderer/table-styles.ts`） | 以 Office 主題色 `#4472C4 #ED7D31 #A5A5A5 #FFC000 #5B9BD5 #70AD47` 近似 Light／Medium／Dark 外觀；淺色條紋為「亮 80%」混白 | 色值與 Excel 實際有些微差（例：Excel 的 Accent1 亮 80% 為 `#D9E1F2`，這裡算出 `#DAE3F3`）；Dark 2、Light 的細節（外框、邊線）是推測。Excel 開檔時會用自己的版本繪製，所以只影響 app 內預覽 |
| Medium 3／Medium 4 的對應                                                 | 依記憶：Medium 3＝橙（accent2）、Medium 4＝灰（accent3），因此把「Medium 4 (orange)」改成 Medium 3                         | 應在 Excel 的表格樣式圖庫確認編號與顏色                                                                                                                                                    |
| Docs 的框線、行號、分欄、頁面框線、樣式（Title／Quote 等）預設值          | 依 Word 預設範本數值                                                                                                       | 未以真 Word 的 PDF 輸出比對（可用 `test/_probe` 的 Word PDF 流程，見記憶 `reference-word-pdf-export-applescript`）                                                                         |
| 頁面設定、主題的 ⌘Z 行為                                                  | 以 Excel 行為為準撰寫                                                                                                      | 未在 Excel 實測 ⌘Z 是否一次一步                                                                                                                                                            |

**怎麼做**

1. 準備同一份測試資料（Sheets：A1:D6 含表頭；Docs：數段文字）。
2. 在 Excel／Word 套用同一組選項，截圖或匯出 PDF；在 GenOffice 重做同樣操作。
3. 差異寫進 QA 文件；顏色差異集中改 `table-styles.ts` 的 `ACCENTS`、`tint`、`shade` 與各 family 分支。
4. Word／Excel 基準與 Playwright 指令見記憶 `reference-office-apps-for-comparison`。

## 2. 稽核最後一輪沒有重跑完整的 7 個分頁

**現況**：最後的修正（Format as Table、頁面設定／主題復原、公式列隱藏）之後，只重跑了：

| 分頁                           | 結果                                                                                                                          |
| ------------------------------ | ----------------------------------------------------------------------------------------------------------------------------- |
| Home                           | 剩 Format Painter（稽核假象：只是啟動格式刷，沒有改動）與 Insert row（單獨測試乾淨，原因未查明，判斷為序列殘留）              |
| Page Layout                    | 剩 Themes › Office；之後加了主題復原，但**沒有再跑這個分頁**，只靠 `e2e/sheets-ribbon-fixes.spec.ts` 的主題測試驗證           |
| View                           | 公式列隱藏的主控台錯誤已用探針確認消失；稽核本身沒在修正後重跑                                                                |
| Insert、Formulas、Data、Review | 只在修正前的完整一輪（2400 寬、略過 Pivot）跑過；Data Validation 的 2 筆 `REOPEN-FAILED` 屬工具重開彈出層失敗，未視為產品問題 |

另外：**PivotTable 被略過**（`AUDIT_SKIP=Pivot`；每格開對話框，會超過 1 小時）。Pivot 目前只靠既有 e2e。

**怎麼做**（在 repo 根目錄；跑稽核期間不要 build、不要同時跑其他 e2e）

```bash
mkdir -p /tmp/audit-sheets && AUDIT_OUT=/tmp/audit-sheets CTX=range LAYOUT=2400 AUDIT_SKIP=Pivot \
  npx playwright test --config e2e/playwright.config.ts e2e/_probe/audit-sheets.probe.spec.ts --timeout=3300000
```

- 只跑某分頁：加 `TABS="Page Layout"`（逗號分隔多個）。全部 7 個分頁約 13 分鐘。
- 結果在 `$AUDIT_OUT/sheets-range-2400/all.summary.txt` 與 `all.findings.json`。
- 等待背景執行用 `pgrep -f "[b]in/playwright test"`；用 `nohup … &` 會立刻回報「完成」，那只是啟動器。
- `e2e/_probe/` 不 commit。
- 預期殘留：Format Painter、Insert row（若仍出現，先用 `e2e/_probe/sheets-undo3.probe.spec.ts` 單獨驗證，再決定是否在稽核序列中找殘留來源）。
- 若時間允許，補跑 `CTX=cell`（單一儲存格選取）與較窄寬度（`LAYOUT=1280`、`1000`）。

## 3. Sheets 仍缺的功能（與 Excel 的差距）

已對照 `ExcelShell.tsx`、`ribbon-actions.ts` 確認：

| 缺口            | 現況                                                                                                                                                    | 建議位置                                                                                         |
| --------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------ |
| Fill            | 只有 Down、Right，沒有 Up、Left、Series…                                                                                                                | `ExcelShell.tsx` Fill 選單；`ribbon-actions.ts` 的 `fill-down`／`fill-right` 旁新增              |
| Clear           | 只有 All、Formats、Contents，沒有 Clear Comments、Clear Hyperlinks                                                                                      | 同上，Clear 選單                                                                                 |
| 選擇性貼上      | Values、Formulas、Formats、Column Widths、All Except Borders；**沒有轉置**、沒有 Values & Number Formatting、沒有選擇性貼上對話框                       | `ribbon-actions.ts` 的 `paste-special:*`                                                         |
| 儲存格群組      | 四顆獨立的插入／刪除列／欄按鈕；插入儲存格、刪除儲存格（向右／下／左／上移動）的對話框只能用 ⌃⇧= 與 ⌘- 叫出，功能區沒有下拉；下拉裡沒有插入／刪除工作表 | `ExcelShell.tsx` Cells 群組；對話框已有（`InsertCellsDialog.tsx`）                               |
| Format as Table | 只有 6 種（Excel 60 種）、沒有 New Table Style、沒有 Table Design 分頁                                                                                  | `ExcelShell.tsx` 選項清單；外觀見 `table-styles.ts`（目前支援 Light 1–14、Medium 1–7、Dark 1–7） |
| 版面配置 → 主題 | 沒有帶即時預覽的圖庫                                                                                                                                    | `page-layout-actions.ts` 的 `theme*`                                                             |

**實作時的注意事項**（都踩過）

- Sheets 凡是「只改編輯日誌（journal）」的動作不會自動進 Univer 的 ⌘Z 堆疊。新增這類功能時，用 `journal-undo.ts` 的 `withJournalUndo`（範例：`page-setup-undo.ts`、`theme-undo.ts`）。表格的復原與重做則由 `table-sync.ts` 跟著 Univer 的 `delete-table`／`add-table` 變動同步日誌。
- 一個功能區指令若跑好幾個 Univer 指令，用 `undo-group.ts` 的 `runAsOneUndoStep`（非同步用 `runAsOneUndoStepAsync`），確保一次 ⌘Z 回復。
- 新增字串要進 `zh.ts` 與其餘 19 個語系分片（`i18n/app/*.ts`）；批次腳本範本是 `e2e/_probe/i18n_sheets_3.py`（用 `e2e/_probe/add_i18n.py` 的 `add`）。
- 功能區控制項要放進 `RibbonFoldGroup`／`fold` 屬性並設 `priority`，確保 1280／1000／900 寬度都不左右捲動（`e2e/sheets-ribbon-fit.spec.ts` 會檢查）。
- 舊的 Sheets e2e 若要點功能區按鈕，需先 `setEditorLayoutWidth(app, '://sheets/', 2400)`，定位時用 `.ribbon-tool:not(.rb-fold-btn)`（折疊按鈕隱藏時仍在 DOM，會造成 strict mode 衝突）。
- CSS 只用佈景 token（見 CLAUDE.md）；Sheets 的 `apps/*/src/main` 沒動，所以不需重建 shell。

## 其他已知但不在三點之內

- Docs 仍缺項目清單在 QA 文件的「Docs gaps vs Word (still open)」。
- 新增字串的 15 種非英語翻譯是機器翻譯，未經母語者審閱。
- 驗證指令：`cd apps/sheets && npx tsc --noEmit && npx vitest run`；e2e 用 `npx playwright test --config e2e/playwright.config.ts e2e/sheets-*.spec.ts`（約 11 分鐘，跑時不要 build）。
