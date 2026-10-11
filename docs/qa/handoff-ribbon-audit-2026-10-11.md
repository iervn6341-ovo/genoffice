# 交接文件：功能區／下拉選單稽核的未完成項目（2026-10-11）

承接 [ribbon-dropdown-audit-2026-10-11.md](ribbon-dropdown-audit-2026-10-11.md)。該文件記錄已修的問題；本文件只列**還沒做完或沒驗證**的三件事，以及接手時怎麼做。

## 現況

- Slides、Docs（缺口已補）、Sheets 已稽核並修正；修改全在工作目錄，**尚未 commit**（是否推到 `qa/local-fixes` 待使用者回覆）。
- 不要 commit：`.claude-flow/`、`apps/docs/.claude-flow/`、`1_Stock_Funds_Rose_______in_____.pdf`、`setup-genoffice-qa-skill.sh`、`e2e/_probe/`、`test/`、`第四次估咪簡報.pptx`。
- 最後一次驗證：Sheets e2e 全部 45 個通過；Sheets 單元 2819 通過、1 失敗（`apps/sheets/tests/xlsx-sidecar-cancel.test.ts`，主程序 sidecar 客戶端，本次未動，但**未在乾淨 HEAD 上驗證**是否本來就失敗）；Docs 單元 2529 通過、e2e 49 通過／6 略過；三個 app 的 `tsc`、eslint、prettier、`tools/check-theme-colors.mjs` 皆過。

## 1. 沒有跟真正的 Word／Excel 並排比對

**現況**：新增的 Docs 與 Sheets 行為是依 Office 文件與記憶寫的，沒有在本機的 Word／Excel 上逐項對照。

**特別需要核對的點**

| 項目 | 目前做法 | 風險 |
|---|---|---|
| Sheets Format as Table 外觀（`apps/sheets/src/renderer/table-styles.ts`） | 以 Office 主題色 `#4472C4 #ED7D31 #A5A5A5 #FFC000 #5B9BD5 #70AD47` 近似 Light／Medium／Dark 外觀；淺色條紋為「亮 80%」混白 | 色值與 Excel 實際有些微差（例：Excel 的 Accent1 亮 80% 為 `#D9E1F2`，這裡算出 `#DAE3F3`）；Dark 2、Light 的細節（外框、邊線）是推測。Excel 開檔時會用自己的版本繪製，所以只影響 app 內預覽 |
| Medium 3／Medium 4 的對應 | 依記憶：Medium 3＝橙（accent2）、Medium 4＝灰（accent3），因此把「Medium 4 (orange)」改成 Medium 3 | 應在 Excel 的表格樣式圖庫確認編號與顏色 |
| Docs 的框線、行號、分欄、頁面框線、樣式（Title／Quote 等）預設值 | 依 Word 預設範本數值 | 未以真 Word 的 PDF 輸出比對（可用 `test/_probe` 的 Word PDF 流程，見記憶 `reference-word-pdf-export-applescript`） |
| 頁面設定、主題的 ⌘Z 行為 | 以 Excel 行為為準撰寫 | 未在 Excel 實測 ⌘Z 是否一次一步 |

**怎麼做**
1. 準備同一份測試資料（Sheets：A1:D6 含表頭；Docs：數段文字）。
2. 在 Excel／Word 套用同一組選項，截圖或匯出 PDF；在 GenOffice 重做同樣操作。
3. 差異寫進 QA 文件；顏色差異集中改 `table-styles.ts` 的 `ACCENTS`、`tint`、`shade` 與各 family 分支。
4. Word／Excel 基準與 Playwright 指令見記憶 `reference-office-apps-for-comparison`。

## 2. 稽核最後一輪沒有重跑完整的 7 個分頁

**現況**：最後的修正（Format as Table、頁面設定／主題復原、公式列隱藏）之後，只重跑了：

| 分頁 | 結果 |
|---|---|
| Home | 剩 Format Painter（稽核假象：只是啟動格式刷，沒有改動）與 Insert row（單獨測試乾淨，原因未查明，判斷為序列殘留） |
| Page Layout | 剩 Themes › Office；之後加了主題復原，但**沒有再跑這個分頁**，只靠 `e2e/sheets-ribbon-fixes.spec.ts` 的主題測試驗證 |
| View | 公式列隱藏的主控台錯誤已用探針確認消失；稽核本身沒在修正後重跑 |
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

| 缺口 | 現況 | 建議位置 |
|---|---|---|
| Fill | 只有 Down、Right，沒有 Up、Left、Series… | `ExcelShell.tsx` Fill 選單；`ribbon-actions.ts` 的 `fill-down`／`fill-right` 旁新增 |
| Clear | 只有 All、Formats、Contents，沒有 Clear Comments、Clear Hyperlinks | 同上，Clear 選單 |
| 選擇性貼上 | Values、Formulas、Formats、Column Widths、All Except Borders；**沒有轉置**、沒有 Values & Number Formatting、沒有選擇性貼上對話框 | `ribbon-actions.ts` 的 `paste-special:*` |
| 儲存格群組 | 四顆獨立的插入／刪除列／欄按鈕；插入儲存格、刪除儲存格（向右／下／左／上移動）的對話框只能用 ⌃⇧= 與 ⌘- 叫出，功能區沒有下拉；下拉裡沒有插入／刪除工作表 | `ExcelShell.tsx` Cells 群組；對話框已有（`InsertCellsDialog.tsx`） |
| Format as Table | 只有 6 種（Excel 60 種）、沒有 New Table Style、沒有 Table Design 分頁 | `ExcelShell.tsx` 選項清單；外觀見 `table-styles.ts`（目前支援 Light 1–14、Medium 1–7、Dark 1–7） |
| 版面配置 → 主題 | 沒有帶即時預覽的圖庫 | `page-layout-actions.ts` 的 `theme*` |

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
