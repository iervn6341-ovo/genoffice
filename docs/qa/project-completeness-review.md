# GenOffice 專案完整性整理與確認

依據 `.agent/skills/genoffice-qa/`（SKILL.md、`references/`、`testing/`）逐項整理：每一條功能是否已實作、有哪些自動化測試、生命週期（復原／重做／存檔重開）是否有測試覆蓋，並以**實際執行**的測試結果確認。

- 基準版本：`qa/local-fixes` @ `9c235d2`，程式碼等同 `f3677de`（`9c235d2` 只新增稽核文件）。
- 執行環境：雲端 Linux 容器，xvfb。沒有 macOS、沒有第二螢幕、沒有 Microsoft Office。
- 前一份文件 `skill-feature-audit.md` 只檢查程式碼是否存在。本文件加上測試覆蓋與實際執行結果。

## 狀態用語

依 skill 的 Status Rules：

- **PASS**：本次實際執行、覆蓋該條目的自動化測試全部通過。
- **PARTIALLY TESTED**：有執行測試，但只覆蓋部分子項（例如只有單元測試，沒有 UI 流程，或沒有存檔重開）。
- **NOT TESTED**：功能有實作，但本次沒有能驗證它的自動化測試，或測試被略過。
- **BLOCKED**：環境無法執行（例如需要第二螢幕或 Office）。
- **FAIL**：執行後失敗且屬於產品問題。本次沒有這類結果。

實作欄：有＝找到實作；部分＝功能不完整；缺＝找不到。「e2e」指 `e2e/*.spec.ts`，「單元」指各 app 或 package 的 `tests/`。

---

## 1. 本次實際執行結果

| 項目                                      | 結果                                                                                                               |
| ----------------------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| `npm run build:all`                       | 通過                                                                                                               |
| `npm run typecheck`                       | 通過                                                                                                               |
| `npm run lint`                            | 通過（0 錯誤、13 警告，都是既有的）                                                                                |
| `npm run format:check`                    | 通過                                                                                                               |
| 單元測試（24 個 workspace，共 12,660 項） | 12,557 通過、11 失敗、92 略過；失敗全部查明原因，見 §1.1。略過中 67 項是 html2docx，設定 Chromium 後通過（下一列） |
| html2docx（設定 `CHROME_PATH`）           | 79/79 通過                                                                                                         |
| Rust sidecar `cargo test`                 | 189/189 通過                                                                                                       |
| Markdown 瀏覽器往返測試                   | 4/4 通過                                                                                                           |
| E2E 全套（174 項）                        | 163 通過、6 略過、5 失敗；5 項失敗都是 `docs-visual` 像素基準，見 §1.2                                             |

### 1.1 單元測試失敗逐項查證（全部是環境因素）

| 失敗                                                           | 原因                                                         | 查證方式                             | 查證結果                    |
| -------------------------------------------------------------- | ------------------------------------------------------------ | ------------------------------------ | --------------------------- |
| sheets `promote-file-atomically` ×3                            | 容器以 root 執行，會無視檔案鎖定與權限                       | 改用非 root 使用者（nobody）執行     | 6/6 通過                    |
| electron-utils `default-save-dir` ×1                           | 同上（root 會無視「資料夾不可寫」）                          | 改用 nobody 執行                     | 9/9 通過                    |
| markdown `close-asset-cleanup` ×1                              | 同上（測試用 `chmod 0` 讓檔案不可讀，root 無視）             | 改用 nobody 執行                     | 7/7 通過                    |
| sheets `pivot-roundtrip.e2e` ×1                                | 容器沒有安裝 LibreOffice                                     | 安裝 `libreoffice-calc-nogui` 後重跑 | 1/1 通過                    |
| font-metrics `font-covering` ×1、pdf `text-insert-fallback` ×1 | 容器裝有 GNU Unifont，連未指派的碼位（U+0378）都有字形       | 暫時移開 Unifont 後重跑，之後已放回  | 4/4、4/4 通過               |
| cli `mcp-http` ×3                                              | 沙箱代理攔截 `fetch` 到 `127.0.0.1`，回應「request blocked」 | 環境規定不能繞過代理                 | **BLOCKED**，無法在本機驗證 |

另外有一個產品面的觀察：`findFontCovering` 在裝有 Unifont 的 Linux 上會選中 Unifont。Unifont 是點陣風格的後備字型，用它嵌入 PDF 文字不理想。建議評估是否排除這類「涵蓋全部碼位」的字型。這只是觀察，本次沒有修改。

### 1.2 E2E 失敗與略過

- **`docs-visual` ×5（像素比對）**：容器的字型組合與 CI runner 不同。我補裝了 CI 用的 Carlito、Caladea、Noto CJK，差異仍在（2,258–7,745 像素，約 1%）。
  - 關鍵查證：在 `69ab301`（最後一次 CI 全綠的版本）重建後跑同一個 spec，五頁的差異像素數和目前 HEAD **完全相同**。
  - 結論：`f89c960` 之後的變更沒有改變這些頁面的渲染。失敗純屬容器字型差異。
  - 狀態：**BLOCKED**（需要 CI runner 或本機 Mac 驗證）。
- **略過 ×6**：5 項需要第二螢幕（簡報者檢視、雙螢幕放映），1 項需要原生拼字檢查。CI 也略過同樣 6 項。
  - 我嘗試用 Xvfb Xinerama 與 `xrandr --setmonitor` 模擬雙螢幕，Electron 仍只看到一個螢幕。
  - 狀態：**BLOCKED**（需要接第二螢幕的 Mac）。

### 1.3 CI 狀態提醒

- CI 只在 PR 到 `main` 或推送到 `main`／`dev_*`／`release_*` 時執行。
- `f89c960`、`e56b130`、`f3677de`、`9c235d2` 是在 PR #1 合併後才推上 `qa/local-fixes`，**從未跑過 CI**。
- 本次在容器內的全套執行是它們唯一的自動化驗證。它們要進 `main` 需要另開 PR，CI 會在那時跑。

---

## 2. Word（Docs）

| 條目                                                                | 實作                           | 自動化測試（本次執行皆通過）                                                                                                                                                               | 生命週期                        | 狀態                                                |
| ------------------------------------------------------------------- | ------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | ------------------------------- | --------------------------------------------------- |
| W-01 文件生命週期（新建、開啟、儲存、另存、改名、未存警告）         | 有                             | 單元：open-file、save-until-persisted、auto-file-name、atomic-write、external-change、doc-dirty；e2e：new-file-tab、mcp-visible-doc、mcp-open-documents                                    | 存檔重開：有                    | PASS                                                |
| W-02 基本編輯（Enter、Backspace、Shift+Enter、多行貼上、CJK）       | 有                             | 單元：tab-key-insert、enter-overtype-para-format、paste-text-marks、cjk-punct-shrink；e2e：docs-word-typing                                                                                | 部分                            | PARTIALLY TESTED（emoji、長段落沒有專門測試）       |
| W-03 插入點與選取（Home/End、雙擊、全選）                           | 有                             | e2e：Home/End、雙擊行尾（docs-word-typing）；單元：dom-range、inactive-selection、all-selection-enter                                                                                      | N/A                             | PASS（Mac 行為：NOT TESTED）                        |
| W-04 字元格式（B/I/U/S、字型、字級、顏色、醒目提示、上下標）        | 有                             | 單元：effective-format-toggle、caret-marks、format-off-clear、set-format-then-type、run-*、char-spacing-zero；e2e：docs-style-bold-toggle                                                  | 復原、重做、存檔重開：有（e2e） | PASS                                                |
| W-05 段落格式（對齊、縮排、首行／凸排、行距、清單、核取方塊、巢狀） | 有                             | 單元：indent、direction、numbering、list-enter-continuation、list-indent-live、checkbox-toggle、char-indent-save；e2e：Tab/Backspace 縮排、段落對話框、清單 AutoFormat、Continue Numbering | 復原：有                        | PASS                                                |
| W-06 樣式與標題                                                     | 部分（沒有 Title／Subtitle）   | 單元：character-styles、style-shortcut、style-gallery-overflow；e2e：Heading 1 流程                                                                                                        | 存檔重開：有                    | PARTIALLY TESTED                                    |
| W-07 表格（插入、Tab、列欄增刪、合併分割、調整大小、框線）          | 有                             | 單元：table-ops、native-table、nested-table-edit、table-borders、table-properties 等 20 多個；e2e：表格移動控點                                                                            | 存檔：單元有                    | PASS（拖曳調整大小：NOT TESTED）                    |
| W-08 圖片（插入、移動、縮放、繞圖、替代文字）                       | 部分（替代文字沒有使用者介面） | 單元：image-drop、image-wrap、image-rotation、picture-dialogs；e2e：拖曳圖片存檔重開                                                                                                       | 存檔重開：有                    | PARTIALLY TESTED                                    |
| W-09 連結                                                           | 有                             | 單元：link-edit、autolink、hyperlink-rels-save                                                                                                                                             | 存檔：有                        | PARTIALLY TESTED（沒有 e2e）                        |
| W-10 分頁、分節、頁首頁尾、頁碼、邊界、方向                         | 有                             | 單元：page-break*、page-margins、section-break-local、layout-history、hf-* 10 多個；e2e：docs-section-break-undo、docs-layout-undo                                                         | 復原：有                        | PASS                                                |
| W-11 剪貼簿（格式貼上、網頁貼上）                                   | 有                             | 單元：paste-options、paste-web-html、para-paste-roundtrip、font-paste-roundtrip；e2e：docs-paste-options、docs-web-paste-font                                                              | 部分                            | PARTIALLY TESTED（跨文件貼上：NOT TESTED）          |
| W-12 快捷鍵                                                         | 有                             | 單元：shortcut-registry、word-shortcut-actions、page-break-shortcut；e2e 用了 ⌘Z、⌘A、⌘B、⌘S                                                                                               | N/A                             | PASS                                                |
| W-13 復原／重做壓力序列                                             | 有                             | 部分涵蓋（各功能的復原測試）                                                                                                                                                               | N/A                             | NOT TESTED（沒有 skill 要求的長序列壓力測試）       |
| W-14／W-15 儲存重開、DOCX 往返                                      | 有                             | docx-engine 1,398 項；convert-regressions、raw-ppr                                                                                                                                         | 有                              | PASS（用 Word 開啟：BLOCKED）                       |
| W-16 異常操作                                                       | 有                             | 單元：encrypted-fixtures、docx-encryption                                                                                                                                                  | N/A                             | PARTIALLY TESTED                                    |
| W-17 大文件效能                                                     | 有（prosemirror-perf）         | 單元：prosemirror-perf、line-factor-incremental                                                                                                                                            | N/A                             | NOT TESTED（沒有量測延遲與記憶體）                  |
| W-18 跨功能流程                                                     | —                              | 沒有端到端的完整報告流程測試                                                                                                                                                               | —                               | NOT TESTED                                          |
| CJK 項目（IME、2 字元縮排、新細明體）                               | 有                             | e2e：段落對話框（字元單位）；單元：cjk-*                                                                                                                                                   | —                               | PARTIALLY TESTED（IME、zh-TW 介面實測：NOT TESTED） |

## 3. Excel（Sheets）

| 條目                                                                               | 實作           | 自動化測試                                                                                                                                              | 生命週期              | 狀態                                                        |
| ---------------------------------------------------------------------------------- | -------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------- | ----------------------------------------------------------- |
| X-01 活頁簿生命週期                                                                | 有             | 單元：close-guard、save-as-toolbar、save-edits-*、save-recovery-mode；e2e：sheets-new-blank、sheets-edit-save                                           | 存檔重開：有          | PASS                                                        |
| X-02 工作表增刪、改名、複製、排序                                                  | 有             | 單元：sheet-nav、sheet-rename-fix、xlsx-sheets、workbook-dsl                                                                                            | 存檔：有              | PARTIALLY TESTED（沒有 UI e2e）                             |
| X-03 各型別輸入                                                                    | 有             | 單元：number-format、short-date、numfmt-*、multiline-cells、long-text-render                                                                            | —                     | PASS                                                        |
| X-04 編輯模式（Enter、Esc、覆寫）                                                  | 有（Univer）   | 只有間接覆蓋（mcp-visible-sheet 填值）                                                                                                                  | —                     | NOT TESTED                                                  |
| X-05／X-06 選取與鍵盤導覽                                                          | 有             | 單元：excel-jump-nav、excel-shortcuts、goto、selection-*；e2e：arrow-collapse、jump-scroll                                                              | —                     | PARTIALLY TESTED（Tab、Enter 導覽是 Univer 內建，沒有測試） |
| X-07 公式與重算                                                                    | 有             | 單元：formula-* 15 多個、function-catalog、array-formulas、rate-function；e2e：mcp-sheet-values                                                         | 存檔：有（快取值 C3） | PASS                                                        |
| X-08 參照（相對、絕對、跨表、結構變動）                                            | 有             | 單元：formula-shift、xlsx-structure、shared-formula-*；e2e：sheets-paste-shared-formula-save                                                            | 存檔：有              | PASS                                                        |
| X-09 公式錯誤                                                                      | 有             | 單元：error-checking、error-value-align、formula-audit                                                                                                  | —                     | PARTIALLY TESTED（**循環參照完全沒有測試**）                |
| X-10 列欄增刪、大小、隱藏                                                          | 有             | 單元：xlsx-structure、structural-delete-guard、row-height-_、column-width-_；e2e：sheets-move-rows、sheets-menu-input-enter、sheets-insert-cells-dialog | 復原：有              | PASS                                                        |
| X-11 合併、取消合併                                                                | 有             | 單元：edit-journal、xlsx-structure（merge 部分）                                                                                                        | 存檔：單元有          | PARTIALLY TESTED（沒有 e2e）                                |
| X-12 格式                                                                          | 有             | 單元：format-cells、xlsx-borders、xlsx-alignment-carry、shrink-to-fit、vertical-alignment-import；e2e：sheets-excel-shortcuts                           | 存檔：有              | PASS                                                        |
| X-13 剪貼簿                                                                        | 有             | 單元：clipboard-_、copy-_、picture-paste、filtered-copy；e2e：sheets-paste-anchor-tile                                                                  | —                     | PASS（跨活頁簿貼上：NOT TESTED）                            |
| X-14 填滿與數列                                                                    | 有             | 單元：fill-range、bulk-fill-undo、flash-fill                                                                                                            | 復原：有              | PARTIALLY TESTED（拖曳填滿 UI：NOT TESTED）                 |
| X-15 排序與篩選                                                                    | 有             | 單元：sort-range、advanced-filter、xlsx-filter；e2e：filter-reopen、filter-outline                                                                      | 存檔重開：有          | PASS                                                        |
| X-16 尋找與取代                                                                    | 有             | 單元：lazy-find、sparse-find、replace-autosearch、workbook-search                                                                                       | —                     | PARTIALLY TESTED（沒有 e2e）                                |
| X-17 凍結、縮放                                                                    | 有             | 單元：xlsx-freeze、sheet-zoom-scale、zoom-to-selection；e2e：凍結復原後存檔                                                                             | 有                    | PASS                                                        |
| X-18 復原壓力                                                                      | 有             | 單元：undo-carry、edit-journal、ai-undo-budget                                                                                                          | —                     | NOT TESTED（沒有 skill 的長序列）                           |
| X-19／X-20 儲存重開、XLSX 往返                                                     | 有             | xlsx-* 40 多個、Rust 189 項；e2e：sheets-edit-save、sheets-xlsm                                                                                         | 有                    | PASS（Excel 開啟無修復提示：BLOCKED）                       |
| X-21 大檔效能                                                                      | 有（串流載入） | 單元：lazy-plan、univer-range-loading、recalc-size-gate、formula-cost                                                                                   | —                     | NOT TESTED（沒有量測）                                      |
| X-22 無效操作                                                                      | 有             | 單元：structural-delete-guard、number-as-text-alert、streamed-formula-guard                                                                             | —                     | PASS                                                        |
| X-23 跨功能流程                                                                    | —              | 無                                                                                                                                                      | —                     | NOT TESTED                                                  |
| Sheets 筆記的互通檢查（CSV、CF、註解、超連結、Unicode 名稱、date1904、快取 `<v>`） | 有             | 每項都有對應單元測試；e2e：sheets-csv-*                                                                                                                 | 存檔：有              | PASS                                                        |

## 4. PowerPoint（Slides）

| 條目                              | 實作             | 自動化測試                                                                                                             | 生命週期     | 狀態                                             |
| --------------------------------- | ---------------- | ---------------------------------------------------------------------------------------------------------------------- | ------------ | ------------------------------------------------ |
| P-01 簡報生命週期                 | 有               | 單元：save-serialization、headless-export；e2e：mcp-visible-deck                                                       | 存檔：有     | PASS                                             |
| P-02 投影片新增、複製、刪除、重排 | 有；**多選：缺** | 單元：history、regenerate-slide、apply-ops-tool（操作層）                                                              | 復原：單元有 | PARTIALLY TESTED（縮圖面板 UI：NOT TESTED）      |
| P-03 版面配置                     | 有               | 單元：layout-tools；e2e：Slides 群組的 Layout 列                                                                       | —            | PARTIALLY TESTED                                 |
| P-04 文字方塊與編輯模式           | 有               | 單元：edit-flow、edit-selection、text-hit-area；e2e：slides-text-editing、Home 分頁插入文字方塊                        | 復原：有     | PASS                                             |
| P-05 圖形（填滿、框線）           | 有               | 單元：draw-shape、insert-gallery；e2e：Shapes 圖庫、Drawing extras                                                     | —            | PASS（旋轉 UI：NOT TESTED）                      |
| P-06 圖片（裁剪、比例）           | 有               | 單元：picture-srcrect-inset、cutout、image-loader                                                                      | —            | PARTIALLY TESTED（沒有 e2e）                     |
| P-07 表格                         | 有               | 單元：table-hit、table-cell-edit-formatting                                                                            | —            | PARTIALLY TESTED                                 |
| P-08 物件選取                     | 有               | 單元：selection-chrome、edit-frame-press                                                                               | —            | PARTIALLY TESTED（框選、Shift 多選：NOT TESTED） |
| P-09 移動、縮放、旋轉             | 有               | 單元：snap、stage-refit-follow                                                                                         | —            | PARTIALLY TESTED（沒有數值化幾何的 e2e）         |
| P-10 對齊與均分                   | 有               | 單元：arrange-ops（alignElements、distributeElements）                                                                 | —            | PASS（單元層）                                   |
| P-11 圖層順序                     | 有               | 單元：edit-ops                                                                                                         | —            | PARTIALLY TESTED                                 |
| P-12 群組、取消群組               | 有               | 單元：只在 op-docs、durable-identity 間接出現                                                                          | —            | NOT TESTED                                       |
| P-13 文字格式                     | 有               | 單元：font-size-step、font-extras-editing、text-highlight；e2e：slides-font-styles（9 項）、slides-home-extras（8 項） | 復原：有     | PASS                                             |
| P-14 剪貼簿                       | 有               | 單元：paste-cascade                                                                                                    | —            | PARTIALLY TESTED                                 |
| P-15 快捷鍵                       | 有               | 單元：keyboard-actions；e2e：⌘B/⌘I/⌘U                                                                                  | —            | PASS                                             |
| P-16 備註                         | 有               | e2e：slides-notes 2 項通過（第 3 項需要第二螢幕，略過）                                                                | 換頁保留：有 | PARTIALLY TESTED                                 |
| P-17 放映                         | 有               | 單元：show-keys、presenter-display；e2e：放映按鍵 N/P/B/W                                                              | —            | PARTIALLY TESTED（簡報者檢視 5 項：BLOCKED）     |
| P-18 佈景主題與背景               | 有               | 單元：theme、svg-theme-retint；e2e：theme-visual-slides                                                                | —            | PASS                                             |
| P-19 復原壓力                     | 有               | 單元：history、undo-routing                                                                                            | —            | NOT TESTED（沒有 skill 的長序列）                |
| P-20／P-21 儲存重開、PPTX 往返    | 有               | pptx-engine 1,025 項、pptx-render 275 項；單元：edit-fidelity                                                          | 有           | PASS（用 PowerPoint 開啟：BLOCKED）              |
| P-22 效能                         | —                | 無                                                                                                                     | —            | NOT TESTED                                       |
| P-23 異常操作                     | 有               | 無專門測試                                                                                                             | —            | NOT TESTED                                       |
| P-24 跨功能流程                   | —                | 無                                                                                                                     | —            | NOT TESTED                                       |

## 5. PDF、Markdown、HTML

| 條目                        | 實作                       | 自動化測試                                                                                    | 狀態                                                 |
| --------------------------- | -------------------------- | --------------------------------------------------------------------------------------------- | ---------------------------------------------------- |
| F-01 開啟、縮放、大綱、縮圖 | 有                         | e2e：pdf-fit-zoom、pdf-outline-layout；單元：pdf-thumb、spread、view-state                    | PASS（旋轉或裁剪頁的顯示：PARTIALLY TESTED）         |
| F-02 文字編輯與插入         | 有                         | e2e：pdf-edit 2 項；單元：text-edit、text-splice、text-wrap、text-block、text-insert-fallback | PASS                                                 |
| F-03 圖片                   | 有                         | e2e：pdf-edit 圖片；單元：cutout、image-bake                                                  | PASS                                                 |
| F-04 頁面操作               | 有                         | e2e：刪頁後匯入；單元：save-pdf（merge、N-up、CropBox）、edit-ops                             | PASS                                                 |
| F-05 註解與遮蔽             | 有                         | e2e：遮蔽的復原與重做；單元：redaction、annotations、annot-delete                             | PASS                                                 |
| F-06 裁剪、頁面大小、浮水印 | 有（中日韓浮水印不可搜尋） | e2e：裁剪對話框；單元：page-op-undo、save-pdf 浮水印                                          | PARTIALLY TESTED（裁剪復原沒有 e2e）                 |
| F-07 儲存與互通             | 有                         | 單元：save-pdf、autosave、page-operation-atomic-write                                         | PASS（用 Preview 開啟：BLOCKED）                     |
| M-01 Markdown 編輯          | 有                         | e2e：markdown-tab（雙擊、粗體）；單元：markdown-nodes                                         | PARTIALLY TESTED（Markdown 的 Home/End：NOT TESTED） |
| M-02 原始碼逐位元保留       | 有                         | 瀏覽器往返 4 項；單元：roundtrip-serializer、source-splice                                    | PASS                                                 |
| M-03 匯出                   | 有                         | 單元：docx-export、print-html                                                                 | PARTIALLY TESTED                                     |
| M-04 復原與搜尋醒目提示     | 有                         | e2e：Ctrl+F 取代全部；單元：find-panel                                                        | PARTIALLY TESTED（編輯後清除醒目提示：NOT TESTED）   |
| H-01 預覽編輯               | 有                         | e2e：html-tab（inspector 雙擊編輯）                                                           | PASS                                                 |
| H-02 插入與拖曳             | 有                         | e2e：html-insert-drag                                                                         | PASS                                                 |
| H-03 原始碼往返             | 有                         | e2e：未編輯存檔逐位元相同；單元：document、ops                                                | PASS                                                 |
| H-04 HTML → DOCX            | 有                         | html2docx 79 項                                                                               | PASS（用 Word 開啟：BLOCKED）                        |

## 6. 共用基礎（shared-office）

| 項目                         | 狀態             | 說明                                                                                         |
| ---------------------------- | ---------------- | -------------------------------------------------------------------------------------------- |
| 復原與重做                   | PARTIALLY TESTED | 各應用程式都有單元與 e2e 覆蓋；skill 要求的「A→B→復原→C 清除重做分支」沒有專門測試           |
| 剪貼簿                       | PARTIALLY TESTED | 跨文件、跨活頁簿貼上沒有測試                                                                 |
| 存檔、自動儲存、匯入匯出     | PASS             | 六個應用程式都有，並有對應測試                                                               |
| 主題（淺色、深色、跟隨系統） | PASS             | e2e：theme-pipeline、theme-visual（5 項）、theme-visual-slides；CI 另有 `check-theme-colors` |
| 執行期錯誤檢查               | NOT TESTED       | e2e 沒有統一收集 console error                                                               |

---

## 7. 缺口總表（依優先順序）

| #   | 缺口                                                                    | 類型         | 建議                                                          |
| --- | ----------------------------------------------------------------------- | ------------ | ------------------------------------------------------------- |
| 1   | `f89c960`、`e56b130`、`f3677de`、`9c235d2` 從未跑過 CI                  | 流程         | 開 PR 到 `main`，讓 CI 跑一次                                 |
| 2   | 簡報者檢視與雙螢幕放映（5 個 e2e）                                      | 環境 BLOCKED | 在接第二螢幕的 Mac 上跑                                       |
| 3   | `docs-visual` 像素基準                                                  | 環境 BLOCKED | 由 CI 或 Mac 驗證（本次已確認新 commit 不影響渲染）           |
| 4   | 循環參照沒有任何測試                                                    | 測試缺口     | 補一個 e2e 或整合測試：輸入 `=A1` 到 A1，確認顯示錯誤且不卡住 |
| 5   | skill 要求的長序列復原壓力（W-13、X-18、P-19）                          | 測試缺口     | 各補一個 e2e                                                  |
| 6   | 跨功能流程（W-18、X-23、P-24）                                          | 測試缺口     | 各補一個 e2e（報告、費用表、商業簡報）                        |
| 7   | 投影片多選                                                              | 功能缺       | 決定是否支援（清單標為 where supported）                      |
| 8   | Title／Subtitle 樣式、圖片替代文字介面                                  | 功能部分     | 決定是否補齊                                                  |
| 9   | 群組、取消群組（P-12）、Sheets 合併儲存格（X-11）、PDF 裁剪復原沒有 e2e | 測試缺口     | 補 e2e                                                        |
| 10  | 效能（W-17、X-21、P-22）沒有量測                                        | 測試缺口     | 加入效能基準測試                                              |
| 11  | Unifont 會被選為 PDF 後備字型                                           | 產品觀察     | 評估排除涵蓋全部碼位的字型                                    |
| 12  | Slides 選取框顏色沒有集中在色表（CLAUDE.md 規則 5）                     | 規範         | 抽成常數表                                                    |
| 13  | README 說 `GENOFFICE_DEBUG_HOOKS` 由 helpers 設定                       | 文件         | 修正文件，或改由 helpers 設定                                 |
| 14  | 所有 Office 比對                                                        | 環境 BLOCKED | 在 Mac 上依 `testing/*.md` 比對 Word、Excel、PowerPoint       |

## 8. 最終判定（依 completion-gate）

| 服務                 | 判定                        |
| -------------------- | --------------------------- |
| Word（Docs）         | PASS WITH KNOWN LIMITATIONS |
| Excel（Sheets）      | PASS WITH KNOWN LIMITATIONS |
| PowerPoint（Slides） | PARTIALLY TESTED            |
| PDF                  | PASS WITH KNOWN LIMITATIONS |
| Markdown             | PASS WITH KNOWN LIMITATIONS |
| HTML                 | PASS                        |

- **PASS WITH KNOWN LIMITATIONS**：已執行的測試全部通過；限制在 §7。
- Slides 評為 PARTIALLY TESTED，因為群組、多選、簡報者檢視、效能、異常操作都沒有被驗證。
- 本次**沒有**發現新的產品失敗。所有失敗都已查證為環境因素，只有 cli `mcp-http` 3 項因代理限制無法在本機驗證。
- 真實使用者的互動、Mac 快捷鍵、以及與 Microsoft Office 的比對，本次都**沒有**直接觀察，全部列為 NOT TESTED 或 BLOCKED。
