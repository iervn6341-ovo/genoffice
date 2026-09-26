# genoffice-qa 技能指引 × 原始碼對照稽核

- 基準：`qa/local-fixes` @ `f3677de`（`add .agent`），2026-09-26。
- 來源：`.agent/skills/genoffice-qa/`（與 `.claude/skills/genoffice-qa/` 內容相同）。
- 方法：**只檢查原始碼**（grep 與閱讀）。沒有執行任何測試，也沒有與 Office 比對。
  - 「有」＝找到實作該功能的程式碼。
  - 「部分」＝有相近功能但不完整。
  - 「未找到」＝原始碼中找不到。
  - 「流程」＝該條是測試流程要求，不是產品功能。
- 所有「有」都只代表**程式碼存在**，不代表行為已驗證。

SKILL.md 本身是開發與測試流程（Phase 1–9、狀態規則），沒有功能條目。功能條目在 `references/*.md` 與 `testing/*.md`，以下依序對照。

---

## Word（Docs）— `testing/word.md` W-01…W-18、`references/word.md`

| 條目                                  | 狀態     | 證據                                                                                                                            |
| ------------------------------------- | -------- | ------------------------------------------------------------------------------------------------------------------------------- |
| W-01 新建／開啟／儲存／另存           | 有       | `apps/docs/src/renderer/App.tsx:1938` openFile；`file-actions.ts:800` saveImpl(saveAs)                                          |
| W-01 多文件、最近文件                 | 有       | shell 分頁（`apps/shell/src/main/index.ts` docsTabs）；`apps/shell/src/shared/home-api.ts:65` RecentEntry                       |
| W-01 重新命名                         | 有       | `App.tsx:1878` onRenamedDocx                                                                                                    |
| W-01 未儲存警告                       | 有       | `apps/shell/src/main/index.ts:2734` 視窗 close 攔截所有有未存變更的分頁                                                         |
| W-02 Enter／Backspace／Delete／插入   | 有       | ProseMirror 基本編輯；`extensions.ts:1246` WordSelectAllDelete                                                                  |
| W-02 Shift+Enter                      | 有       | `extensions.ts:1232` Shift-Enter → hardBreak                                                                                    |
| W-02 多行貼上                         | 有       | `editor/paste-text.ts:71` pasteTextSlice                                                                                        |
| W-02 Unicode／CJK／emoji              | 有       | CJK 排版：`protected-render.ts`、`cjk-punct-shrink.ts`                                                                          |
| W-03 點擊、方向鍵、全選、Shift+方向鍵 | 有       | ProseMirror 內建；全選刪除 `extensions.ts:1246`                                                                                 |
| W-03 Home/End（Mac）                  | 有       | `packages/ui/src/mac-line-keys.ts` handleMacLineBoundaryKey                                                                     |
| W-03 選字／選段落（雙擊／三擊）       | 有       | ProseMirror 內建；雙擊行尾 `editor/double-click-line-end.ts:18`                                                                 |
| W-04 粗體／斜體                       | 有       | `editor/marks.ts`；有效格式 `editor/effective-format.ts`（含樣式繼承）                                                          |
| W-04 底線／刪除線                     | 有       | `marks.ts:75`、`marks.ts:94`；Mod-u `marks.ts:89`                                                                               |
| W-04 字型／字級／顏色／醒目提示       | 有       | docTextStyle mark：`marks.ts`（sizeHalfPoints、font、color、highlight）；`editor/text-color.ts`                                 |
| W-04 上標／下標                       | 有       | `marks.ts:559`、`convert.ts:3068` vertAlign                                                                                     |
| W-04 游標處輸入沿用格式               | 有       | `editor/caret-marks.ts:116` CaretMarksMemory                                                                                    |
| W-05 對齊                             | 有       | `editor/direction.ts:75` setSelectionAlign                                                                                      |
| W-05 縮排／首行／凸排                 | 有       | `editor/indent.ts:10`；段落對話框 `components/ContextMenu.tsx:753` SpecialIndent                                                |
| W-05 行距／段距                       | 有       | `extensions.ts:221/231` lineSpacing、spaceBefore                                                                                |
| W-05 項目符號／編號／巢狀             | 有       | docListItem（kind、numId、ilvl）；`editor/numbering.ts`                                                                         |
| W-05 核取方塊清單                     | 有       | `editor/checkbox-toggle.ts`                                                                                                     |
| W-06 內文、標題 1–3                   | 有       | `components/Ribbon.tsx:573` STYLE_GALLERY                                                                                       |
| W-06 Title／Subtitle                  | **部分** | 內建樣式庫只有內文與標題 1–3；空白範本（`packages/docx-engine/src/blank.ts`）也沒有 Title／Subtitle。只有文件本身定義時才能套用 |
| W-06 自訂樣式                         | 有       | styleUpserts（`doc-dirty.ts:22`、`ai/style-ops.ts`）                                                                            |
| W-07 插入表格、Tab 移動               | 有       | `extensions.ts:3179` goToNextCell（最後一格 Tab 新增列 `:3182`）                                                                |
| W-07 插入／刪除列欄、合併／分割       | 有       | `extensions.ts:23` addRowAfter 等；`ContextMenu.tsx:16/18` mergeCells、splitCell                                                |
| W-07 調整大小、框線                   | 有       | `editor/table-sizing.ts:170`；columnResizing；`editor/table-properties.ts`                                                      |
| W-08 圖片插入、移動、縮放、繞圖、對齊 | 有       | `editor/image-drop.ts`；imageOffsetXEmu、imageWrap（`convert.ts:697-719`）                                                      |
| W-08 替代文字                         | **部分** | 只找到 AI 工具參數 `ai/floating-ops.ts:206` altText。沒有找到供使用者編輯替代文字的介面                                         |
| W-08 標號（caption）                  | 有       | SEQ 欄位標號渲染 `extensions.ts:4390`（References 分頁）                                                                        |
| W-09 連結建立、編輯、移除             | 有       | `marks.ts:112` link mark；`components/ribbon-insert-tab.tsx:563` 連結對話框                                                     |
| W-10 分頁、分節符號                   | 有       | `editor/page-break.ts`；`App.tsx` insertSectionBreak（C6）                                                                      |
| W-10 頁首頁尾、頁碼                   | 有       | `editor/hf-text.ts:16` {PAGE}/{NUMPAGES}；`App.tsx:2354` applyPgNumFormat                                                       |
| W-10 邊界、方向、紙張大小             | 有       | Layout 分頁 onSection；已納入復原 `editor/layout-history.ts`                                                                    |
| W-11 剪貼簿、格式貼上、跨文件         | 有       | `editor/paste-options.ts`、`paste-web-html.ts`、`components/PasteOptionsChip.tsx`                                               |
| W-12 ⌘C／X／V／Z／A／B／I／U／S、重做 | 有       | UndoRedo：Mod-z、Shift-Mod-z、Mod-y（`@tiptap/extensions`）；⌘S `App.tsx:4596`；⌘T 凸排 `App.tsx:4756`                          |
| W-13 復原／重做                       | 有       | `extensions.ts:3` UndoRedo（`@tiptap/extensions`）                                                                              |
| W-14／W-15 儲存重開、DOCX 往返        | 有       | `packages/docx-engine`（parseDocx／saveDocx）；`editor/convert.ts` pmDocToSavePlan                                              |
| W-16 異常匯入                         | 有       | `file-actions.ts:457` appOpenFailed 提示，不崩潰                                                                                |
| W-17 大文件效能                       | 有       | `editor/prosemirror-perf.ts:117` installProseMirrorPerf                                                                         |
| W-18 跨功能流程                       | 流程     | 由上列功能組成                                                                                                                  |
| 筆記：段首 Tab／Backspace 縮排        | 有       | `editor/decoration-extensions.ts`（f89c960）；spec `docs-word-typing.spec.ts:227`                                               |
| 筆記：CJK 定位點 480／420、字元單位   | 有       | `packages/docx-engine/src/blank.ts:153`；`ContextMenu.tsx:812`                                                                  |
| 筆記：清單 AutoFormat、⌘Z 還原標記    | 有       | `editor/list-autoformat.ts`                                                                                                     |
| 筆記：IME 組字                        | 有       | `direction.ts:144`、`revisions.ts:628` view.composing 處理                                                                      |
| 筆記：新細明體預設                    | 有       | `packages/docx-engine/src/parse-props.ts:1165` zh-tw → PMingLiU；`font-list.ts:37`                                              |
| 筆記：已知問題（字元單位存成 twips）  | 屬實     | `packages/docx-engine/src/generate.ts:1490` 寫 `w:firstLineChars="0"`，改用 twips                                               |

## Excel（Sheets）— `testing/excel.md` X-01…X-23

| 條目                                          | 狀態              | 證據                                                                                                                                   |
| --------------------------------------------- | ----------------- | -------------------------------------------------------------------------------------------------------------------------------------- |
| X-01 開啟、儲存、另存、關閉警告               | 有                | `renderer/save-actions.ts:94` handleSave；`main/sheets-main.ts:2257` requestSheetsClose                                                |
| X-02 新增、改名、複製、排序、刪除、切換工作表 | 有                | `ribbon-actions.ts:365` add_sheet；COPY_SHEET_COMMAND（`App.tsx:192`）；`shared/desktop-api.ts:1376/1388` remove-sheet、reorder-sheets |
| X-03 各型別輸入（日期、時間、%、布林）        | 有                | Univer 解析加 `renderer/numfmt-fix.ts` 修正                                                                                            |
| X-04 編輯模式、Enter／Esc                     | 有                | Univer 內建；`ribbon-actions.ts:1172` isCellEditing                                                                                    |
| X-05 選取列、欄、全選                         | 有                | Univer 內建；`renderer/excel-shortcuts.ts`（整列／整欄）                                                                               |
| X-06 方向鍵、Tab、Enter                       | 有（Univer 內建） | 應用程式沒有另寫，由 Univer 提供                                                                                                       |
| X-06 Home、Ctrl+Home/End、PageUp/Down         | 有                | `App.tsx:1724` registerExcelShortcuts、`excel-jump-nav.ts`；`ribbon-actions.ts:1161` page-row/col                                      |
| X-07 公式與重算                               | 有                | Univer 公式引擎；IronCalc 後備（`shared/desktop-api.ts:1022`）；`formula-closure.ts`                                                   |
| X-08 參照（含跨工作表）、複製調整             | 有                | Univer；存檔端 `packages/xlsx-gateway/src/gateway/xlsx-structure.ts` 參照平移                                                          |
| X-09 錯誤、循環參照                           | 有                | EXCEL_ERROR_LITERALS；循環參照由 Univer engine-formula 處理                                                                            |
| X-10 插入／刪除列欄、調整大小、隱藏           | 有                | `ribbon-actions.ts:1016` insert_rows 等；`app-constants.ts` AXIS_ATTR_MUTATIONS                                                        |
| X-10 插入／刪除儲存格（位移）                 | 有                | `renderer/cell-shift.ts`、`InsertCellsDialog.tsx`（C4）                                                                                |
| X-11 合併、取消合併                           | 有                | `ribbon-actions.ts:1403`；MERGE_MUTATIONS 寫入 journal                                                                                 |
| X-12 字型、填滿、框線、對齊、換行、數值格式   | 有                | `ribbon-actions.ts:1335/1345/1432`；`FormatCellsDialog.tsx`；`ExcelShell.tsx` NUMBER_FORMAT_SHORTCUTS                                  |
| X-13 剪貼簿、選擇性貼上、外部貼上             | 有                | `ribbon-actions.ts:343` paste-value；`clipboard-tsv.ts`；`copy-materialize.ts`；`clipboard-anchor-tile.ts`                             |
| X-14 填滿、數列                               | 有                | AUTO_FILL_COMMAND（`App.tsx:2592`）；`bulk-fill-undo.ts`                                                                               |
| X-15 排序（多欄）、篩選                       | 有                | `ribbon-actions.ts:1593` sort-custom；FILTER_COMMAND_PATTERN；`AdvancedFilterDialog.tsx`                                               |
| X-16 尋找、取代（含串流區塊外）               | 有                | `renderer/sparse-find.ts`；open-find-dialog                                                                                            |
| X-17 凍結、縮放                               | 有                | `ribbon-actions.ts:785` setFreeze；zoomPercent                                                                                         |
| X-18／X-19 復原壓力、儲存重開                 | 有                | edit-journal 與 gateway 存檔                                                                                                           |
| X-20 XLSX 往返                                | 有                | `packages/xlsx-gateway/src/gateway/xlsx-gateway.ts:570`；Rust sidecar `apps/sheets/native/xlsx-engine`                                 |
| X-21 大檔效能                                 | 有                | 串流載入：preloadComplete、loadVisibleRange                                                                                            |
| X-22 無效操作                                 | 有                | 各處 gating 與訊息（appNeedFullLoadSort 等）；#ERROR! 過濾 `save-actions.ts:208`                                                       |
| 筆記：CSV 分隔符、sep=、UTF-16                | 有                | `main/sheets-main.ts:2992`；`packages/xlsx-gateway/src/gateway/csv-import.ts:55`                                                       |
| 筆記：CF 不支援規則原樣保留                   | 有                | `xlsx-cf.ts:165` UNLOADABLE_CF_TYPES                                                                                                   |
| 筆記：未動的註解保留格式                      | 有                | noteStates（`shared/desktop-api.ts:1753`）                                                                                             |
| 筆記：範圍超連結                              | 有                | hyperlinkEdits（`xlsx-package-io.ts:90`）                                                                                              |
| 筆記：Unicode 名稱與共用公式                  | 有                | `apps/sheets/native/xlsx-engine/src/shared_formulas.rs`（整字元複製，含測試）                                                          |
| 筆記：date1904 驗證日期                       | 有                | `xlsx-dv.ts` DATE1904_OFFSET_DAYS（C2）                                                                                                |
| 筆記：公式快取 `<v>`                          | 有                | `renderer/formula-values.ts` journaledFormulaValues（C3）                                                                              |

## PowerPoint（Slides）— `testing/powerpoint.md` P-01…P-24

| 條目                                                   | 狀態       | 證據                                                                                                                                |
| ------------------------------------------------------ | ---------- | ----------------------------------------------------------------------------------------------------------------------------------- |
| P-01 儲存、另存、關閉警告                              | 有         | `renderer/file-actions.ts:80/105`；`main/slides-main.ts:627` requestSlidesClose                                                     |
| P-02 新增、複製、刪除、重排、切換投影片                | 有         | `renderer/slide-actions.ts:12` 等；縮圖拖曳 thumbDragProps                                                                          |
| P-02 投影片多選                                        | **未找到** | 縮圖與投影片瀏覽模式點擊只設 current（`App.tsx:3516` setCurrent）。只有自訂放映對話框有 slideIndices                                |
| P-03 版面配置                                          | 有         | `slide-actions.ts:27` addSlideWithLayout；`shared/ipc.ts:692` 版面清單                                                              |
| P-04 文字方塊與文字編輯模式                            | 有         | `renderer/TextEditOverlay.tsx`；插入 `RibbonInsertTab.tsx`                                                                          |
| P-05 圖形、填滿、框線、旋轉                            | 有         | `draw-shape.ts`；`arrange-actions.ts:218` rotateSelected；`FormatPane.tsx`                                                          |
| P-06 圖片、裁剪、比例                                  | 有         | `CropOverlay.tsx`；`SlideCanvas.tsx:1375` keepRatio                                                                                 |
| P-07 表格列欄、格式、調整                              | 有         | `renderer/table-actions.ts:11` tableStructureOp、`:38` 合併、`:61/:77` 調整                                                         |
| P-08 點選、框選、多選                                  | 有         | `SlideCanvas.tsx:750` marquee                                                                                                       |
| P-09 移動、縮放、旋轉、吸附                            | 有         | `SlideCanvas.tsx:1392` 等尺寸吸附；`snap.ts`                                                                                        |
| P-10 對齊、均分                                        | 有         | `arrange-actions.ts:55/63` alignSelected（含 distribute-h/v）                                                                       |
| P-11 圖層（上移、下移、最前、最後）                    | 有         | `arrange-actions.ts:159` reorderSelected                                                                                            |
| P-12 群組、取消群組                                    | 有         | `arrange-actions.ts:13/33`                                                                                                          |
| P-13 文字格式、項目符號、行距                          | 有         | `bullet-presets.ts`；`components/ribbon-shared.tsx:290` lineSpacing                                                                 |
| P-14 物件剪貼簿、跨投影片貼上                          | 有         | `clipboard-actions.ts`；`main/paste-cascade.ts`                                                                                     |
| P-15 快捷鍵（刪除、方向鍵微移、⌘D）                    | 有         | `keyboard-actions.ts:266/311/319`、`:298` ⌘D                                                                                        |
| P-16 備註                                              | 有         | `components/NotesEditor.tsx:60`；`main/notes-view.ts`                                                                               |
| P-17 放映（方向鍵、PageUp/Down、B/W、Esc、簡報者檢視） | 有         | `slideshow-utils.ts` showKeyCommand（`:152` B、`:156` W）；`main/presenter-show.ts`、`components/PresenterView.tsx`                 |
| P-18 佈景主題、背景                                    | 有         | `themes.ts`；`components/FormatBackgroundPane.tsx`                                                                                  |
| P-19 復原路由                                          | 有         | `undo-routing.ts`（文字與簡報層級分流）                                                                                             |
| P-20／P-21 儲存重開、PPTX 往返                         | 有         | `packages/pptx-engine/src/index.ts:686` savePptx                                                                                    |
| P-22 效能                                              | 流程       | 沒有對應的特定程式碼可查，需實測                                                                                                    |
| P-23 邊界操作                                          | 有         | 刪除最後一張的處理（`keyboard-actions.ts:252`）；`FormatPane.tsx:617` clamp                                                         |
| 筆記：⌘⇧> 字級、變更大小寫                             | 有         | `shared/ipc.ts:223/226` fontSizeStep、Change Case                                                                                   |
| 筆記：選取框顏色（CLAUDE.md 規則 5）                   | **部分**   | `SlideCanvas.tsx:276` selectionChromeColor 依投影片明暗取色，但色值直接寫在函式內，沒有集中在畫布色表（規則要求的 constants table） |

## PDF／Markdown／HTML — `testing/pdf-markdown-html.md`

| 條目                                                  | 狀態             | 證據                                                                                                    |
| ----------------------------------------------------- | ---------------- | ------------------------------------------------------------------------------------------------------- |
| F-01 縮放（符合寬度等）、大綱、縮圖                   | 有               | `apps/pdf/src/renderer/App.tsx:5975` fitWidth；`PdfThumb.tsx`                                           |
| F-02 文字編輯（字型、字級、顏色）、插入文字、重新量測 | 有               | `shared/ipc.ts:12` validateTextEdits；textInserts；`text-block.ts` reflowOverflows                      |
| F-03 圖片插入、取代、裁剪、去背、刪除                 | 有               | `shared/ipc.ts:364/390`；`ImageDialogs.tsx` CropDialog、CutoutDialog                                    |
| F-04 旋轉、刪除、重排、匯入頁                         | 有               | rotatePages／deletePage／setPageOrder 操作；`shared/ipc.ts:21` insertPdf                                |
| F-04 N-up 遵守 CropBox 與 /Rotate                     | 有               | `main/save-pdf.ts:635/652/673`                                                                          |
| F-05 螢光筆、繪圖、印章、遮蔽（含復原）               | 有               | `edit-state.ts`（redactions bucket）；`main/redaction.ts`；`StampDialog.tsx`                            |
| F-05 合併時攤平註解外觀                               | 有               | `main/save-pdf.ts:560` flattenAnnotAppearances                                                          |
| F-06 裁剪與頁面大小可復原                             | 有               | `main/page-op-history.ts`；restorePageOp（C8）                                                          |
| F-06 浮水印文字可搜尋                                 | 有（僅 WinAnsi） | `main/save-pdf.ts:1000` drawInvisibleStampText；中日韓文字仍只有圖片                                    |
| F-07 儲存、另存副本                                   | 有               | `shared/ipc.ts:42` saveAsRequest；savePdfToPath                                                         |
| M-01 Home/End、雙擊行尾、清單、標題、程式碼區塊       | 有               | `apps/markdown/src/renderer/main.tsx:23`；`editor/double-click-line-end.ts`；`editor/CodeBlockView.tsx` |
| M-02 原始碼逐位元保留                                 | 有               | `App.tsx:170` originalSourceRef（round-trip 模式）                                                      |
| M-03 匯出 DOCX／HTML／PDF                             | 有               | `shared/ipc.ts:28/29` exportDocx、exportPdf 等                                                          |
| M-04 搜尋醒目提示                                     | 有               | `editor/searchHighlight.ts`（編輯後是否清除未在程式碼中確認，需實測）                                   |
| H-01 預覽區雙擊編輯、Home/End、Esc 提交               | 有               | `apps/html/src/renderer/preview/inspector.js:911` dblclick、`:964` Home/End、`:981` Escape              |
| H-02 插入後可編輯（版本門檻）、拖曳                   | 有               | `App.tsx` editAfterLoadRef {sid, version}（0096771）                                                    |
| H-03 最小化原始碼差異                                 | 有               | `document/patch.ts`（Patch／PatchSet，所有編輯都編譯成原始碼區段替換）                                  |
| H-04 HTML → DOCX                                      | 有               | `packages/html2docx/src/convert.ts`                                                                     |

## 共用基礎（`references/shared-office.md`）與測試說明（`testing/README.md`）

| 條目                                                                                    | 狀態         | 證據或說明                                                                                                                                                                                                                                                                                                                                    |
| --------------------------------------------------------------------------------------- | ------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 復原／重做、剪貼簿、存讀檔、匯入匯出、縮放捲動                                          | 有           | 分見各應用程式                                                                                                                                                                                                                                                                                                                                |
| 自動儲存                                                                                | 有           | 六個應用程式都有 autoSave（docs、sheets、slides、markdown、html 的 ipc；pdf `App.tsx`）                                                                                                                                                                                                                                                       |
| launchShell 選項 lang／openFile／onboardingSeen／settings／foreground／env／userDataDir | 有           | `e2e/helpers.ts:26-40`（另有 chromiumArgs、videoDir）                                                                                                                                                                                                                                                                                         |
| 1280×800、背景模式、GENOFFICE_E2E_FOREGROUND                                            | 有           | `e2e/helpers.ts:22/94/96`                                                                                                                                                                                                                                                                                                                     |
| 狀態掛鉤 `__aidocs`                                                                     | 有           | `apps/docs/src/renderer/App.tsx:5102`                                                                                                                                                                                                                                                                                                         |
| 狀態掛鉤 `__genofficeDebug`「由 helpers 設定 GENOFFICE_DEBUG_HOOKS」                    | **文件不符** | `e2e/helpers.ts` 沒有設定這個變數，是各 Sheets spec 自己設 `process.env.GENOFFICE_DEBUG_HOOKS = '1'`（例如 `sheets-arrow-collapse.spec.ts:10`）。掛鉤本身在 `apps/sheets/src/renderer/App.tsx:3855`                                                                                                                                           |
| Sheets ⌘⇧S／⌘⇧Z 是原生選單快捷鍵                                                        | 有           | `apps/shell/src/main/index.ts:3998` 等；`apps/sheets/src/main/sheets-main.ts:4089`                                                                                                                                                                                                                                                            |
| 引用的 spec 檔                                                                          | 有           | docs-word-typing、docs-paragraph-indent、docs-lists-tables-pictures、docs-section-break-undo、docs-layout-undo、docs-style-bold-toggle、sheets-excel-shortcuts、sheets-insert-cells-dialog、sheets-csv-_、pdf-edit、pdf-fit-zoom、pdf-outline-layout、pdf-shift-arrow-selection、markdown-tab、html-tab、html-insert-drag、slides-_，全部存在 |
| 測試 PDF `1_Stock_Funds_Rose_______in_____.pdf`                                         | 不在 repo    | 文件本來就說明不提交，符合                                                                                                                                                                                                                                                                                                                    |

---

## 結論

- 絕大多數條目都能在原始碼找到對應功能。
- 需要注意的有：
  1. **P-02 投影片多選：未找到。** 縮圖面板和投影片瀏覽模式點擊只選一張。清單標注「where supported」，所以算未支援，不算錯誤。
  2. **W-06 Title／Subtitle：部分。** 內建樣式庫與空白範本都沒有這兩個樣式。
  3. **W-08 替代文字：部分。** 只有 AI 工具能設定，沒有使用者介面。
  4. **Slides 選取框顏色：部分。** 色值寫在函式內，不符合 CLAUDE.md 規則 5 要求的畫布色表。
  5. **README 文件不符：** `GENOFFICE_DEBUG_HOOKS` 不是由 `e2e/helpers.ts` 設定。
- 限制：本稽核**只檢查程式碼**。行為、復原、存檔重開和 Office 對照都沒有執行，所以全部維持 **NOT TESTED**。
