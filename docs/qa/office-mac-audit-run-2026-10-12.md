# Microsoft 365 Mac 細項實測：執行紀錄

日期：2026-10-12（Asia/Taipei）。起始 HEAD：`6578a439`。

使用者已恢復實測，依三位檢查員、三位實作者與監督的分階段安排進行。本紀錄承接 `handoff-office-mac-tool-details-audit-2026-10-12.md`；原交接文件保留暫停當時狀態。

## 分工與執行設定

| 角色           | agent                    | 模型／思考                        | 狀態         |
| -------------- | ------------------------ | --------------------------------- | ------------ |
| Docs 檢查員    | `/root/docs_inspector`   | GPT-6 Luna／Extra High（`xhigh`） | 本輪檢查完成 |
| Sheets 檢查員  | `/root/sheets_inspector` | GPT-6 Luna／Extra High（`xhigh`） | 本輪檢查完成 |
| Slides 檢查員  | `/root/slides_inspector` | GPT-6 Luna／Extra High（`xhigh`） | 本輪檢查完成 |
| 三套實作者     | 待檢查結果後分批啟動     | GPT-6 Sol／High（`high`）         | 尚未啟動     |
| 監督與整合驗收 | `/root`                  | 沿用目前設定                      | 本輪完成     |

模型設定遵循對話中直接指定的 High；交接檔後來出現的 Medium 不作為覆蓋該指示的依據。

## 執行規則

- 最大同時四個 agents，包含 root，因此分階段執行。
- 檢查階段不修改產品程式碼；各 agent 只維護自己的報告與必要的獨立測試／探針。
- 原生 Office 共用桌面依 Word → Excel → PowerPoint 排程；僅操作空白測試文件，保留使用者原文件。
- Electron E2E 跨 agent 串行執行。純 Node／單元測試可並行。不得在 E2E 執行期間重建 app。
- 舊 broad ribbon passes 不作為細項一致性證明；區分實機觀察、官方 Mac 文件、來源判讀、實測結果與未驗證項目。
- 每個可交付問題包含操作步驟、預期／實際結果、資料語意與儲存重開、來源位置、影響及驗收條件。
- 共用數學／符號／UI 模組由監督協調單一修改負責人，避免多位實作者同時修改。

## 預定交付

1. `docs/qa/docs-office-mac-details-audit-2026-10-12.md`
2. `docs/qa/sheets-office-mac-details-audit-2026-10-12.md`
3. `docs/qa/slides-office-mac-details-audit-2026-10-12.md`
4. 整合優先順序與複驗紀錄（本文件更新）。

三份應用稽核報告已建立，已取得原生細項、來源盤點、引擎及實際 Electron 操作證據。這是本輪細項稽核結果；沒有任何一套被標為完全一致。每份報告的覆蓋矩陣列出已驗、來源判讀及未驗範圍。

## 監督的共用數學模組覆核

- `packages/docx-engine`：`vitest run tests/math.test.ts`，37 個既有案例通過。
- `packages/pptx-engine`：`vitest run tests/equation-table.test.ts`，2 個既有案例通過，確認匯入的原生方程式 XML 能在周圍文字修改時保留。這與 Slides 的新建方程式入口插入一般文字方塊是不同情況。
- 直接呼叫 `latexToOmml` 的八項結果保存在 `docs/qa/office-mac-math-probe-2026-10-12.json`：三種 fraction 指令產物相同；指定次方根與矩陣可轉換；boxed／overset／underset 不支援。
- 上述通過只表示目前已涵蓋的引擎案例通過，不代表完整 Office 結構圖庫、UI、存檔重開或原生方程式編輯均一致。

## 第一批已確認差異與交付順序

以下是已取得證據的實作方向，不是已完成的修改。個別項目以三份報告中的實機／來源／測試標籤為準。

| 優先順序                | Docs                                                                          | Sheets                                                                 | Slides                                                                       | 共同驗收要求                                                         |
| ----------------------- | ----------------------------------------------------------------------------- | ---------------------------------------------------------------------- | ---------------------------------------------------------------------------- | -------------------------------------------------------------------- |
| 1：頁首頁尾／日期／範圍 | 原生樣式圖庫、分節導覽及連結、欄位／圖片編輯流程；保留已存在的 first/odd/even | 圖片／字型／路徑、頁首變體與頁面版面流程；不可把引擎保留等同 UI 可建立 | 固定／自動日期模式回填、目前頁／全部範圍、標題頁排除、備忘稿／講義、格式保留 | UI 設定 → 存檔 → 關閉重開 → 再編輯；檢查欄位、範圍與格式均保留       |
| 2：方程式／符號         | 已有 OMML 插入，缺結構編輯圖庫和完整符號流程；converter 子集不足              | 方程式目前 rasterize 成 PNG；符號附加儲存格文字                        | 新建方程式是普通文字方塊；缺 Symbol；匯入的 OMML 有保留路徑                  | 新建與匯入分開測；文件物件類型、再次編輯和 Undo/Redo，不能只檢查外觀 |
| 3：格式對話框           | Font Advanced、Paragraph 分頁／定位點／中文排版；TOC 自訂與註腳設定           | 對齊、邊框、底線效果、圖樣、Protection 回填、資料剖析精靈              | 母片格式控制、動畫進階、Setup/Record、圖表／圖片細項                         | 修改一項不覆蓋其他未改屬性；同功能區／右鍵／窗格狀態一致             |

自訂放映另有檔案可攜性問題：`App.tsx` 與 `show-actions.ts` 以檔案路徑作 localStorage key；本機同路徑可重讀，不代表已保存至 PPTX `custShowLst`。

實作者分工仍為 GPT-6 Sol／High。共享的數學結構、符號資料和通用 UI 需由一位負責人先定契約，再由各應用接線；每一修正仍需檢查員獨立複驗。

## 環境排查與證據保存

- 原生 Office 在空白測試文件取證；使用者簡報沒有保存修改。Word 已取得頁碼、目錄、插入表格、表格屬性等細項；Excel 補取得水平 8、垂直 5 個對齊項目及資料驗證三個分頁；PowerPoint 補取得方程式 11 類結構入口、符號圖庫與轉換控制。每個未展開的細項仍列待驗。
- Excel 控制連線後續重複逾時，PowerPoint 結構圖庫後回報 screen capture size `(0,0)`；root 獨立重新綁定 PowerPoint 仍同樣失敗。這些是實機工具限制，不當成產品功能失敗。
- 最初 Docs 與 Sheets 的 sandbox 內 Electron launches 約 0.4 秒即 SIGABRT，沒有進入 UI。Sheets 完整 log：`/private/tmp/genoffice-sheets-audit-20261012.log`。macOS crash stack 位於 `_RegisterApplication`／`NSMenuBar`；root 在 sandbox 外執行已成功 launch 並操作 Docs 頁首，確認原先啟動限制可繞由授權的正常桌面測試環境排除。
- Sheets 檢查員曾啟動後中止 `build:all`，造成 ignored out 部分更新；root 待所有 Electron 時段釋放後完成一次完整 `npm run build:all` 恢復一致產物。log：`/private/tmp/genoffice-office-audit-build-20261012.log`。產品源碼沒有改動。
- 原始三張使用者附件已保存在 `docs/qa/office-mac-evidence-2026-10-12/`，README 明確區分使用者參考圖與本輪測試圖。
- 具體 UI 操作及保存結果另於本文件末尾更新；locator 或 strict-mode 失敗先修測試，不直接報產品 bug。

### 新增實測缺陷：折疊群組內的符號選單被關閉

Docs 在 1280px 測試視窗：Insert → Symbols 折疊群組 → Symbol。50 個 symbol buttons 已進 DOM，但 `≤` 因上層群組關閉而不可見，不能點選。這是 UI 行為缺陷，與此前按鈕含圖示造成的 locator 錯誤不同。

來源覆核：`packages/ui/src/RibbonFoldGroup.tsx:91` 的 click handler 遇任何 button 都關閉群組，而 Docs `components/ribbon-insert-tab.tsx:1338` 將 symbol palette 放在該群組裡。其他應用採相同元件的下拉仍需獨立驗證，不能因共享程式碼就斷言全數同樣失敗。

驗收：窄／寬視窗各走符號與方程式等巢狀選單，入口和下一層都可見及操作；點外部、Escape 和完成命令仍正常關閉。共用元件修正需單一負責人。Docs 其餘細項以正常放大測試視窗繼續，不強制點擊隱藏字元，也不刪除 CSS。

## 實際操作、保存與重開結果

### Docs

專屬 E2E 1 case 通過，涵蓋頁首文字、頁碼格式（大寫羅馬數字，從 5 開始）、符號 ≤、LaTeX 分數及段落 1.5 倍行距。保存 DOCX 的六項 marker 全部成立；關閉原程序後重新開啟，確認文字、MathML `mfrac`、頁首及頁碼格式回填。Font dialog 已開啟讀取／取消，沒有修改字型值，因此不宣稱完成 Font 格式往返。

相關 docx-engine 9 files／228 tests 通過，包含 root 另跑的 math 37 項，不重複加總。寬視窗可完成符號插入；1280px 下另取得上述折疊群組失敗，保留為獨立缺陷。證據：[成功 JSON](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-success/evidence.json>)、[DOCX](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-success/docs-audit.docx>)、[窄視窗失敗](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-folded-symbol-failure/evidence.json>)。

### Sheets

原六個 E2E cases 的寬視窗 run 有四項通過操作／保存／重開：Format Cells 的 Locked Off、逗號文字分欄、New Comment 的 Notes 編輯、Print pane 的頁面設定。New Comment 保存的是 legacy comments＋VML，而非 threaded comments；這是已確認的語意差異。

focused run 測三項，未重跑上述四項：符號／方程式、API 種入 DV、真實 UI 建立 DV。符號／方程式最初因錯誤預設圖片 alt 為 `Equation.png` 而未完成；root 查輸出 drawing name 是 `Picture 2`，修正為 scratch 唯一圖片後只補跑此 case，最終 1 passed（7.3s）。Ω 在 Greek 分類插入、保存、關閉重開成功；LaTeX 方程式保存為圖片（drawing/media），重開後唯一 `.xlsx-image` 可見且 `naturalWidth > 0`。這證明目前圖片降級路徑正常，不表示原生結構方程式對等。持久證據：[XLSX](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/sheets-e2e/symbol-equation-final/symbol-equation-audit.xlsx>)、[驗證摘要](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/sheets-e2e/symbol-equation-final/evidence.json>)、[最終 log](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/sheets-e2e/symbol-equation-final/final-e2e.log>)。

DV 已走真實 UI 的 Add Rule → Dropdown → 新增 Open、Closed → Done，再保存、關閉、重開。它與 API builder 產出同樣的 JSON `formula1`；gateway 將 JSON 逐字包成 quoted literal 寫入 XLSX，而非 Excel 清單需要的逗號分隔 literal。兩項 DV 由 soft parity assertion 記錄為失敗；現況保存／重開捕獲成功不等於格式對等。Microsoft 對 List Formula1 的定義是逗號分隔值或工作表參照。[官方定義](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.add)。尚未用原生 Excel 打開此輸出檔，因此不宣稱已實測原生 Excel 的下拉／修復行為。

Sheets 共七個獨立案例，跨最後有效 runs 的結果為五項現況操作／保存／重開通過、兩項 DV parity 失敗；並非一次完整七-case suite 的統計。純 renderer／gateway 11 files／181 tests 通過；沒有覆蓋此次現行 Univer JSON 清單到 Excel literal 的映射。最終持久檔案及各次執行差異見 [Sheets 報告](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/sheets-office-mac-details-audit-2026-10-12.md>)，不把首次 sandbox launch 或測試 selector 錯誤計為產品 bug。

### Slides

修正測試定位及文字分段讀取後，最終 2 cases 均執行到保存 XML 檢查，2 failed；均為具體功能差異的 soft assertions，非 app launch／locator 問題：

- 頁尾固定日期第一次保存：兩頁固定、零頁 `datetime1`。關閉、重開，再開對話框時 Auto Update 被勾選。再次套用全部後：兩頁變為 `datetime1`，固定日期零頁，頁尾仍在兩頁。這證明回填改變日期模式。
- 新方程式 `x² + y² = z²` 保存為 Cambria Math 一般 `p:sp txBox="1"`，沒有 OMML／`a14:m`；Insert 的 Symbol control 為零。此 case 沒有關閉重開，因此不列新方程式 UI 重開通過。匯入 OMML 的既有引擎保留路徑另行測試，不與新建混同。

Slides 相關引擎測試共 43 項通過（HF 2、SmartArt／master 8、animation／transition 31、既有 OMML 2），不等於介面完全一致。證據：[HF JSON](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/slides-e2e/header-footer/header-footer-evidence.json>)、[Equation JSON](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/slides-e2e/equation/equation-evidence.json>)、[最終 E2E log](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/slides-e2e/final-e2e.log>)。

## 後續實作交付與界線

本輪只新增稽核報告、opt-in 探針及測試證據；尚未修改產品源碼。第一批可交付實作者的問題為：共用折疊選單、Slides 日期模式回填與套用範圍、資料驗證清單映射，以及各應用的結構方程式／符號功能。功能名稱、可見控制與保存的物件類型需一併驗收。

原生 Office 版本已記錄，Mac 專屬選項以實機／官方適用版本確認；未取得的 Excel／PowerPoint 深層介面保留待驗。原生 Excel 實際打開導出的資料驗證檔、PowerPoint Symbol 字元插入情境、其餘動畫／放映及備忘稿／講義深層設定尚無完整實機證據。不能由本輪 pass 數推論全部工具細項等同 Office。

## 最後檢查

- 三份稽核 spec 的 ESLint 通過；格式由 Prettier 核對／修正。
- 新增 spec 全部以 `GENOFFICE_OFFICE_AUDIT` opt-in，正常測試預設不執行這些已知 parity 失敗案例。
- 產品 tracked diff 為空。保留原有 untracked 檔案，未提交／推送。
- 測試文檔、原始 XML、截圖與 logs 保存在 `office-mac-evidence-2026-10-12/`；各報告連結指向持久副本，原始 JSON 保留暫存出處。
- 本輪仍有明列未驗範圍，尤其原生 Excel／PowerPoint 控制工具失效後的深層選單；此結果不是完整 Microsoft 365 Mac 相容認證。
