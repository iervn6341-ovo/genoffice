# GenOffice Docs 對照 Microsoft 365 Word for Mac：細項稽核

日期：2026-10-12（Asia/Taipei）  
程式碼基準：HEAD `6578a439`（沿用 handoff 的工作樹；本稽核沒有修改產品程式碼）  
原生版本：Microsoft Word 16.113.3，build 16.113.26092714  
範圍：`apps/docs`、`packages/docx-engine`。此份報告是 Docs 檢查員的第一批細項證據，不代表所有 Ribbon 子選單均已完整對齊。

## 證據狀態與執行限制

- **實機 UI**：在 Word 空白 scratch 文件中檢查指定對話框和情境工具列；沒有保存或改動使用者文件。Word 對話框的補充操作摘錄在[監督追加取證](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/word-native-supervisor-observations.md>)；原生 CUA AX 截圖未能匯出為完整逐項檔案，故把實際操作、看到的控制項、是否執行操作逐項記錄在此。
- **官方 Mac 文檔**：只有明確列出 Mac/Microsoft 365 的來源標作 Mac 文檔。一般 Word 支援頁面只佐證其明寫的 Word 行為，不當成 Mac 專屬細節。
- **原碼判讀**：能證明目前 UI 程式碼提供哪些控制項、處理什麼資料；不等於 GenOffice UI 已實際運作。
- **引擎測試**：Vitest 可證明 `.docx` OOXML 產生／解析／往返行為。這些測試沒有把輸出交給 Word 開啟，也不是 GenOffice UI 的儲存重開驗證。
- **GenOffice UI 實測**：專屬 Electron probe [docs-office-mac-details-audit.spec.ts](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/e2e/docs-office-mac-details-audit.spec.ts>) 已從真正 Docs UI 輸入 scratch 文字，操作 Header、Page Number Format、Symbol、Equation、Font 與 Paragraph；儲存後檢查六項 OOXML marker，關閉第一個 Electron，再用 `openFile` 開第二個 Electron，檢查重開內容、MathML 分數結構、頁首文字，以及頁碼格式 dialog 回填值。最終窄測 **1 passed**。首次兩次 sandbox 內 Electron 在 selector 前 SIGABRT；root 從 macOS `.ips` 看到 `_RegisterApplication`／`NSMenuBar` 註冊階段 stack。相同測試在授權的 sandbox 外 launch 成功；此差異是執行環境限制，不能報成 Docs UI 結果。最終執行命令為：

  ```sh
  GENOFFICE_E2E_NO_VIDEO=1 GENOFFICE_DOCS_AUDIT_DIR=/private/tmp/genoffice-docs-audit-root/docs-office-mac-details-evidence ./node_modules/.bin/playwright test --config e2e/playwright.config.ts e2e/docs-office-mac-details-audit.spec.ts --workers=1 --reporter=line
  ```

  成功證據已複製到 repo：[evidence.json](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-success/evidence.json>)、[stage-log.jsonl](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-success/stage-log.jsonl>)、[final-state.png](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-success/final-state.png>)、[docs-audit.docx](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-success/docs-audit.docx>)，以及同目錄內的 [Header UI](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-success/header-dialog.png>)、[頁碼格式](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-success/page-number-format-options.png>)、[Symbol palette](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-success/symbol-palette.png>)、[Equation modal](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-success/equation-modal.png>)、[Font dialog](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-success/font-dialog.png>)、[Paragraph dialog](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-success/paragraph-dialog.png>)、[重開文件](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-success/reopened-document.png>)與[重開後頁碼回填](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-success/reopened-page-number-format.png>)。UI marker 為 seed ✓、頁碼 upperRoman/start=5 ✓、頁首 ✓、≤ 符號 ✓、分數 OMML ✓、1.5 行距 `w:line=360` ✓。重開後 `mfrac` DOM 數量為 1，頁首可見，Page Number Format dialog 回填 `I, II, III, …` 與 `5`。這只驗證此窄案例；其他功能仍依 coverage matrix 標為未驗。

- **折疊 Symbol 的窄窗缺陷**：在 1280 寬視窗的失敗截圖和 stage JSON 中，`Symbols` 組折疊時，點內層 Symbol 按鈕會被共享 `RibbonFoldGroup` 的 click handler 關閉父層，造成 `.symbol-palette` 雖在 DOM 但不可見、無法選字；放寬 scratch 視窗後才完成符號實測。這是可重現的 GenOffice UI 缺陷，不是 locator 問題：[失敗截圖](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-folded-symbol-failure/final-state.png>)、[失敗 evidence.json](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-folded-symbol-failure/evidence.json>)；具體 handler 是[共享 RibbonFoldGroup](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/packages/ui/src/RibbonFoldGroup.tsx:91>)，巢狀 palette 在[Insert Symbol UI](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/docs/src/renderer/components/ribbon-insert-tab.tsx:1338>)。
- probe 程式碼 lint 通過。probe 在此輪成功執行後才加上 `test.skip(!process.env.GENOFFICE_OFFICE_AUDIT, 'Opt-in Office detail audit')`；一般 E2E 預設略過，刻意執行需設定 `GENOFFICE_OFFICE_AUDIT=1`。本輪成功命令發生在加 gate 前。

## Coverage matrix

`✓` 有指定範圍證據；`△` 僅部分控制／情境實測；`—` 未驗。原生欄是本輪實際操作到的項目，不能推成該分頁所有子功能均完整。UI 存檔重開欄只標窄測實際檢查的功能；引擎測試的 `saveDocx → parseDocx` 不等於 Word 互通。

| 範圍                 | Word Mac 實機 UI                                                     | 官方 Mac 文檔                            | GenOffice 原碼                                                                                            | 引擎測試／語意                                                 | GenOffice UI 操作                                             | 存檔／重開                                                            |
| -------------------- | -------------------------------------------------------------------- | ---------------------------------------- | --------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------- | ------------------------------------------------------------- | --------------------------------------------------------------------- |
| Home                 | ✓ 主分頁；原生 Font/Paragraph dialogs 另列                           | △ 行距等一般說明                         | ✓ `Ribbon.tsx` Home controls                                                                              | △ 字型／段落模型與原始 OOXML                                   | △ 本輪只從右鍵開 Font/Paragraph                               | △ Paragraph 1.5 的 OOXML marker 通過；未重開 dialog 核對              |
| Insert               | ✓ 主分頁；Header、Page Number、Symbol、Equation、Table 深查          | ✓ 頁碼、符號、表格、註腳                 | ✓ `ribbon-insert-tab.tsx`                                                                                 | △ 頁碼、方程式、表格、頁首頁尾有引擎覆蓋                       | ✓ Header、頁碼格式、≤、`\frac{x}{y}`                          | ✓ DOCX save→close→openFile 重開，內容與頁碼 UI 回填已驗               |
| Draw                 | ✓ 主分頁曾打開                                                       | —                                        | ✓ `DrawTab`                                                                                               | — 未跑專項                                                     | — 只查源碼                                                    | —                                                                     |
| Design               | ✓ 主分頁曾打開                                                       | —                                        | ✓ `DesignTab`                                                                                             | △ section/settings 路徑存在                                    | — 只查源碼                                                    | —                                                                     |
| Layout               | ✓ 主分頁及 table Layout、Formula、Sort dialogs                       | △ AutoFit 文件可參照；一般佈局未逐項對照 | ✓ `LayoutTab`/Table Layout controls                                                                       | △ section、頁碼、AutoFit 等有用例；Formula/Sort 專項未找到     | △ 頁碼格式由 Insert 入口實測；其他 Layout 未操作              | △ 本輪已重開驗頁碼格式；其餘未驗                                      |
| References           | ✓ 主分頁；TOC/註腳有實機檢查                                         | ✓ TOC、註腳／章末註                      | ✓ `ribbon-references-tab.tsx`                                                                             | ✓ TOC、notes fields 有部分往返                                 | — 本輪無 GenOffice References UI 操作                         | △ 引擎有；GenOffice UI 未往返                                         |
| Mailings             | ✓ Word 主分頁曾打開                                                  | —                                        | Mac TABS 不含 `mailings`；未見 GenOffice mail merge UI                                                    | — 未跑                                                         | —                                                             | —                                                                     |
| Review               | ✓ Word 主分頁曾打開                                                  | —                                        | ✓ `Ribbon.tsx` review 分支                                                                                | — 本輪未跑 review 專項                                         | — 只查源碼                                                    | —                                                                     |
| View                 | ✓ Word 主分頁曾打開                                                  | —                                        | ✓ `ViewTab`                                                                                               | — 本輪未跑 view 專項                                           | — 只查源碼                                                    | —                                                                     |
| Mac File／Backstage  | Word 無 Windows 式 File Ribbon；原生選單未逐項深查                   | —                                        | △ Backstage 有新建、開啟、最近檔、資訊、匯出                                                              | — 未跑                                                         | — 未操作 GenOffice File UI                                    | —                                                                     |
| Print                | △ 檢查 Mac 列印窗格及 Word 專屬選項，未送出工作                      | — 本輪未找到專屬 Print 文檔引用          | ✓ `PrintDialog` 有 preview、printer、range、odd/even、page setup、per sheet、scaling、reverse、background | △ 來源讀取；未測 printer/output                                | — 未操作 GenOffice Print                                      | — 未印、無輸出/重開                                                   |
| Table Design／Layout | ✓ Insert Table、Properties、Formula、Sort 對話框已查看；後兩者未執行 | ✓ Mac 插表與 AutoFit                     | ✓ 有樣式、選項、框線、尺寸、AutoFit、repeat header、properties；未找到 Formula/Sort 對應入口              | △ autoFit/cell margin roundtrip；table formula/sort 專項未找到 | — 本輪未操作 GenOffice 表格 UI                                | △ docx-engine table cases；Formula/Sort 未驗                          |
| Picture Format       | — 未在 Word 選取圖片開啟 context                                     | —                                        | ✓ 源碼有 crop、remove bg、replace、rotate、flip、wrap/align、size/reset                                   | — 本輪無圖片專項往返                                           | — 未操作 GenOffice Picture UI                                 | —                                                                     |
| Chart contextual     | — 未在 Word 選取圖表開啟 context                                     | —                                        | △ Insert 有 bar/line/pie + data grid；`Ribbon.tsx` 未見 chart context tab                                 | — 未驗 chart package                                           | — 未操作 GenOffice Chart UI                                   | —                                                                     |
| Equation contextual  | ✓ Word 插入後 11 類結構選單逐一展開                                  | △ LaTeX/UnicodeMath 說明為一般 Word 文檔 | △ 8 個預設、LaTeX dialog/MathML preview；未見結構 gallery context                                         | ✓ OMML 與 converter probes                                     | ✓ Equation UI 插 `\frac{x}{y}`；重開 live DOM 有 1 個 `mfrac` | ✓ OOXML OMML marker 與 GenOffice 重開 DOM `mfrac` 已驗；Word 開啟未驗 |
| Font dialog          | ✓ Font/Advanced 兩分頁                                               | △ 通用 Word 字型說明                     | ✓ 三下拉、四 checkbox 等有限 controls                                                                     | △ `rawRPr` 保存未建模屬性                                      | ✓ dialog opened、3 dropdown/4 checkbox；未改值                | — dialog 重開回填未驗                                                 |
| Paragraph dialog     | ✓ Indents & Spacing、Line and Page Breaks、Asian Typography 三分頁   | △ Word 支援頁為一般功能說明              | ✓ 對齊、行距、左右/特殊縮排、段前後距                                                                     | △ keepNext/keepLines/widowControl/tabStops 有模型/解析         | ✓ dialog opened、3 dropdown；UI 設為 1.5 lines                | △ `w:line=360` marker 已保存；重開後未重新打開 dialog 回填            |

**本輪可重現的執行結果**：`packages/docx-engine` 直接執行 Vitest，第一組 7 個檔案、164 tests passed：`math.test.ts`、`header-footer-rich.test.ts`、`hf-variants.test.ts`、`note-props.test.ts`、`notes.test.ts`、`sections.test.ts`、`table-edit.test.ts`。第二組 2 個檔案、64 tests passed：`field-display.test.ts`、`protected-text-edit.test.ts`。合計本檢查員跑過 9 個 suite files、228 tests；不合併計算 root 另跑的 LaTeX probe。測試命令：

```sh
cd packages/docx-engine
../../node_modules/.bin/vitest run tests/math.test.ts tests/header-footer-rich.test.ts tests/hf-variants.test.ts tests/note-props.test.ts tests/notes.test.ts tests/sections.test.ts tests/table-edit.test.ts
../../node_modules/.bin/vitest run tests/field-display.test.ts tests/protected-text-edit.test.ts
```

## 優先差異與可交實作者的驗收條件

### P1 — 頁首頁尾目前是簡化文字編輯流程，尚未呈現原生欄位／樣式工作流

**證據狀態：實機 UI ✓；官方 Mac 文檔 △（頁碼操作）；原碼判讀 ✓；引擎測試 ✓；GenOffice Header/頁碼 UI ✓ 並實際儲存、關閉、重開；Word 開啟 GenOffice 輸出未驗。**

**原生操作及實際所見：**在 Word scratch 文件走 `Insert > Header`，頁首下拉有 21 個內建樣式；進入頁首後出現 Header & Footer context controls，可見日期、Field、圖片、切換頁首/頁尾、前/後節、Link to Previous、Show Document Text、距頁頂/底與 Close。頁碼路徑 `Insert > Page Number > Format Page Numbers` 可設定編號格式、續前節／指定起始值；另有 Include chapter number 與章節樣式/分隔符控制（空白文件中章節欄位 disabled）。原生編輯區能保留混合內容與欄位。Microsoft Mac 頁碼說明也列出 Current Position、Format Page Numbers、對齊與 PAGE field 插入。[官方頁碼說明](https://support.microsoft.com/en-us/word/add-page-numbers-to-a-header-or-footer-in-word)

**GenOffice 目前實際能從原碼確認的行為：**Insert 頁首頁尾下拉是單一 text input + OK/Cancel（`ribbon-insert-tab.tsx:110-153,1184-1233`），頁碼下拉含上/下各左中右六位置、格式與移除（同檔 `:1235-1294`）。格式 modal 有 7 個選項：decimal、dash、lower/upper letter、lower/upper Roman、chineseCounting，另有 start 欄位，且將值套用到游標所在 section（`App.tsx:2517-2553,6701-6749`）。第一頁不同、奇偶頁不同**已有實作**（`ribbon-insert-tab.tsx:1296-1312`），不列缺項。頁面上的 HeaderFooterArea 可雙擊編輯，但為每段純文字表面；圖片為 `contentEditable=false`、讀取顯示（`HeaderFooterArea.tsx:110-115,218-263`）。

**資料語意與往返：**引擎資料模型有多段 rich paragraphs、PAGE/NUMPAGES 標記（`types.ts:834-851`），儲存計畫合併既有 header part（`patch.ts:723-792`）。測試證明多段格式化頁首、真正 PAGE field、表格與圖片 bytes 可保留（`header-footer-rich.test.ts:32-71,447-475,510-540`）。頁碼 `w:pgNumType` 可解析及寫回 section；測試有 upperRoman/start=5 的保存重解析（`section.ts:288-318`、`sections.test.ts:579-609`）。所以主要差異是控制流程/編輯能力，不是「引擎完全不支援頁首」或「頁碼不儲存」。

**GenOffice UI 實測：**Header menu 用單一 input 輸入 `Audit header` 並按 OK，重開 DOCX 後頁首文字仍可見。Page Number Format UI 列出 decimal、numberInDash、lower/upper letter、lower/upper Roman、chineseCounting 共 7 種；實際設 `upperRoman`、起始值 `5`，存檔 `sectPr/w:pgNumType` marker 正確，關閉後重開再開 Format dialog，格式及起始值都回填正確。細節與 artifacts 在上述成功 evidence。

**差異、影響：**目前使用者不能從 UI 選 21 種頁首版型、插入日期/欄位/圖片、按節瀏覽或切換 Link to Previous；多段格式和圖片雖可保留，但不代表 UI 能逐一編輯。頁碼基本格式及 section start 已由 UI 驗證，但格式清單仍比 Word 短，且沒有章節編號/樣式/分隔符設定。對品牌頁首頁尾、分節頁碼及混合內容文件屬高影響。

**驗收案例：**

1. 新建三節文件，在第 2 節取消 Link to Previous 並套一個頁首頁尾樣式；只改第 2 節，其他節及 first/odd/even variants 不變。儲存、關閉、重新開啟後各節連結旗標與文字仍正確。
2. 插入日期欄位、PAGE、NUMPAGES 與一張標誌圖片，設定多段/左右對齊；保存 `.docx`，檢查 `w:fldSimple` 或 begin/instrText field、圖片 relationship、section references，重開確認可編輯且不丟圖、不把 field 展平成字串。
3. 頁碼格式 UI 含原生格式與 chapter number 控制；續前節應在目標節移除 start override，指定起始值應在對應 `sectPr/w:pgNumType` 寫正確 `w:fmt`/`w:start`；關閉重開後 UI 回填完全相同。章節選項在無 heading style 時清楚 disabled。

### P1 — 符號是 50 字元快捷板，方程式沒有原生結構編輯工具列

**證據狀態：實機 UI ✓；官方一般 Word 文檔 ✓、方程式頁非 Mac 專屬；原碼判讀 ✓；引擎測試 ✓；GenOffice Symbol/Equation UI 操作 ✓；輸出 DOCX close/reopen 後 MathML `mfrac` DOM ✓；Word 開啟此輸出未驗。**

**原生操作及實際所見：**Word `Insert > Symbol > Advanced Symbol` 有 Symbols 與 Special Characters 兩分頁、字型／subset、字元碼與相關快捷資訊。Word `Insert > Equation` 插入後，方程式 context ribbon 有 11 類結構 gallery；scratch 文件逐一展開，總計 230 個 gallery 項目：Fractions 9、Scripts 8、Radicals 6、Integrals 21、Large Operators 40、Brackets 41、Functions 27、Accents 27、Limits and Logs 8、Operators 21、Matrices 22。此數是該版本的範本項目數，不是唯一數學字元數。Microsoft Symbol 說明指出字型會決定可用符號，且 More Symbols 有字型、Special Characters、subset 與 character code；該頁適用 Word 通用版，原生 Mac 分頁則另由實機確認。[Microsoft 符號說明](https://support.microsoft.com/en-us/word/insert-a-symbol-in-word)

**GenOffice 目前能確認的資料路徑：**符號板固定 50 個字符，選取時 `insertContent(s)` 插入一般文字（`ribbon-insert-tab.tsx:57-108,1318-1353`）；沒有字型、類別、subset/字元碼或特字符分頁。此處不應說 Docs 完全沒有 Symbol 按鈕，實際上有獨立 Symbol panel。Equation gallery 只有 8 個預設；自訂 modal 輸入 LaTeX、顯示 MathML preview 並可切 inline（`EquationModal.tsx:8-45,117-164,166-205`），存檔資料是 OMML（`editor/equation.ts:19-50`）。

**測試／實際結果：**本輪實際從 Symbol palette 插入 `≤`，從 Equation UI 輸入 `\frac{x}{y}`；存檔 `word/document.xml` 有 `<m:f>` 與 `x` OMML marker。關閉第一個 Electron、使用 `openFile` 重開後，live DOM 的 `<mfrac>` 數量為 1，截圖顯示分數結構仍是上下分子／分母。這證明此線性分數 UI/save/reopen 窄例，不代表所有 LaTeX 都可用。`math.test.ts:254-267` 涵蓋 OMML 保存與解析；root 另跑 converter probes 並存於 [office-mac-math-probe-2026-10-12.json](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-math-probe-2026-10-12.json>)：`\frac`、`\dfrac`、`\tfrac` 產出相同 OMML；`\sqrt[3]{x}`、`bmatrix` 可轉；`\boxed`、`\overset`、`\underset` 回報 unsupported command。那些是 converter probes，不是 UI / Word 往返，且不把 root 的 probe 數重複加進本報告測試總數。Microsoft LaTeX/UnicodeMath 支援頁說明 Word 線性輸入與 Professional conversion；頁面是一般 Word 說明，沒有足夠證據把具體 UI 控制斷言為 Mac 版本。[Microsoft 線性方程式說明](https://support.microsoft.com/en-gb/word/linear-format-equations-using-unicodemath-and-latex-in-word)

**影響與驗收：**複雜公式難以靠滑鼠逐結構建立，線性/專業格式互轉、UnicodeMath 與符號搜尋不足。驗收需：(a) 在 equation selection/context toolbar 逐項插入分數、上下標、根號、積分、矩陣等，不靠測試 helper 設定 editor model；(b) 符號視窗按字型/subset 查找並插入一個不在現有 50 字集合的字符；(c) 使用 LaTeX 與 UnicodeMath 兩種輸入，錯誤要在 preview 給出定位提示；(d) 保存 DOCX 並重新開啟，確認 OMML 能重新編輯、矩陣/分數結構未扁平化；(e) 對 `\boxed`、`\overset`、`\underset` 明確提供支援或可理解錯誤，而不是靜默丟失。

### P1 — Font / Paragraph dialog 缺少原生進階與分頁控制項

**證據狀態：實機 UI ✓；官方通用 Word 文檔 △；原碼判讀 ✓；資料模型 △；GenOffice Font dialog 開啟/控件實測 ✓（未改值）；Paragraph dialog UI 設 1.5 並保存 OOXML marker ✓；重開文件內容已驗，Font/Paragraph dialog 回填未驗。**

**原生實際操作：**選取文字並開 `Font…`，Font/Advanced 兩頁。Advanced 有字元比例、間距、位置、kerning 等；Font 頁含底線樣式/顏色、雙刪除線、小型大寫/全大寫、隱藏等控制。以段落右鍵 `Paragraph…`，實機看到 Indents & Spacing、Line and Page Breaks、Asian Typography 三頁，含定位點/分頁旗標和中文排版選項。Microsoft 的段落說明列出 Single/1.5/Double/At least/Exactly/Multiple、左右縮排、首行/懸掛、段前後距；另一支援頁列 Keep with next、Keep lines together、Page break before、Widow/Orphan control。後兩頁不是本次特定 Mac build 的細項證據，Mac 實際內容以 native observation 為準。[段落間距說明](https://support.microsoft.com/en-us/word/adjust-indents-and-spacing)、[分行與分頁說明](https://support.microsoft.com/en-us/word/line-and-page-breaks)

**目前 UI 源碼：**Font dialog 有 font/style/size 三下拉、顏色及 underline/strike/superscript/subscript 四個 checkbox（`ContextMenu.tsx:545-747`）；沒有第二個 Advanced 頁面、雙底線/底線顏色、字型效果或 kerning/character spacing/position 控制。Paragraph dialog 有 alignment、line spacing multiples/atLeast/exact、left/right、none/first/hanging、段前後距（`ContextMenu.tsx:792-885,918-1010`）；沒有 tab stop 編輯器、keepNext、keepLines、widow/orphan、page-break-before、Asian Typography 專屬旗標。

**資料語意：**引擎保留未建模字元格式於 `rawRPr`，包含多種 Word 字型屬性；段落模型有 `keepNext`、`keepLines`、`widowControl`、`tabStops`（`types.ts:41-108,456-497`；解析 `parse.ts:3387-3403,3428`）。所以要把此問題標作「UI 不能檢視/編輯這些設定」，不能籠統判定 docx-engine 不支持。另一方面，本輪沒有測試使用者在 dialog 改另一項格式後，這些保留值是否都能保持，因此 acceptance 必須包含。

**驗收案例：**(1) 從 Word DOCX 匯入帶 kerning/char spacing、不同 East Asian/Latin font、tab stop、keepNext/keepLines/widowControl 的段落；只修改字號/行距，保存、關閉、重開後未碰屬性仍逐項等價。 (2) 新增/編輯 Font Advanced 的 scale/spacing/position/kerning 與底線樣式/顏色，在 XML 對應 OOXML 有值且重新開啟回填。 (3) Paragraph dialog 的分頁旗標及 tabs 用實際 UI 勾選/編輯；驗證段落跨頁結果與 `w:keepNext`, `w:keepLines`, `w:widowControl`, `w:pageBreakBefore`, `w:tabs`。 (4) Asian Typography 對話框至少能保留未改值；支援的功能須有可見 UI，不接受只因 raw XML 存在就算 parity。

### P2 — 頁碼清單可用但格式範圍較窄，章節編號設定缺少

**證據狀態：原生 Format Page Numbers dialog 實機 ✓；官方 Mac 頁碼文檔 ✓；原碼 ✓；`pgNumType` 引擎往返 ✓；GenOffice UI 設 upperRoman/start=5、存檔及重開 dialog 回填均 ✓。**

原生格式下拉除 decimal、dash、字母、羅馬數字外，還看見 full-width、circled、數種中文/日文計數格式等；對話框有 Include chapter number，並可選 chapter begins with style 和 separator。GenOffice 原碼只提供前述 7 格式和 section start。這會讓格式特殊的報告/章節文件無法從 UI 建立或改設定，即使現有其他值可能仍隨 OOXML 保留。驗收：讀入每個 native `w:fmt` 目前值須被回填；UI 能設定所承諾的格式；chapter number 能與 heading numbering 綁定且分隔符生效；`start` 空白代表續前節、非空值寫至正確節；每節重開時回填一致。官方 Mac 說明主要證明插入、位置、欄位和格式入口，不當成所有格式清單的權威枚舉。[Microsoft 頁碼說明](https://support.microsoft.com/en-us/word/add-page-numbers-to-a-header-or-footer-in-word)

### P2 — TOC 是真 field，但缺少 Custom Table of Contents 視窗及更新模式

**證據狀態：Word 原生 gallery/custom dialog 實機 ✓；官方 Mac/M365 文檔 ✓；原碼 ✓；engine field roundtrip ✓；UI 未驗。**

**原生步驟：**`References > Table of Contents` 看到五個 Manual 與五個 Automatic gallery 樣式。開 `Custom Table of Contents` 有格式、層級、顯示頁碼、靠右對齊、leader、hyperlink、Modify、Options；Options 可將段落樣式映射到層級，Modify 可編輯 TOC 1–9。Microsoft 說明確認可以顯示/隱藏/對齊頁碼、設定 leader/format/levels、修改層級樣式；選項名稱以該官方內容和實機為準。[自訂目錄](https://support.microsoft.com/en-au/word/format-or-customize-a-table-of-contents-in-word)

**GenOffice：**References UI 只有 Insert TOC 與 Update TOC 兩個按鈕（`ribbon-references-tab.tsx:562-590`）；插入從 heading 建立內容，更新會替換整個 TOC field（同檔 `:79-145,397-408`）。引擎產生真正 dirty TOC field，指令固定 `TOC \o "1-N" \h \z \u`，固定 dot leader tab stop `9350`（`generate.ts:2212-2247`）；測試確認 field structure 和 parser roundtrip 為 tocLine blocks（`field-display.test.ts:494-522`）。這比靜態假文字強，但沒有可讓使用者配置的 Custom TOC options；也沒有分開「只更新頁碼」及「更新整個目錄」的操作。Microsoft 有更新頁碼/整個目錄的官方指示；本輪沒有在 Word 內實際更新 TOC，因此不宣稱操作結果已驗。[Microsoft 更新目錄](https://support.microsoft.com/en-us/word/update-a-table-of-contents)

**驗收案例：**在同一份文件中設 4 層 heading 並把自訂段落樣式映射到層級 2；插入 TOC 時切換頁碼顯示、右對齊、不同 leader、層數、template/format 和 hyperlink；只更新頁碼不得改 heading/樣式，完整更新應重建 entries。保存重開後保留 field、目錄設定和手動修改的 TOC styles；改 heading 後更新需正確移除/加入行並更新頁碼。

### P2 — 註腳可插入/編輯正文，但 footnote/endnote 設定 UI 未提供

**證據狀態：原生實際插入過註腳；設定 dialog 本輪未實機打開；官方 Word Mac/M365 設定文檔 ✓；原碼 ✓；engine 解析/編號測試 ✓；GenOffice UI 往返未驗。**

Word Mac 實機在空白 scratch 用 `References > Insert Footnote` 成功插入；沒有手動開啟 Footnote and Endnote dialog。官方 Mac 說明確認可改 position、number format、整份文件或目前 section scope，並可 restart each section。[官方 Mac 註腳/章末註說明](https://support.microsoft.com/en-US/Word/add-footnotes-and-endnotes-in-word-for-mac)

GenOffice References 區只有 Insert Footnote / Insert Endnote（`ribbon-references-tab.tsx:595-624`）；App 用 multiline PromptModal 編輯 note body（`App.tsx:2688-2696,6517-6544`）。引擎 NoteProps 有 position、number format、start、continuous/eachSect/eachPage（`types.ts:555-565`），能讀取、套 settings 與 section override 並按 body 引用順序計號（`section.ts:204-225`、`parse.ts:5681-5759`）。但本輪檢視的 UI 沒有 note settings dialog/action；`eachPage` 需要 pagination，解析器目前在尚未分頁時按 continuous 處理。測試涵蓋設定讀取、settings/section precedence 和 eachSect numbering（`note-props.test.ts:36-97`）。

**驗收案例：** UI 對 footnote 與 endnote 分別提供位置、數字格式、開始號碼、連續/每節重編與全文件/目前節 scope；切換 scope 後只寫入正確的 `settings.xml` 或目標 `sectPr`。多節 fixture 預期頁註在頁底/文字下、章末註在文件末/節末；重編號跨節正確；save/reopen 回填一致。對 `eachPage` 若 pagination 尚不支援，應限制選項並解釋，而非 UI 看似已設定但結果連續編號。

### P2 — 表格有完整部分 contextual 功能，但插入對話框與原生細項不同

**證據狀態：Word Mac Insert Table dialog/表格 context 實機 ✓；官方 Mac 插入表格文檔 ✓；原碼 ✓；engine AutoFit roundtrip ✓；GenOffice UI 未驗。**

**原生操作：**`Insert > Table > Insert Table` 對話框本機預設 5 columns × 2 rows；AutoFit 選項是 Initial column width (`Auto`/specific)、AutoFit to contents、AutoFit to window，另有 Set as default for new tables。原生 Table Design 見 Header Row、Total Row、Banded Rows、First/Last Column、Banded Columns、style/shading/borders。Table Layout 見選取/格線、Properties、Draw Table、Eraser、Delete/Insert row/column、Merge/Split Cells、Split Table、AutoFit、row/column dimensions、分佈、cell alignment、text direction、cell margins、Sort、Repeat Header Rows、Convert to Text、Formula。Properties 有 Table/Row/Column/Cell/Alt Text；Row 包含 exact height、allow row to split、repeat header。Microsoft Mac 插入表格說明確認 10×8 以外自訂尺寸、AutoFit 三模式、設為預設、Draw Table 與情境 tabs。[Word for Mac 插入表格](https://support.microsoft.com/en-us/word/insert-a-table-in-word-for-mac)

**GenOffice 原碼：**Insert Table Modal 只接受 rows/columns（預設同為 5×2），直接呼叫 `insertTableAt`，沒有 AutoFit Behavior 或 set-as-default（`ribbon-insert-tab.tsx:495-545`）。Table Design 有預設樣式、表格樣式選項、shading、邊框與色/粗細（`Ribbon.tsx:2398-2603`）；Table Layout 有刪除/插入、merge/split cells、尺寸、三種 AutoFit、repeat header、properties（`Ribbon.tsx:2605-2859`）。Table Properties 有 autofit/wrap 等（`Ribbon.tsx:4232-4325`）。本輪對照過程中，未在這些 app contextual controls 找到 Draw/Eraser、Split Table、Sort、Convert to Text、Formula、row split 或各種高度/列分散選項；有待實作者以全程式搜尋和新 UI 操作確認後決定缺項，避免把入口差異當作不支援。

**原生 Formula/Sort 追加觀察：**Word `Table Layout > Formula` 開啟公式文字欄（初始 `=`）、數字格式下拉、Paste Function 與 Paste Bookmark。Paste Function 清單共 18 項：ABS、AND、AVERAGE、COUNT、DEFINED、FALSE、IF、INT、MAX、MIN、MOD、NOT、OR、PRODUCT、ROUND、SIGN、SUM、TRUE；數字格式有 7 項：`#,##0`、`#,##0.00`、`$#,##0.00;($#,##0.00)`、`0`、`0.00`、`0.00%`、`0%`。Sort dialog 有第一至第三階排序、遞增/遞減、標題列選擇、Options；Options 可選 Tab/Comma/Other 分隔、大小寫視為相異及排序語言。當時第三階與標題列選擇 disabled，dialogs 均取消，沒有計算或排序結果。細項記錄在[監督追加取證](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/word-native-supervisor-observations.md>)。

**Formula/Sort 的程式碼證據界線：**root 對 Docs renderer 與 `docx-engine` 搜尋 `tableFormula`、`tableSort`、`FormulaModal`、`FormulaDialog`、`SUM(ABOVE)`、`ribbonTableFormula`、`ribbonFormula` 沒找到對應入口。這是 UI gap 的候選證據，不足以推論引擎對通用 field 或未知 OOXML 的保存／解析能力；不得只以 raw XML 保留推稱 formula 可用。

**資料語意/測試：**engine 有 `autoFit`、cell margins 等模型；產生器映射 fixed/contents/window 為 OOXML `tblLayout` / `tblW`（`generate.ts:2129-2144`），`table-edit.test.ts:201-225` 覆蓋 AutoFit/cell margin/float/style options roundtrip。這不能證明本輪插入 dialog 或 contextual app UI 已實際操作。

**驗收案例：**新插表時支援三種 AutoFit/Initial width/default 行為，另驗 12×9 自訂尺寸；在同一表格實際透過 Table Design 設 Header/Total/Banded/First/Last，透過 Layout 畫線、eraser、split table、sort、formula 或明確界定產品不支援項。Formula 以有數字的兩列 fixture 驗 `=SUM(ABOVE)`、更新欄位值及數字格式；Sort 以 header row + 三鍵 fixture 驗升/降序、分隔符與語言；修改 row split/repeat header/cell margins 後保存重開，逐項比對 `tblPr`, `tblGrid`, `trPr`, `tcMar` 等 OOXML 與頁面輸出。

### P2 — 折疊 Ribbon 下 Symbol palette 會被父組收起

**證據狀態：1280 寬 Electron 實際操作失敗 ✓；共享元件原碼判讀 ✓；放寬至 2200 後 Symbol 操作成功 ✓；存檔/重開窄窗流程未驗。**窄窗時 `Symbols` group 折疊，點子按鈕會觸發[共享 `RibbonFoldGroup`](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/packages/ui/src/RibbonFoldGroup.tsx:91>)內部 close；[Symbol palette](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/docs/src/renderer/components/ribbon-insert-tab.tsx:1338>)是子節點而非 portal，因此父 fold 關閉時 palette 雖存在 DOM 卻不可見，無法選取。已保存的窄窗失敗 evidence 在 [PNG](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-folded-symbol-failure/final-state.png>) 與 [JSON](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-folded-symbol-failure/evidence.json>)。這是 1280 寬折疊狀態下的 UI bug；Equation/Symbol 在寬窗成功不能代表折疊路徑通過。

**實作者驗收：**在 1280 或更窄視窗展開 Insert > Symbols fold group，再點 Symbol；palette 必須保持可見且能插入字符。Equation 子按鈕及 gallery 也要在同一 fold 狀態下能開啟並操作，不因子按鈕 click 收起父層。另測寬視窗控制不回歸。不能用 hidden click、刪 CSS 或直接呼叫 editor API 作為通過條件。

### 後續需先補 native/app 證據的情境項目

- **Picture Format：**GenOffice 原碼可見裁切、去背、替換、90°旋轉、翻轉、wrap/align、寬高與 reset（`Ribbon.tsx:2187-2395`），但本輪未在原生 Word 選取圖片打開 picture context，亦未成功在 GenOffice UI 實測。不能列出可靠的 native 差異或標為 parity pass。
- **Chart：**GenOffice Insert Chart dialog 只有 bar/line/pie、title 與 data grid（`ribbon-insert-tab.tsx:328-489,900-923`）；`Ribbon.tsx` 未見 chartFormat/chartDesign context tab。本輪未在原生 Word 選取 chart，也未跑 chart-specific save/reopen，故「原生 contextual options 與具體缺項」仍待驗。
- **File/Print：**macOS Word 主分頁不包含 Windows File Ribbon；Word 原生選單與完整 File/Word Options 本輪未深查。原生 Print dialog 已只讀查看：份數、自動分頁、全部/目前頁/選取/From–To/自訂頁面範圍、雙面、快速預覽與 PDF menu；Word 專屬區另見 all/odd/even、Print What 六項（Document、Document Properties、Markup List、Styles、AutoText Entries、Key Assignments）、Print form data 與 PostScript text-over-print。沒有改設定或送印。詳細摘錄在[監督追加取證](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/word-native-supervisor-observations.md>)。GenOffice Backstage 接線有 New/Open/Recent/Info 與 PDF/HTML/Images export（`App.tsx:6618-6640`）。PrintDialog 原碼有 printer/settings/preview/range/page setup/per-sheet/scaling/options（`PrintDialog.tsx:67-93,299-512`），但本輪沒有操作 GenOffice Print，也未用任何 printer 實際列印，故無可對照輸出或 save/reopen 結果。媒體品質、送紙、driver job security 等系統印表機選項不是此 Office parity 缺口。
- **大 Ribbon 各分頁：**Word Home、Insert、Draw、Design、Layout、References、Mailings、Review、View 主分頁有實機開過；app Mac Ribbon 清單是 Home/Insert/Draw/Design/Layout/References/Review/View，沒有 Mailings（`Ribbon.tsx:360-374`）。除表格、方程式、字型/段落與頁首頁尾/TOC/註腳等上列項，其他分頁的 dropdowns、對話框及保存重開流程沒有逐項覆蓋。不能把之前的 broad ribbon pass 算成細項驗收。

## 首批可交實作者的問題

1. **頁首頁尾（P1）：**先定義 section/variant/link-to-previous 的編輯資料模型，再補 Header & Footer contextual 工作流、field/date/picture 插入與 style gallery；保住現有引擎 rich paragraphs/PAGE/NUMPAGES/非文字內容的 roundtrip。
2. **Equation/Symbol（P1）：**Equation selection/context toolbar 加常用 OMML 結構插入與線性/專業轉換；符號升級為可按字型/類別搜尋的 dialog。擴充 converter 前先把 unsupported command 集合做成明確的 acceptance corpus。
3. **Font/Paragraph（P1）：**加入 Font Advanced 與 Paragraph pagination/tabs/Asian typography 可編輯 controls；優先以 fixture 驗證只改一個欄位時原始 OOXML 其他未改屬性不被扁平化。
4. **TOC / Footnote settings（P2）：**TOC custom options / update modes 和 footnote/endnote section-level settings 對接既有 fields/NoteProps，不需重造 OOXML 基礎。
5. **Table insert/context（P2）：**讓 Insert Table dialog 對齊 AutoFit initial/contents/window/default；對照 Word 實機已看到的 Formula/Sort dialog，先決定支援範圍，再提供可往返 fixture；不把通用 raw XML 留存等同功能實作。
6. **Folded Symbol / Equation 子選單（P2）：**修共享 fold group 的點擊傳播／關閉行為，確保折疊時 Symbol palette 與 Equation gallery 仍可見、可選；以 1280 寬及寬視窗都實際操作驗收。

## 來源與證據檔

- 主要原碼入口：[Insert tab — Header/Page Number/Symbol/Equation/Table](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/docs/src/renderer/components/ribbon-insert-tab.tsx:57>)、[HeaderFooterArea](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/docs/src/renderer/components/HeaderFooterArea.tsx:110>)、[EquationModal](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/docs/src/renderer/components/EquationModal.tsx:8>)、[Font/Paragraph ContextMenu](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/docs/src/renderer/components/ContextMenu.tsx:545>)、[Page Number Format UI](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/docs/src/renderer/App.tsx:6708>)、[References tab](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/docs/src/renderer/components/ribbon-references-tab.tsx:79>)、[Table/Picture/Context ribbon](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/docs/src/renderer/components/Ribbon.tsx:2187>)、[PrintDialog](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/docs/src/renderer/components/PrintDialog.tsx:67>)、[docx-engine MathML/OMML parsing](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/packages/docx-engine/src/math.ts:768>)、[Table engine tests](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/packages/docx-engine/tests/table-edit.test.ts:201>)。
- Word native supplemental evidence: [word-native-supervisor-observations.md](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/word-native-supervisor-observations.md>)。Docs Electron 成功與折疊失敗的 JSON、PNG、stage log、DOCX 保存在 [docs-success](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-success/evidence.json>) 與 [docs-folded-symbol-failure](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-evidence-2026-10-12/docs-folded-symbol-failure/evidence.json>)。
- 本輪沒有新增或修改產品源碼；自有 UI probe 是 [docs-office-mac-details-audit.spec.ts](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/e2e/docs-office-mac-details-audit.spec.ts>)，預設以 `GENOFFICE_OFFICE_AUDIT` opt-in 執行。
- root 的 converter 輸出：[office-mac-math-probe-2026-10-12.json](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/office-mac-math-probe-2026-10-12.json>)。原始 handoff：[handoff-office-mac-tool-details-audit-2026-10-12.md](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/handoff-office-mac-tool-details-audit-2026-10-12.md>)。
