# 其他服務是否有同樣問題：跨服務唯讀稽核（2026-10-21）

- 範圍：Docs / Sheets / Slides / PDF / Markdown / HTML，對使用者回報的 5 個問題逐項檢查。
- 性質：唯讀稽核，未修改 `apps/`、`packages/` 原始碼，未 commit。探針在 `e2e/_probe/xs-*.probe.spec.ts`（未 commit），原始結果在 `/private/tmp/claude-501/xs/*.jsonl`。
- 建置狀態：Docs / PDF / Markdown / HTML 的 `out/` 為 10-21 00:49–00:50（原始碼自 10-16 後無變動）；Sheets `out/` 為 10-21 21:09、Slides 為 10-21 21:37（另兩個子代理正在修，探針量到的是它們的中途建置）。
- 驗證層級標籤：【GUI】Playwright 實測；【Chromium】Docs 以 vite dev server + 無頭 Chromium（假 preload）實測，非 Electron；【合成 DnD】以 DragEvent 序列模擬瀏覽器拖放（見第 4 項限制）；【讀碼】僅讀程式碼推論；【Word 實測】Word for Mac 的 AppleScript `grow font`。

## 總表

| 問題 | Docs | Sheets | Slides | PDF | Markdown | HTML |
|---|---|---|---|---|---|---|
| 1 Undo 後右鍵選單不同 | 無：選單由 `editor.state` 即時算，GUI 插入表格→Undo→右鍵與插入前一致 | （原問題） | 無：GUI 表格／文字方塊插入、插列後 Undo，選單項目皆與前相同 | N/A：只有縮圖選單，固定 6 項（GUI 前後相同） | N/A：原生 Electron 選單，每次依 params 現建 | N/A：同 Markdown |
| 2 右鍵選單長度太短 | 無：`.ctx-menu` 無 max-height；15 項表格選單 455–468px，四種位置都完整顯示、不捲動。小風險：子選單不翻轉，視窗 700px 時超出 | （原問題） | 反向問題：表格儲存格選單 29 項 887px，視窗 760px 時底部 127px 被切且不能捲動 | 輕微有問題：縮圖選單夾位用 190px，實高 198px，貼底時被切 8px | N/A：原生選單由 OS 控制 | N/A：同 Markdown |
| 3 增大／縮小字型級距 | 大致無（與 Word 梯度一致、連點可累加）；缺口：72pt 以上不再放大、混合字級被壓平成同一值 | 有問題：梯度缺 22（20→24）；混合字級壓平成同一值 | （原問題） | N/A：只有數字輸入框 | N/A：無字型大小 UI | 低：線性 ±2px，非 Office 梯度（設計差異，未 GUI） |
| 4 滑鼠選取時文字被移到他欄 | 有問題（條件式）：按在已選取文字上拖到他欄＝文字被搬走；CellSelection 拖放會把表格結構弄壞（合成 DnD） | （原問題 4a） | （原問題 4b） | N/A：無表格選取 | 有問題（條件式）：同 Docs（合成 DnD） | N/A：無表格選取 |
| 5 多物件操作 undo 粒度 | 無：GUI 2x2 儲存格範圍 8 種操作、4 段落 4 種格式、取代全部，皆一次 Ctrl+Z 還原；圖片／形狀無多選（N/A） | 無（undo 粒度）：貼上／填滿／排序／多範圍刪除皆 1 步；但另有「多範圍時功能區指令只套用到最後一個範圍」 | （原問題） | 無（讀碼）：一個動作一批 op 一步，無多選 | 無：GUI 全選粗體／斜體、清單、表格插列／刪欄皆 1 步 | 無（讀碼）：patch 一步，無多選 |

## 逐項細節

### 1. Undo 之後右鍵選單內容與操作前不一樣

**Docs：無。** 選單項目在 render 當下由 `editor.state` 計算，沒有快取：`apps/docs/src/renderer/components/ContextMenu.tsx:94-102`（selection／canEdit／canComment）、`:158-162`（inTable）。
- 【GUI】`docs1.jsonl` Q1.a/e：段落右鍵選單（10 項）→ 以功能區插入 3x3 表格 → Ctrl+Z → 同位置右鍵，項目完全相同。Q1.g：表格儲存格選單（15 項）→ Insert Below → Undo → 右鍵，項目完全相同。
- Q1.f 空段落的「New Comment」呈停用，是因游標處沒有字詞（`canComment`），屬選取相依而非快取殘留。
- 附帶：【GUI】在 2x2 CellSelection 內右鍵，選取維持 4 格（`docs2.jsonl`）。2026-10-20 報告的 D-03（右鍵折疊範圍）在此建置未重現。

**Slides：無。** `App.tsx:3511-3515` 的 `ctxItems` useMemo 依賴 `ctxMenu, slides, selectedIds, slide…`，且選單於 mousedown 即關閉，每次右鍵重新建立。【GUI】`slides.jsonl`：畫布選單（貼上／新增投影片／背景格式）在「插入表格→Undo」「插入文字方塊→Undo」後與基準相同；表格儲存格選單在「插入下方列→Undo」前後相同。

**PDF／Markdown／HTML：N/A。** Markdown、HTML、PDF 內容區右鍵都走 `installContextMenu`（`packages/electron-utils/src/context-menu.ts:269-282`，每次 `context-menu` 事件用 `params` 現建原生選單，無快取）。PDF 另有縮圖選單（`apps/pdf/src/renderer/App.tsx:8563-8620`），項目固定，只有 `pageCount <= 1` 時停用刪除；【GUI】旋轉→Undo 前後項目相同。

### 2. 右鍵選單長度太短

**Docs：無。** `.ctx-menu` 只有 `position: fixed; padding…`，沒有 max-height／overflow（`apps/docs/src/renderer/styles.css:10251-10260`）。位置修正在 `ContextMenu.tsx:117-127`（右緣往左收、下緣往上翻）。【GUI】`docs1.jsonl` Q2：表格內最長 15 項選單高 455px（視窗 760）／468px（視窗 800），`overflowY: visible`、`scrollH == clientH`，儲存格在 y=140/400/610/700 時都完整落在視窗內。
- 小風險：選單沒有 max-height，視窗高度低於約 470px 時會被切且無捲動；有圖片／清單／欄位時再多 2–5 項（估約 600px）。
- 小風險：`.ctx-submenu { left: calc(100% - 4px) }`（`styles.css:10316-10320`）不翻轉。【Chromium】視窗 700px 寬、在表格內右鍵：子選單 right=837px > 視窗 700px（`docs4-web.jsonl`）；1280／900px 寬正常。

**Slides：不是「太短」，但有反向問題（選單過長且不能捲動）。** `.ctx-menu` 無 max-height（`apps/slides/src/renderer/styles.css:5227-5236`），`ContextMenu.tsx:26-34` 只把 y 夾到 `max(0, innerHeight - h - 4)`，選單比視窗高時就從 y=0 往下溢出。【GUI】表格儲存格右鍵 29 項、高 887px，視窗 760px：`rect.y=0, bottom=887`，底部約 127px（Align Middle／Align Bottom／Format Table／Delete 一帶）點不到、也不能捲動（`slides.jsonl`）。文字方塊選單 17 項／514px 正常。
- 影響：視窗高度低於約 890px 的使用者。13 吋筆電預設視窗常中招。

**PDF：輕微有問題。** 縮圖選單位置夾到 `innerHeight - 190`（`App.tsx:6999-7006`），實際高度 198px（`styles.css:558-567` 的 padding 5px + 6 項各約 31px）。【GUI】`pdf2.jsonl`：最後一張縮圖右鍵，`y=570, bottom=768`，視窗 760，底部被切 8px（最後一項「Insert blank page」下緣）。

**Markdown／HTML：N/A**（原生選單，長度由 OS 控制）。

**順帶提醒：** `packages/ui/src/dropdown.css:79` 的 `.gs-dd-pop` 固定 `max-height: 320px`（不論視窗多高），影響各 App 對話框內的下拉選單，不是右鍵選單；未測，僅標示。

### 3. 增大／縮小字型級距

**Office 梯度的查證**
- 【Word 實測】本機 Word for Mac 的 AppleScript `grow font`（連續 28 次，起點 5pt）：5.5, 6.5, 7.5, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72, 80, 90, 100, …, 150。`shrink font` 反向，並在 5 以下繼續 4, 3。混合字級：10pt+20pt 一次 grow 後各自為 10.5 與 22（每段各走自己的下一級），shrink 後回 10 與 20。
- Excel／PowerPoint：Excel 的 AppleScript 沒有 grow font 指令，System Events 送按鍵被系統拒絕（error 1002），所以**無法用 App 實測**。Excel 梯度（8,9,10,11,12,14,16,18,20,22,24,26,28,36,48,72）屬【推論】，PowerPoint 梯度採使用者提供的清單。

**Docs：大致無，有兩個缺口。**
- 級距表 `apps/docs/src/renderer/components/Ribbon.tsx:510-513`、步進 `:1659-1670`，連點合併在 `font-size-step.ts:34-35`（用 `pending.size` 累加）。
- 【GUI】`docs1.jsonl` Q3：預設 11pt 連按 6 次 → 12, 14, 16, 18, 20, 22；縮小 → …12, 11, 10.5, 10；明確 8pt 縮小 → 7.5, 6.5, 5.5；非梯度 11.5 放大 → 12, 14；連點兩下（無等待）11→14。以上與 Word 實測梯度一致，沒有「按兩次才變一級」。
- 缺口 A（有）：72pt 再放大仍是 72（Word 為 80, 90, …）；縮小到 5 以下不再變小（Word 為 4, 3, …）。
- 缺口 B（有）：混合字級被壓平。`font-size-step.ts:34` 只取一個 base（第一段或 `fallbackSize`），再對整個選取 `setMark`。【GUI】`AAAA`=10pt、`BBBB`=20pt 全選放大一次 → 兩段都變 10.5（Word 為 10.5 與 22）。

**Sheets：有問題。** 梯度 `apps/sheets/src/renderer/ExcelShell.tsx:174-175`（`8,9,10,11,12,14,16,18,20,24,26,28,36,48,72`）、步進 `:192-199`、起點 `echoSize = selectionFormat?.fontSize ?? 11`（`:3037`）、按鈕 `:3133/3140`。
- 【GUI】`sheets.jsonl`：11 起連按 9 次 → 12, 14, 16, 18, 20, **24**, 26, 28, 36（缺 22，Excel 清單有 22【推論】）；縮小 → 28, 26, 24, 20, 18, 16, 14, 12, 11, 10, 9, 8；連點兩下 11→14（無累加問題）。
- 混合字級：A1=10、A2=20 同選一次放大 → 兩格都變 11（以 echoSize=第一格為 base 統一套用；Excel 為各格各走一級）。
- 附帶發現：多範圍選取（A1:A3＋C1:C3）按增大字型，只有最後一個範圍 C1:C3 變 12（`sheets2.jsonl`）。

**PDF：N/A**（只有數字輸入框，`App.tsx:7776-7792`，無增大／縮小按鈕）。**Markdown：N/A**（無字型大小 UI）。**HTML：低**，浮動工具列 `apps/html/src/renderer/components/FloatToolbar.tsx:62-63,116,129` 為線性 ±2px、最小 8px，非 Office 梯度；使用者每按一下都有可見變化，屬設計差異（未 GUI）。

### 4. 用滑鼠選取表格時文字被移到其他欄

**重要限制：** 這個測試環境無法讓瀏覽器「原生」啟動拖放。【GUI】在 Electron 與無頭 Chromium 都試過：一般段落裡「已選取字詞 → 按住拖曳」的對照組，Playwright 的真滑鼠不會觸發 `dragstart`（只有 mousedown／mouseup）。所以下列拖放結果是【合成 DnD】：在 DOM 上依序送出 `dragstart → dragenter → dragover → drop → dragend`，由 ProseMirror 的真實 handler 處理，對照組（段落內搬字）確實搬動了字，證明序列有效。**「真實滑鼠是否一定會在這些情況啟動原生拖放」未驗證。**

**Docs：有問題（條件式）。** 【Chromium＋合成 DnD】`dnd-synth.jsonl`：
- 已選取 A1 全部文字，在選取上開始拖、放到 B2：A1 變空，文字被接到 B2 尾端（`EchoAlpha one two`）；放 C1、B1 同理。一次 Ctrl+Z 可還原。
- 先拖出 2x2 CellSelection（A1:B2）：DOM selection 非空（文字「Echo」，`selectedCells=4`），從該格文字開始拖到 C3：A1:B2 內容被清空，整個 2x2 切片被貼進 C3 一帶，結果列的欄數錯亂（`||Charlie|| / ||Foxtrot|| / Golf|Hotel|India|Alpha one two|Bravo / Delta|Echo|||`）。一次 Ctrl+Z 可還原。
- 【GUI 真滑鼠】沒有既有選取時拖曳選取（A1 → C2、A1:B2 再拖到 C3）：內容不變，只產生 CellSelection（`docs2.jsonl` Q4.a/c）。所以「單純拖曳選取」本身不會改內容，要有「按在已選取文字上」才會變成搬字。
- 程式碼：沒有任何 dragstart 防護。`apps/docs/src/renderer/App.tsx:1482` 的 `handleDrop` 只處理圖片檔；`editor/table-handle.ts:140-149` 只處理表格移動把手。ProseMirror 預設 dragstart 對 CellSelection 會取 `selection.content()`（整塊儲存格切片）。

**Markdown：有問題（條件式）。** 【Electron＋合成 DnD】`md-dnd.jsonl`：雙擊選取 A2 的「Alpha」後拖到 C4：A2 變 `one two`，C4 變 `thirdAlpha`；CellSelection（A2:B3）從 B3 文字拖到 C4：表格被撕壞（`Name|Qty|Note|| / ||first|| / ||second|| / Charlie|7|third|Alpha one two|3 / Bravo|5|||`），GFM 存檔會輸出壞表格。一次 Ctrl+Z 皆可還原。程式碼同樣沒有防護（`apps/markdown/src/renderer/editor/extensions.ts:83-86` 表格設定；`localImage.ts:150` 的 `handleDrop` 只處理圖片檔）。

**Sheets、Slides：原問題（4a、4b）。** **PDF、HTML：N/A**（無表格範圍選取；HTML 預覽的拖曳是刻意的「移動元素」功能）。

### 5. 多物件操作 undo 粒度

**Docs：無（GUI）。**
- 【GUI】`docs2.jsonl`：2x2 儲存格範圍上的粗體、字級 20、儲存格底色、置中、Delete 清空、合併、刪兩列、直接輸入 X，每個動作一次 Ctrl+Z 都完整還原。
- 【Chromium】`docs3-web.jsonl`：4 段落套 Heading 1、粗體、置中、項目符號；尋找取代「全部取代」4 處，皆一次 Ctrl+Z 還原。
- 多選圖片／形狀／表格的方向鍵移動：N/A，Docs 沒有物件多選（2026-10-20 報告 D-20：Shift＋點兩張圖片仍是單一 NodeSelection，方向鍵移動游標而非物件；本次未重測）。
- 備註：2026-10-20 報告 D-02（多格選取段落對齊只套最後一格）本次只確認「一次 Undo 可還原」，未重驗是否每格都有套用。

**Sheets：undo 粒度無問題。** 【GUI】`sheets.jsonl`：複製 A1:D3 貼到 F1、填滿向下、排序、多範圍 Delete、多範圍清除內容，皆 1 次 Undo 還原。程式有 `undo-group.ts`（`runAsOneUndoStep`）把一個指令的多筆 Univer 指令合併。另有一個非 undo 的問題：多範圍時功能區指令只作用於最後一個範圍（見第 3 項；粗體按鈕在多範圍下的探針結果為「無變化」，原因未查明，未驗證）。

**PDF：無（讀碼）。** 所有編輯走 `applyEditOps`，註解寫明「One batch = one undo step」（`App.tsx:1761-1782`），`pushUndo` 存整份快照（`:1734-1740`）；套用標記跨頁、旋轉多頁、刪留言串都是單一 op 批次（`:2258-2285`、`:3736-3741`、`:4915-4935`）。PDF 編輯器沒有多選註解／物件。GUI 只確認縮圖旋轉後 Undo 鈕可用，未做完整的 Undo 還原比對。

**Markdown：無。** 【GUI】`md.jsonl`：全選粗體、全選斜體、輸入文字、清單三項切換、表格浮動列插入下方列、刪除欄，皆一次 Ctrl+Z 還原 DOM。

**HTML：無（讀碼）。** 修改以 `applyPatches` 在 CodeMirror 單一交易內完成（`apps/html/src/renderer/source/SourceEditor.tsx:14-16`；呼叫處 `App.tsx:414`）；單元素選取，無多選。

## 「有問題」項目：位置、重現、建議、影響

### A. Slides 表格右鍵選單過長且不能捲動（問題 2 的反向）
- 位置：`apps/slides/src/renderer/styles.css:5227-5236`（`.ctx-menu` 無 max-height／overflow）、`apps/slides/src/renderer/components/ContextMenu.tsx:26-34`（只夾 y）。
- 重現：視窗高度 760px，插入表格，在儲存格上右鍵：選單 887px，底部 127px 看不到。
- 修法：`.ctx-menu { max-height: calc(100vh - 8px); overflow-y: auto; }`，並讓 `ContextMenu.tsx` 的夾位用 `Math.min(h, innerHeight - 8)`；或把表格項目拆成子選單（Insert／Delete／Align）縮短。
- 影響：僅 Slides；表格與多選時的長選單。Slides 另一個子代理正在改該區，建議交給它處理。

### B. PDF 縮圖選單貼底被切 8px
- 位置：`apps/pdf/src/renderer/App.tsx:6999-7006`（寫死 `innerHeight - 190`）。
- 修法：改量測實際高度（ref＋`useLayoutEffect`，同 Slides 作法），或把常數改為 `innerHeight - 206`。
- 影響：只有 PDF 縮圖列最下方幾張；很輕微。

### C. Docs 增大／縮小字型的兩個缺口
- 位置：`Ribbon.tsx:510-513`（梯度表）、`:1659-1670`（步進）、`font-size-step.ts:34-35`（單一 base）。
- 重現：全選含 10pt 與 20pt 的兩段文字按「增大字型」→ 兩段都變 10.5pt（Word 為 10.5 與 22）；72pt 文字按增大不變（Word 為 80）。
- 修法：梯度表在 72 之後改為 +10（到 1638 封頂），5 以下改為每次 −1（到 1 封頂）；步進改為對選取內每個 `docTextStyle` run 各自找下一級（`state.doc.nodesBetween` 迭代 run，一次 transaction 寫入）。
- 影響：Docs 字型按鈕與 `⇧⌘.`／`⇧⌘,`；`ReturnType<createFontSizeStepper>` 的測試 `apps/docs/tests/font-size-step.test.ts` 需要補混合字級案例。

### D. Sheets 增大／縮小字型：缺 22 級、混合字級壓平
- 位置：`ExcelShell.tsx:174-175`、`:192-199`、`:3037`、`:3133/3140`。
- 重現：11pt 連按增大 → …20 → 24（缺 22）；A1=10、A2=20 同選按一次增大 → 都是 11。
- 修法：梯度補 22；放大／縮小改為逐格套用（用各格自己的字級走一級），而非 `font-size:${stepFontSize(echoSize, …)}` 統一值。
- 影響：Sheets 功能區兩顆按鈕；與 Sheets 子代理正在改的檔案是否重疊請先確認（`ExcelShell.tsx` 目前不在 git 變更清單內）。

### E. Docs／Markdown：拖放造成文字搬移與表格結構損壞
- 位置：Docs `apps/docs/src/renderer/App.tsx:1415-1490`（editorProps 只有圖片 `handleDrop`）；Markdown `apps/markdown/src/renderer/editor/extensions.ts:83-86`（表格設定）與 `localImage.ts:150`。
- 重現（合成 DnD，見第 4 項限制）：在表格內先選取某格文字（雙擊或三擊）或拖出儲存格範圍，再按在選取文字上拖到別欄 → 文字被搬走；CellSelection 被拖時整塊切片貼入目標格，表格列欄數錯亂。一次 Ctrl+Z 可還原。
- 建議修法：在 editorProps 加 `handleDOMEvents.dragstart: (view, e) => { if (view.state.selection instanceof CellSelection) { e.preventDefault(); return true } return false }`，並在 `handleDrop` 檢查被拖的 slice 含 `tableRow`／`tableCell`／`tableHeader` 時，若落點在表格內就吞掉（`return true`）。單一儲存格內文字拖移若要保留 Word 行為，可只允許在同一儲存格內搬動。兩個 App 各加一處，約 20 行，共用一個 helper 放 `packages/ui`。
- 影響：Docs、Markdown 所有表格；Sheets／Slides 另案（原問題 4a／4b）。

## 誠實的限制

1. 第 4 項的「拖放」全是合成事件，不是原生拖放；真滑鼠在 Electron／Chromium 都無法在此測試環境啟動原生拖放（對照組也失敗）。真實使用者是否在相同點擊情境下啟動拖放、以及「CellSelection 的 DOM selection 只涵蓋最後一格文字」是否足以啟動原生拖放，未驗證。
2. Docs 的問題 1–5（含 Q3、Q4 的 Docs 結果）大部分在 Electron 與 Chromium 兩邊都跑過（`docs1.jsonl`／`docs1-web.jsonl` 一致）；Q4 的合成 DnD、Q5 的 Heading／取代全部、Q2 子選單寬度只在 Chromium（`*-web.jsonl`，vite dev server + 假 preload，非 Electron），已於結束時關閉 dev server。
3. Slides 量測的是另一個子代理 10-21 21:37 的中途建置；Sheets 是 21:09 的建置。結果可能在它們修完後改變。
4. Excel／PowerPoint 的增大字型梯度無法用 App 實測（Excel 無 AppleScript 指令、System Events 按鍵被拒），Excel 的 22 級屬推論；PowerPoint 梯度採使用者提供清單，未另外驗證。Word 梯度是實測。
5. PDF、HTML 的第 5 項、HTML 的第 3 項只有讀碼；PDF 縮圖旋轉的 Undo 還原未做完整 GUI 比對；HTML 完全沒開 GUI。
6. Sheets 多範圍粗體探針結果為「無變化」，不確定是按鈕選擇器問題還是真的沒套用；多範圍字型只套最後一個範圍則是明確量到。
7. Docs 圖片／形狀的多選與方向鍵移動只引用 2026-10-20 報告，本次未重測；Docs D-02（多格段落對齊）未重驗。
8. 使用 Electron 前取得鎖花了很久（另兩個子代理持續持有），實際取得後每個規格各取放一次；曾用 `pkill -f` 結束自己的等待迴圈，這可能也誤殺了其他代理同樣指令的等待迴圈（它們的指令列相同），請留意是否有其他代理看到「等鎖失敗」。目前鎖已釋放、無殘留 Electron 程序。
