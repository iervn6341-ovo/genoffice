# Slides 物件選取／操控／格式設定 與 PowerPoint for Mac 差異評估（2026-10-20）

範圍：GenOffice Slides（分支 `qa/local-fixes`，未 commit）對照 Microsoft PowerPoint for Mac 16.113.3（Microsoft 365，繁體中文介面）。
本文件同時記錄本輪完成的兩項修正（問題 1：選取文字後改格式選取消失；問題 2：表格儲存格範圍選取），最後附給後續實機驗證者（Codex）的「驗證手冊」。

## 0. 觀察方式與可信度（務必先讀）

| 標記             | 意思                                                                                                                                                                                                                    |
| ---------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **探針實測**     | GenOffice 以 Playwright Electron、前景視窗、真實滑鼠／鍵盤事件實測（`e2e/_probe/`，結果在本文各表）                                                                                                                     |
| **PPT 實機**     | 2026-10-20 在本機 PowerPoint for Mac 16.113.3 以 AppleScript 物件模型＋無障礙（AX）背景觀察，新增空白簡報「簡報1」，結束後已關閉且未儲存（未碰任何使用者檔案）                                                          |
| **PPT 已知行為** | 依 PowerPoint 365 Mac 的既有操作行為（記憶／官方文件），**本輪未實機驗證**；原因：背景 AX 無法在投影片畫布上點擊、拖曳、送出方向鍵、開右鍵選單，電腦全螢幕控制多次被「user interrupt」中止。列為 Codex 驗證手冊的待驗項 |

其他限制：PowerPoint 的圖表／SmartArt／圖片／媒體插入需要啟動 Excel 或檔案授權對話框（選單項在背景為停用、AppleScript 開檔會跳授權視窗，已取消），因此這四類只有 GenOffice 側探針資料，PowerPoint 側為已知行為。
原生色彩選取器（NSColorPanel）與原生 `<select>` 彈窗無法用 Playwright 驅動，本輪**未驗證**，列在 §6 風險。

## 1. 問題 1：選取文字後改字型等格式，選取被取消

### 1.1 重現（為何 Playwright「單純點擊」測不到）

`e2e/_probe/slides-sel1.probe.spec.ts`（`launchShell({ foreground: true })`，所有操作用 `page.mouse` 真實 down/up、多段 move，鍵盤用 `page.keyboard`）。修正前實測結果：

| 流程                                                       | 修正前                               |
| ---------------------------------------------------------- | ------------------------------------ |
| 字級／字型：只點箭頭（caret）再選項目                      | 選取保留（與使用者先前測的相同）     |
| 字級輸入框打值＋Enter、字型輸入框打字＋Enter               | 選取保留                             |
| **先點字級輸入框（游標進框），再點箭頭選項目**             | **編輯結束（`edit:false`）選取消失** |
| **先點字型輸入框，再點箭頭選字型**                         | **編輯結束**                         |
| **字型框打 `Ari`，再點清單項目**                           | **編輯結束**                         |
| **點字級框後再按 Bold 按鈕**                               | **編輯結束**                         |
| Bold／斜／底線、⌘B/I/U、增大字級、對齊、雙擊選字、三擊選段 | 保留                                 |

### 1.2 根因

`TextEditOverlay` 內有一個 document 層 `mousedown`（capture）監聽：只要**焦點不在編輯器**（例如停在功能區字級／字型輸入框）且按下的位置不在 `[data-keep-edit]` 內，就 `commit()`＝結束編輯。字級／字型下拉清單、Bold 等按鈕沒有標 `data-keep-edit`，所以「先點輸入框（焦點停放）→ 再點清單項目」會被誤判成「點到編輯器外」。單純點箭頭時按鈕 `preventDefault` 讓焦點留在編輯器，故不會觸發，這就是 Playwright 只做「單純點擊」測不出來的原因。同一機制還造成：

1. 點功能區／窗格的空白處、標籤文字（非可聚焦元素）→ 焦點掉到 body → 編輯器 `blur(relatedTarget=null)` → 提交。
2. 格式窗格操作（文字方塊對齊、自動調整、邊界…）會 `applySlide`，`editNode` 變新物件 → 編輯器被**重新填入**：選取退到文末，且尚未提交的輸入內容被丟掉。
3. 焦點停放在輸入框時，編輯器文字的反白消失（PowerPoint 會保留反白）。
4. 字級／字型框按 Esc 後焦點落在 body，下一次點擊就提交。
5. 新增儲存格選取模式後發現：`CmField` 按 Enter 會先 `blur()`，同一個 keydown 往 window 冒泡時 `activeElement` 已是 body，被全域鍵盤處理當成「表格上按 Enter」。

### 1.3 修正（共用、不綁死 Slides）

新增 `packages/ui/src/keep-edit.ts`（由 `@genoffice/ui` 匯出），供任何 contentEditable 編輯器使用：

- `KEEP_EDIT_ATTR`（`data-keep-edit`）：把功能區根、任務窗格、文字對話框標成「中性區」。
- `classifyEditPress(frame, target)`：把一次按下分成 `in-editor / neutral / neutral-field / outside`。
- `installKeepEditGuard({ editor, frame?, keeper, commit })`：一個編輯 session 的 capture 按下守衛。中性區內的**按鈕／選單項／空白**：焦點在編輯器就 `preventDefault` 保持；焦點停放在欄位則先 `keeper.restore()` 把焦點與選取還給編輯器；**欄位**（input/select/textarea/label 內 input/contenteditable）讓欄位自己取得焦點；只有**所有中性區之外**的按下才提交。
- `createSelectionKeeper(name)`：選取快照（`remember/save/restore/set/peek/forget/clearPaint`），停放期間用 CSS Custom Highlight API（`::highlight(name)`）把保留的選取畫出來。
- `handleEditorBlur(e, keeper)`：編輯器 blur 若移往中性區，記住選取並吸收這次 blur（不提交）。

Slides 端整合（`apps/slides/src/renderer/`）：

- `TextEditOverlay.tsx`：改用上述共用模組（仍匯出 `classifyEditPress / saveEditSelection / restoreEditSelection / rememberEditSelection / forgetEditSelection` 供既有呼叫端與測試）；填入 effect 改成「同一個編輯 session（sourceId／儲存格、zoom norm、caret、replaceWith）內，模型刷新時：**使用者已改動 → 保留 DOM 與選取**；未改動 → 重新填入並用字元位移把選取帶過去」。
- `Ribbon.tsx`、`components/FormatPane.tsx`：根元素加 `data-keep-edit=""`。
- `RibbonHomeTab.tsx`：字級／字型框 Esc → `restoreEditSelection()`（焦點與選取回編輯器）。
- `styles.css`：`::highlight(go-kept-selection)`（`Highlight / HighlightText` 系統色，無寫死 hex）。
- `keyboard-actions.ts`：`inTextField(e.target)` 也看事件目標（自己 blur 的欄位不再洩漏按鍵）；`RibbonTableTabs.tsx` 的 `CmField` Enter／Esc 加 `preventDefault`。

**是否可套用到 Sheets／Docs**：可以，邏輯與 Slides 無關。Sheets 側回報的「功能區按鈕點完焦點留在按鈕、儲存格／圖形內編輯時用功能區改格式選取與焦點消失」是同一類根因。做法：(1) 功能區根與任務窗格加 `data-keep-edit=""`；(2) 編輯器掛載時 `installKeepEditGuard({ editor: () => el, keeper, commit })`，卸載時呼叫回傳的清理函式；(3) 編輯器 `onBlur` 先 `handleEditorBlur(e, keeper)`、回傳 true 就不提交，`onFocus` 呼叫 `keeper.clearPaint()`；(4) 套用格式前呼叫 `keeper.restore()`；(5) 在該 app 的 CSS 加 `::highlight(<name>)`。Sheets 的編輯面若是 Univer 自己的 canvas 編輯器而非 DOM contentEditable，則只能用第 (1)(2) 的按下守衛概念，需另行接 Univer 的焦點 API（本輪未碰 Sheets／Docs）。
`keepEditorFocusOnRibbonPress`（`ribbon-focus.ts`）仍保留，是只擋按鈕焦點的舊版。

### 1.4 修正後驗證（探針實測）

- `slides-sel1`：上表所有流程皆 `edit:true`、選取字串完整；新增雙擊選字、三擊選段、真實拖曳選取（`slides-sel7`）後再操作皆保留。
- `slides-sel3`：停放焦點時反白仍顯示（`CSS.highlights` 大小 1，選取後清為 0）、字級／字型實際套用到選取（24px／Georgia）、字級框 Esc 回編輯器。
- `slides-sel6`：開格式窗格、已輸入尚未提交的文字，在窗格文字方塊「垂直對齊／自動調整」操作後，仍在編輯、選取與輸入內容保留（修正前會被重新填入）。
- 整物件選取（未進入編輯）改格式：Bold、字級、字型、對齊後 Konva Transformer 仍為 1，連續兩次操作正常。
- 單元測試：`apps/slides/tests/edit-keep-selection.test.ts`（9）、`packages/ui/tests/keep-edit.test.ts`（7）。

## 2. 問題 2：表格儲存格範圍選取

### 2.1 設計

三個狀態：**物件選取**（Transformer 控制點，沿用）→ **儲存格範圍選取**（新）→ **儲存格文字編輯**（沿用）。

- 資料：`TableCellSel { sourceId, anchor, focus }`（App `cellSel`）。座標用渲染樹的 `row`＝a:tr 索引、`col`＝a:tc 索引（等於格線欄）。高亮區塊＝anchor 與 focus 的矩形，並**擴張到不切斷任何合併儲存格**（`table-selection.ts` 的 `expandToMerges`）。
- 純函式模組 `table-selection.ts`：`selectionRect / cellsInRect / rectBounds / rowSel / colSel / allSel / stepFrom / moveSel / extendSel / tabTarget / clampSel / rangeToTsv / parseTsv`。
- 視覺與指標：`components/TableCellLayer.tsx`。與表格同旋轉／翻轉的 DOM 層，每格一個隱形命中區（四周留 5px 縫隙，讓 Konva 欄寬／列高拖曳點與 8 個變形控制點仍可點）；選取區半透明底＋外框，顏色來自新的 `canvas-colors.ts`（依投影片底色明暗取淺／深，不寫死在元件，CSS 無新增 raw 色）；表格左緣每列、上緣每欄、左上角各有一條按下即選整列／整欄／整表的把手（hover 才顯示）。
- 鍵盤：`table-select-actions.ts` 的 `handleCellSelectionKey`（在 `keyboard-actions.ts` 的 `editing||inField` 之後、全域快速鍵之前）。
- 作用於範圍的指令：新增 `cells` 參數貫穿 IPC → pptx-ops → pptx-engine。

### 2.2 實作範圍

| 項目             | 行為                                                                                                                                                                                                                                                                                          |
| ---------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 選取方式         | 表格已被選取後：按一下格＝選該格；拖曳＝矩形範圍；Shift＋點擊延伸；Shift＋方向鍵延伸焦點角；方向鍵收合並移動；列／欄／左上角把手（可拖曳、可 Shift 延伸）；⌘A＝全部儲存格；⌘／⌥＋方向鍵在此模式不會位移表格                                                                                   |
| 與文字編輯切換   | 雙擊／Enter／F2 進入該格編輯；直接打字＝以該字取代並進入編輯；編輯中 Esc／提交＝回到「該格被選取」；格選取時 Esc＝回到整表物件選取                                                                                                                                                            |
| Tab／Shift＋Tab  | 格選取時依閱讀順序移動（跨合併格一步）；最後一格 Tab 自動加一列；編輯中的 Tab 導覽沿用原行為                                                                                                                                                                                                  |
| 設計分頁         | 底紋、框線（內／外框等 `borderMode` 由引擎對範圍外接矩形處理）作用於所選儲存格                                                                                                                                                                                                                |
| 版面配置分頁     | 插入上／下方列、左／右方欄＝**一次插入選取範圍跨越的列／欄數**；刪除列／欄＝刪掉所選的列／欄；合併儲存格（新增 `merge-range`，任意矩形）／分割（`split-range`）；列高／欄寬欄位套用到所有選取的列／欄；平均分配列高／欄寬僅作用於選取的列／欄（跨多列／欄時）；水平／垂直對齊作用於所選儲存格 |
| Home 分頁        | 粗／斜／底線／刪除線／上下標、字型、字級、增減字級、顏色、螢光筆、大小寫、字元間距、項目符號、行距、水平對齊、文字格式——全部只作用於所選儲存格（B/I/U 切換依所選儲存格判斷全開／全關）；⌘B/I/U 同                                                                                             |
| 清除             | Delete／Backspace＝清空所選儲存格文字、保留格式（新 op `clearTableCells`），**不會刪表格**                                                                                                                                                                                                    |
| 複製／剪下／貼上 | ⌘C／⌘X 把範圍轉成 TSV 寫入系統剪貼簿（與 Excel／Word 範圍互通）；⌘V 以選取左上格為起點貼入（可貼 Excel 範圍、含引號換行；最多 50 格，超過提示 `appStatusTablePasteLimit`）；一次一步復原                                                                                                      |
| 右鍵             | 在選取範圍內右鍵＝保留範圍並對範圍操作（多格時選單出現「合併儲存格」，插入／刪除／對齊／底紋皆為範圍版）；在範圍外右鍵＝只選該格                                                                                                                                                              |
| 復原／重做       | 每個指令一個 transaction＝一步（探針驗證：清除兩格、⌘Z 一次還原、⇧⌘Z 重做）                                                                                                                                                                                                                   |
| 存檔重開         | 合併／分割／插入／清除／範圍格式皆為 XML 手術，`packages/pptx-engine/tests/table-range.test.ts` 驗證 `savePptx → openPptx` 往返                                                                                                                                                               |
| 旋轉／翻轉表格   | 圖層與表格同旋轉；探針於 25° 旋轉表格拖曳選到正確的 2×3 區塊                                                                                                                                                                                                                                  |

### 2.3 限制（刻意不做或尚未做）

1. **第一次點未選取的表格＝選整個表格物件**（沿用，可拖曳搬移）；要點第二下才進入儲存格選取。PowerPoint 第一下就把插入點放進格子。
2. 範圍複製只帶**純文字**（TSV），不帶儲存格格式／底色；貼上到範圍沿用目的格格式；貼上上限 50 格。
3. 含合併儲存格的表格仍**拒絕插入／刪除列欄**（引擎既有限制，狀態列提示）；PowerPoint 可以。
4. 沒有 PowerPoint「選取」下拉（選取儲存格／欄／列／表格）、檢視格線、儲存格邊界、表格內文字方向、手繪表格／橡皮擦、表格效果。
5. 範圍底紋對合併格只寫入錨點格（渲染以錨點為準）。
6. 群組內的表格不顯示選取圖層（座標系不同）；表格在簡報放映／縮圖不受影響。
7. 復原後元素選取會被清掉（既有行為），需再點一次表格。
8. 沒有拖曳移動儲存格內容、沒有 Shift＋點擊跨表格邊緣的延伸捲動。

### 2.4 探針實測結果

`slides-tbl2`：30 項全部 PASS（單格、Shift 延伸、Shift＋方向鍵、方向鍵收合、列／欄／全表把手、拖曳列把手、輸入進入編輯、Esc 回選取、Bold 只作用於 2×2、⌘B 反轉、置中、底紋整列、底紋後選取保留、Delete 清除、表格不被刪、復原一步、重做、貼上區塊、合併、合併後選取、分割、插入兩列、插入後選取、刪除兩列、Esc、Enter 編輯、Esc 回選取）。
`slides-tbl3`：右鍵（範圍內保留／範圍外只選該格、選單出現「合併儲存格」、選單插入兩列）、欄寬拖曳點仍可用、列高欄位套用到兩列、距離分配後選取保留。
`slides-tbl4`：旋轉 25° 的表格拖曳選取。

## 3. 非文字物件行為評估

結果欄：**一致／有差異／缺少**。優先序：P0 影響日常編輯、P1 常用、P2 少用。「依據」標示是探針實測、PPT 實機或 PPT 已知行為。

### 3.1 形狀

| #   | 項目                                                       | 結果         | GenOffice 觀察（探針實測）                                                                                                                                                   | PowerPoint 依據                                                                                                                                                                                                                      | 優先                         |
| --- | ---------------------------------------------------------- | ------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | ---------------------------- |
| S1  | 單選、Shift／⌘＋點擊多選、再點取消                         | 一致         | 點 A＝1；Shift＋點 B＝2；Shift＋再點 B＝1；⌘＋點 C＝加選                                                                                                                     | PPT 已知行為                                                                                                                                                                                                                         | —                            |
| S2  | 框選（完全包含才選取）                                     | 一致         | 起點在空白處拖出包住 A、B 的框＝選到 2 個                                                                                                                                    | PPT 已知行為                                                                                                                                                                                                                         | —                            |
| S3  | ⌘A 全選                                                    | 一致         | 全選所有可變形物件（連接線無變形把手，不計）                                                                                                                                 | PPT 已知行為                                                                                                                                                                                                                         | —                            |
| S4  | Tab／⇧Tab 循環                                             | 一致         | 依 z 順序循環，每次選 1 個                                                                                                                                                   | PPT 已知行為                                                                                                                                                                                                                         | —                            |
| S5  | 方向鍵微調                                                 | 有差異       | 方向鍵＝1px、⇧＋方向鍵＝10px；⌘／⌥＋方向鍵與方向鍵相同（沒有「微調」層級）                                                                                                   | PPT 已知行為：方向鍵依格線／預設步距移動，另有細微調修飾鍵；**待實機**量測移動距離                                                                                                                                                   | P2                           |
| S6  | 8 個控制點縮放（自由／Shift 等比／Alt 由中心／Alt＋Shift） | 一致         | 孤立矩形：Shift＝比例 2.4 維持；Alt＝中心不動；Alt＋Shift 兩者皆成立（有其他物件時「同尺寸吸附」可能讓比例略偏，屬刻意功能）                                                 | PPT 已知行為                                                                                                                                                                                                                         | —                            |
| S7  | 旋轉把手：Shift 15°                                        | 一致         | 指標在 37°＋Shift＝30°                                                                                                                                                       | PPT 已知行為                                                                                                                                                                                                                         | —                            |
| S8  | 旋轉：不按 Shift 的自由旋轉                                | 有差異       | 無 Shift 時指標 37°＝36.5°（連續），但靠近 45° 倍數 ±5° 會自動吸附（44°→45°）                                                                                                | PPT 已知行為：自由旋轉不自動吸附；**待實機**                                                                                                                                                                                         | P2                           |
| S9  | 黃色調整把手                                               | 一致（部分） | 圓角矩形選取時出現 1 個黃色把手；`adjust-handles.ts` 不涵蓋 gear6/9、mathNotEqual、circularArrow                                                                             | PPT 已知行為：幾乎所有預設圖案都有                                                                                                                                                                                                   | P2                           |
| S10 | 群組／取消群組、進入群組內選取                             | 一致         | ⌘G 群組 3 個形狀＝1 個群組；雙擊進入群組並選到子物件（選取數 1）；Esc 離開；⌘⇧G 取消。表格／圖表不能被群組（PPT 圖表可群組、表格不行）                                       | PPT 已知行為                                                                                                                                                                                                                         | P2（圖表群組）               |
| S11 | 排列順序                                                   | 一致         | 右鍵選單與功能區有移到最上／最下／上移／下移                                                                                                                                 | PPT 實機：圖案格式分頁有「上移一層／下移一層／選取範圍窗格」                                                                                                                                                                         | —                            |
| S12 | 選取範圍窗格（Selection Pane）                             | 缺少         | 無                                                                                                                                                                           | PPT 實機：圖案格式分頁有「選取範圍窗格」按鈕                                                                                                                                                                                         | P1                           |
| S13 | 對齊與分散                                                 | 有差異       | 有對齊（左中右上中下）與水平／垂直分散（需 3 個以上）；**沒有「對齊投影片／對齊選取的物件」切換**                                                                            | PPT 已知行為：對齊選單含「對齊投影片」與「對齊所選物件」                                                                                                                                                                             | P1                           |
| S14 | 複製貼上、⌘D、Option 拖曳複製                              | 一致         | ⌘C⌘V＋1、⌘D＋1、Option 拖曳＋1（物件數 8→9→10→11）                                                                                                                           | PPT 已知行為                                                                                                                                                                                                                         | —                            |
| S15 | 鎖定長寬比                                                 | 一致         | 格式窗格「大小與屬性」有「Lock aspect ratio」                                                                                                                                | PPT 實機：圖案格式分頁大小區有高度／寬度與鎖定長寬比                                                                                                                                                                                 | —                            |
| S16 | 格式窗格選項                                               | 有差異       | 圖案選項：填滿與線條／效果／大小與屬性；文字選項：文字方塊（對齊、方向、自動調整、邊界）可用，**文字填滿、文字外框、文字效果顯示為「尚未可編輯」**                           | PPT 已知行為：文字選項有文字填滿與外框／文字效果／文字方塊                                                                                                                                                                           | P1                           |
| S17 | 選取形狀時的情境功能區分頁                                 | 有差異       | 分頁名「Shape Tools」；內容只有 Change Shape、樣式庫、Fill、Outline、Effects。**文字方塊（text box）選取時完全沒有情境分頁**                                                 | PPT 實機：分頁「圖案格式」（形狀、文字方塊、連接線都有），群組含插入圖案（圖庫、文字方塊、編輯圖案、合併圖案）、圖案樣式、文字藝術師樣式、替代文字、排列（上移／下移／選取窗格／對齊／組成群組／旋轉）、大小（高度／寬度）、格式窗格 | P1                           |
| S18 | 右鍵選單（形狀）                                           | 有差異       | 剪下複製貼上、排列、旋轉／翻轉、編輯文字、變更圖案、超連結、對齊、格式化圖案、刪除；**無浮動迷你工具列（字型／填滿／外框）**、無「儲存為圖片」「大小與位置」「設為預設圖案」 | PPT 已知行為（右鍵時上方出現迷你工具列）                                                                                                                                                                                             | P1（迷你工具列）、P2（其餘） |
| S19 | 文字方塊選取時的右鍵／格式化                               | 有差異       | 選單有「Format Text Box」，但缺情境分頁（見 S17）                                                                                                                            | 同 S17                                                                                                                                                                                                                               | 併入 S17                     |

### 3.2 圖片

| #   | 項目                     | 結果   | GenOffice 觀察                                                                                                                       | 依據                                                                                           | 優先 |
| --- | ------------------------ | ------ | ------------------------------------------------------------------------------------------------------------------------------------ | ---------------------------------------------------------------------------------------------- | ---- |
| P1  | 圖片格式分頁             | 有差異 | 「Picture Format」：校正、色彩、移除背景、取代圖片、重設圖片、透明度、外框、效果、樣式、圖片品質（壓縮）、替代文字、裁剪、旋轉／翻轉 | PPT 已知行為：另有美術效果、圖片版面配置、裁剪下拉（形狀／長寬比）、邊框／效果、排列與大小群組 | P1   |
| P2  | 裁剪模式                 | 一致   | 按「Crop」進入裁剪覆疊層（拖曳把手、Enter 確認、Esc 取消）；右鍵選單有「Crop Picture」                                               | PPT 已知行為                                                                                   | —    |
| P3  | 裁剪成圖形／長寬比       | 缺少   | 無「裁剪為圖形」「長寬比」                                                                                                           | PPT 已知行為                                                                                   | P1   |
| P4  | 改變形狀（圖片）         | 缺少   | 右鍵「Change Shape…」僅出現在形狀／文字方塊，圖片沒有                                                                                | PPT 已知行為                                                                                   | P2   |
| P5  | 重設圖片、替代文字、壓縮 | 一致   | 都有（壓縮僅 PNG／JPEG）                                                                                                             | PPT 已知行為                                                                                   | —    |
| P6  | 右鍵選單                 | 有差異 | 無迷你工具列、無「儲存為圖片」「大小與位置」「取代圖片」在右鍵                                                                       | PPT 已知行為                                                                                   | P2   |

### 3.3 圖表

| #   | 項目                           | 結果   | GenOffice 觀察                                                                                       | 依據                                                   | 優先 |
| --- | ------------------------------ | ------ | ---------------------------------------------------------------------------------------------------- | ------------------------------------------------------ | ---- |
| C1  | 圖表工具分頁                   | 一致   | 選圖表出現「Chart Design」「Chart Format」兩個分頁                                                   | PPT 已知行為：「圖表設計」「格式」                     | —    |
| C2  | 選取圖表本身 vs 圖表內元素     | 缺少   | 只能選整個圖表物件；點圖表內的圖例、座標軸、資料列、資料點不會單獨選取（資料點顏色只能在格式窗格改） | PPT 已知行為：點圖表元素可單獨選取並出現元素的格式窗格 | P1   |
| C3  | 圖表元素按鈕（＋／筆刷／濾斗） | 缺少   | 無，改由 Chart Design 的「Add Chart Element」下拉                                                    | PPT 已知行為                                           | P1   |
| C4  | 右鍵選單                       | 有差異 | 僅通用項＋「Format Chart」；無「編輯資料」「變更圖表類型」                                           | PPT 已知行為                                           | P2   |

### 3.4 表格（本輪已實作範圍選取，其餘差異）

| #   | 項目                                                                | 結果           | 說明                                                                                                                         | 依據                                                                                                                                                                 | 優先 |
| --- | ------------------------------------------------------------------- | -------------- | ---------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---- |
| T1  | 拖曳多格、整列整欄、整表、Shift 延伸、方向鍵、Tab、清除、貼上、合併 | **一致（新）** | 見 §2                                                                                                                        | PPT 已知行為                                                                                                                                                         | —    |
| T2  | 第一次點擊表格                                                      | 有差異         | 選整個表格物件；第二下才進入儲存格選取（PPT 第一下即在格內放插入點）                                                         | PPT 已知行為                                                                                                                                                         | P1   |
| T3  | 表格設計分頁                                                        | 有差異         | 有六個樣式選項、樣式庫、底紋、框線、畫筆（樣式／粗細／顏色）；無「效果」、快速樣式（文字藝術師樣式）、手繪表格、橡皮擦       | PPT 實機：標題列／合計列／帶狀列／首欄／末欄／帶狀欄、樣式庫、網底、框線、效果、快速樣式、畫筆樣式／粗細／色彩、手繪表格、橡皮擦                                     | P2   |
| T4  | 表格版面配置分頁                                                    | 有差異         | 有刪除、四個插入、合併／分割、列高／欄寬、平均分配、對齊、表格大小；缺「選取」下拉、檢視格線、文字方向、儲存格邊界、排列群組 | PPT 實機：選取、檢視格線、刪除、插入×4、合併／分割、列高／欄寬、平均分配、對齊×6、文字方向、儲存格邊界、表格大小、排列（上移／下移／對齊／組成群組／旋轉）、格式窗格 | P2   |
| T5  | 含合併儲存格的表格插入／刪除列欄                                    | 有差異         | 被拒絕（狀態列提示）                                                                                                         | PPT 已知行為：可執行                                                                                                                                                 | P1   |
| T6  | 複製範圍保留格式                                                    | 缺少           | 只有純文字 TSV                                                                                                               | PPT 已知行為：保留格式、可貼成新表格                                                                                                                                 | P1   |
| T7  | 拖曳移動儲存格內容、拖曳選取跨邊捲動                                | 缺少           | 無                                                                                                                           | PPT 已知行為                                                                                                                                                         | P2   |

### 3.5 線條／連接線、SmartArt、WordArt、媒體、方程式、預留位置、縮圖窗格與備忘稿

| #   | 項目                     | 結果         | GenOffice 觀察                                                                                                              | 依據                                         | 優先         |
| --- | ------------------------ | ------------ | --------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------- | ------------ |
| L1  | 直線／箭頭選取與端點黏附 | 一致         | 選取後顯示端點把手；端點靠近形狀邊中點會吸附並建立連接（stCxn/endCxn）；無縮放把手與旋轉把手                                | PPT 實機：連接線選取後同樣只出現圖案格式分頁 | —            |
| L2  | 線條情境分頁             | 有差異       | 「Shape Tools」（同 S17 內容）；PPT 為「圖案格式」且有線條樣式、箭頭等群組                                                  | PPT 實機（連接線選取→「圖案格式」）          | P1（併 S17） |
| L3  | 轉折連接線的黃色調整把手 | 待確認       | 未驗證轉折點把手                                                                                                            | PPT 已知行為：有                             | P2           |
| M1  | SmartArt                 | 缺少         | 以「群組形狀」實作，選取時只有「Shape Tools」；無 SmartArt 設計／格式分頁、無文字窗格、無版面切換                           | PPT 已知行為                                 | P1           |
| M2  | WordArt                  | 有差異       | 插入預設文字方塊樣式；無「文字藝術師樣式」群組，文字填滿／外框／效果在格式窗格文字選項標示為尚未可編輯                      | PPT 已知行為                                 | P1（併 S16） |
| M3  | 媒體（影音）             | 有差異       | 雙擊開啟播放覆疊；選取時無「影片格式／播放」情境分頁、無剪輯／書籤／淡入淡出                                                | PPT 已知行為                                 | P2           |
| M4  | 方程式                   | 有差異       | 有「Equation」情境分頁與對話框編輯；PPT 為就地編輯                                                                          | PPT 已知行為                                 | P2           |
| M5  | 預留位置（placeholder）  | 一致（基本） | 空白預留位置顯示灰字提示；單擊進入編輯；版面配置／重設可用；未逐項驗證刪除後還原                                            | PPT 已知行為                                 | P2           |
| M6  | 投影片縮圖窗格           | 一致（既有） | Shift／⌘ 多選、拖曳排序、右鍵選單（新增／複製／刪除／章節）已有 e2e：`e2e/slides-thumb-multiselect.spec.ts`；本輪**未重跑** | PPT 已知行為                                 | —            |
| M7  | 備忘稿窗格               | 待確認       | 有備忘稿編輯器；格式化行為未逐項驗證                                                                                        | PPT 已知行為                                 | P2           |

### 3.6 優先序統計

- **P0（已於本輪修正，2 項）**：①選取文字後改格式選取／編輯被取消（§1）；②表格無法跨儲存格選取、底紋／格式只能逐格（§2）。其餘未發現新的 P0。
- **P1（未修，12 項）**：S12 選取範圍窗格；S13 對齊投影片／所選物件；S16 格式窗格文字填滿／外框／效果；S17／S19／L2 形狀、文字方塊、線條情境分頁內容與文字方塊完全沒有分頁；S18 右鍵迷你工具列；P1 圖片格式分頁內容；P3 裁剪為圖形／長寬比；C2 圖表內元素選取；C3 圖表元素按鈕；T2 表格第一下進格；T5 合併表格插入刪除列欄；T6 範圍複製保留格式；M1 SmartArt。（M2 併入 S16）
- **P2（未修，約 14 項）**：S5 微調層級；S8 自由旋轉吸附；S9 調整把手涵蓋；S10 圖表群組；S18 其餘選單項；P4 圖片改變形狀；P6；C4；T3；T4；T7；L3；M3 媒體分頁；M4 方程式就地編輯；M5；M7。

前幾名建議：(1) 文字方塊補上 Shape Format 情境分頁（`components/context-tabs.ts` 的 `contextElementTypeForNode` 對 `node.type === 'text'` 回傳 null，是造成缺分頁的一行原因；未修，因會牽動分頁自動切換與既有測試，需另行評估）；(2) 右鍵迷你工具列；(3) 表格第一下進格＋範圍複製保留格式；(4) 圖表內元素選取；(5) 選取範圍窗格。

### 3.7 順手修正的小錯

僅有與本輪修正直接相關的小項（均列於 §1.3）：`CmField` Enter／Esc 缺 `preventDefault`、`inTextField` 未看事件目標、字級／字型框 Esc 不還焦點。未做任何報告中其他差異的實作。

## 4. 本輪改動檔案

Slides：`apps/slides/src/renderer/{TextEditOverlay.tsx, App.tsx, action-context.ts, style-actions.ts, table-actions.ts, table-layout-actions.ts, keyboard-actions.ts, context-menu-items.ts, styles.css}`、`components/{Ribbon.tsx, FormatPane.tsx, RibbonHomeTab.tsx, RibbonTableTabs.tsx}`；新增 `table-selection.ts、table-select-actions.ts、canvas-colors.ts、components/TableCellLayer.tsx`；`i18n/app/*.ts`（20 語系各加 `appStatusTablePasteLimit`，zh／zh-TW／ja／ko 翻譯，其餘英文）；`apps/slides/src/shared/ipc.ts`、`apps/slides/src/main/slides-main.ts`（**含 main，需重建 `@genoffice/shell`**）。
共用：`packages/ui/src/{keep-edit.ts, index.ts}`。
引擎／ops（最小改動）：`packages/pptx-engine/src/index.ts`（`merge-range／split-range`、`editTableStructure` 的 `count`、`clearTableCellText`、`setElementFont／setElementParagraphFormat` 的 `cells`）、`packages/pptx-ops/src/ops/{registry.ts, table-ops.ts, text-ops.ts}`、`prompts/ops/{table.md, text.md}`（新 op `clearTableCells`）。
測試：`apps/slides/tests/{edit-keep-selection.test.ts, table-selection.test.ts}`、`packages/ui/tests/keep-edit.test.ts`、`packages/pptx-engine/tests/table-range.test.ts`。
探針（不 commit）：`e2e/_probe/slides-sel-common.ts、slides-sel1…7、slides-hold、slides-tbl1…4、slides-objs、slides-objs2…4`（`.probe.spec.ts`）。

## 5. 驗證結果（2026-10-20）

`npm run test -w @genoffice/slides`：105 檔 1023 通過（1 skipped）；`@genoffice/ui` 59、`@genoffice/pptx-engine` 1141、`@genoffice/pptx-ops` 55 通過；`tsc`（slides／ui／pptx-engine／pptx-ops）無錯；eslint 0 error；prettier、`node tools/check-theme-colors.mjs` 通過；`npm run build -w @genoffice/slides` 與 `-w @genoffice/shell` 成功。

## 6. 未驗證風險

1. 原生色彩選取器（「更多色彩」）與原生 `<select>` 彈窗：Playwright 驅動不了，電腦控制被中斷。理論上色彩輸入屬 `data-keep-edit` 欄位、`onPointerDown` 會先存選取，但視窗失焦（NSColorPanel）時編輯器的 blur 路徑未實機走過。
2. PowerPoint 側的畫布互動（方向鍵步距、右鍵選單內容、自由旋轉是否吸附、表格拖曳選取細節、Select 下拉）為已知行為，需實機確認。
3. 表格範圍選取在高 DPI／縮放極端值、群組內表格、簡報放映中的互動未測。
4. `TextEditOverlay` 的「模型刷新時保留已改動編輯器」在使用者同時用 AI／尋找取代改同一文字方塊時，會以編輯器內容為準（不再被模型覆蓋）；可接受但行為已變。
5. 圖表／SmartArt／圖片／媒體的 PowerPoint 側情境分頁內容為已知行為，未實機展開。

## 7. 驗證手冊（給 Codex，不用看程式碼）

### 7.0 前置

```
cd "/Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice"
npm run build -w @genoffice/slides && npm run build -w @genoffice/shell
GENOFFICE_E2E_NO_VIDEO=1 npx playwright test --config e2e/playwright.config.ts e2e/_probe/<檔名>.probe.spec.ts
```

探針會前景開啟一個 Electron 視窗（約 20～50 秒），輸出 `RESULT`／`CHECK`／`OBS` 行，截圖在 `/private/tmp/claude-501/obj-slides/`。一次只跑一個。實機驗證請用 `npm run shell` 或打包的 app，語言用英文較好對照。

### 7.1 修正 1：選取文字後改格式，選取與編輯狀態保留

| 步驟（真實滑鼠）                                                                                                             | 預期                                                           |
| ---------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------- |
| 新增文字方塊輸入 `Hello world sample text`，⌘A 全選。**點字級輸入框**（游標進框）→ 點字級箭頭 → 選一個字級                   | 文字仍在編輯、仍全選；字級已變；選取在點字級框期間仍以反白顯示 |
| 同上但點**字型輸入框**→ 點字型箭頭 → 選字型                                                                                  | 同上，字型已變                                                 |
| 字型框打 `Ari`（清單過濾）→ 點清單第一項                                                                                     | 編輯與選取保留，字型套用                                       |
| 點字級框後直接按 **Bold** 按鈕                                                                                               | 編輯與選取保留，文字變粗                                       |
| 字級框打 `30` 按 Enter；字型框打 `Arial` 按 Enter                                                                            | 套用到選取，選取保留，焦點回文字                               |
| 點字級框後按 **Esc**                                                                                                         | 焦點與選取回到文字，繼續打字會取代選取                         |
| 雙擊選字／三擊選段／拖曳選字，再做上述任一操作                                                                               | 同上                                                           |
| 螢光筆、字色、項目符號、行距、大小寫、字元間距、⌘B／⌘I／⌘U                                                                   | 選取保留（探針 `slides-sel2` 另有）                            |
| 打開**格式窗格**（Home → Format Pane）→ 文字選項 → 文字方塊 → 垂直對齊選「Middle」、自動調整選「Shrink」，操作前先多打幾個字 | 仍在編輯、選取保留、剛打的字沒有消失（`slides-sel6`）          |
| 只選取整個物件（Esc 一次後）改字型／字級／粗體／對齊，連做兩次                                                               | 物件選取框仍在、格式兩次都套用                                 |
| **（需實機）**「更多色彩」原生色彩選取器選色；右鍵浮動格式列（目前沒有，屬 P1 缺少）                                         | 預期：色彩套用於原先選取且仍在編輯；若選取消失請回報步驟       |

對應探針：`slides-sel1`（修正前會有 4 行 `edit:false`）、`slides-sel2`（調色盤／項目符號／行距等）、`slides-sel3`（反白、實際套用、Esc）、`slides-sel6`（窗格）、`slides-sel7`（拖曳／三擊）、`slides-sel4/5`（整物件）。
單元測試：`apps/slides/tests/edit-keep-selection.test.ts`、`packages/ui/tests/keep-edit.test.ts`。
已知限制：無右鍵浮動格式列；原生色彩選取器未驗；探針中「點第二層長清單項目超出視窗」會點到畫布，屬探針座標問題而非程式問題。

### 7.2 修正 2：表格儲存格範圍選取

準備：Insert → Table → 4×4。先在表格上點一下（選取整個表格），之後在儲存格內操作。

| 操作                                                      | 預期                                                                                                     |
| --------------------------------------------------------- | -------------------------------------------------------------------------------------------------------- |
| 在格內按一下                                              | 該格出現半透明底與外框                                                                                   |
| 從 (1,1) 拖到 (2,3)                                       | 2×3 範圍被選取                                                                                           |
| Shift＋點另一格；Shift＋方向鍵                            | 範圍延伸；方向鍵（無 Shift）收合成單格並移動                                                             |
| 表格左緣某列把手、上緣某欄把手、左上角把手                | 選整列／整欄／全部；沿把手拖曳可延伸多列／多欄；hover 把手會顯示                                         |
| ⌘A                                                        | 全部儲存格                                                                                               |
| 直接打字／Enter／F2／雙擊                                 | 進入該格編輯（打字會取代原內容）；Esc 回到該格被選取；再按 Esc 回到整表物件選取                          |
| 編輯中 Tab／Shift＋Tab；選取模式下 Tab／Shift＋Tab        | 依閱讀順序移動；最後一格 Tab 自動新增一列                                                                |
| 選 2×2 後 Home 按 Bold、⌘B、置中、字色                    | 只有 4 格變化；再按 Bold／⌘B 切換回來                                                                    |
| 選一整列 → Table Design → Shading 選色；Borders → Outside | 只該列變色／框線                                                                                         |
| 選 2×3 → Table Layout → Merge Cells；再 Split Cells       | 合併成一格（文字併入）、分割回 6 格；合併後仍選取該格                                                    |
| 選兩列 → Insert Row Above；Delete → Delete Row            | 一次插入 2 列、選取仍指向原來的格；刪除一次刪 2 列                                                       |
| Row Height 欄位輸入 2.5 Enter（選了兩列）                 | 兩列都變 2.5 cm；選取保留                                                                                |
| Distribute Rows（選了多列）                               | 只平均所選的列                                                                                           |
| 選兩格 Delete／Backspace；⌘Z；⇧⌘Z                         | 文字被清空（格式保留）、表格不會被刪；⌘Z 一次還原兩格；⇧⌘Z 重做（⌘Z 後表格選取會被清掉，需再點一下表格） |
| 選 2×2 ⌘C，點別格 ⌘V；從 Excel 複製範圍後 ⌘V              | 以選取左上格為起點貼入文字；可雙向與 Excel 互通（僅文字）                                                |
| 右鍵：範圍內／範圍外                                      | 範圍內保留範圍並出現「Merge Cells」、插入／刪除／對齊／底紋皆作用範圍；範圍外只選該格                    |
| 欄線上拖曳（欄寬）、表格外框附近拖曳（搬移）              | 欄寬仍可拖；在邊緣 8 個控制點與外緣帶狀區仍是縮放／搬移                                                  |
| 旋轉表格 25° 後拖曳選取                                   | 選取區塊隨表格旋轉且正確                                                                                 |
| 存檔 → 重開                                               | 合併、清除、底紋、字型格式皆保留                                                                         |

對應探針：`slides-tbl1`（拖曳）、`slides-tbl2`（30 項檢查，看 `SUMMARY 0 fail of 30`）、`slides-tbl3`（右鍵／欄寬／列高）、`slides-tbl4`（旋轉）。
單元／引擎測試：`apps/slides/tests/table-selection.test.ts`（14）、`packages/pptx-engine/tests/table-range.test.ts`（10，含存檔往返）、`apps/slides/tests/op-docs*.test.ts`（新 op 文件與範例）。
已知限制：見 §2.3（第一下點擊選物件、純文字複製、合併表格不能插入刪除列欄、貼上上限 50 格、群組內表格不顯示圖層）。

### 7.3 物件行為差異評估的實機補驗項（PowerPoint 側）

請用 PowerPoint for Mac 逐項確認 §3 標「PPT 已知行為」且未驗證的內容，特別是：S5 方向鍵步距與細微調修飾鍵、S8 自由旋轉是否吸附、S13 對齊選單的「對齊投影片」、S18 右鍵迷你工具列與選單項目、T2 表格第一下點擊行為、C2／C3 圖表內元素與元素按鈕、M1 SmartArt 分頁。GenOffice 側資料可用 `slides-objs`、`slides-objs2`、`slides-objs3`、`slides-objs4` 重現（輸出 `OBS` 行）。
