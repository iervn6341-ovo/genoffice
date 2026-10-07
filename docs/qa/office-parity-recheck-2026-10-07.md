# 與 MS Office 對照的完整重測（進行中，2026-10-07 暫停）

目標：把 GenOffice 各應用和本機的 Word／Excel／PowerPoint 逐項比對，不一樣就修正。
這份文件記錄做到哪裡、找到什麼、還剩什麼，下次可直接接續。**所有修改都還沒 commit。**

## 1. 基準與工具

- **完整 e2e 基準：** 204 通過、10 略過、0 失敗（47.2 分鐘），在今天的 Word 對齊修改**之前**量的。之後只重跑了相關檔案，**完整 e2e 還沒重跑**。
- **單元測試（最近一次整套）：**
  - 文件編輯器 2,498 個：2,497 通過、1 個失敗——`tab-key-insert` 的「清單項目中間按 Tab 仍縮排」，Tab 行為改成和 Word 一致後這個測試過時。
  - docx 解析 1,401 個：1,400 通過、1 略過。
  - 之後已把失敗的測試改成 Word 實測的行為（該檔 7/7），**整套還沒重跑**。
- **比對工具（都在 `test/_probe/`，不進 git）：**
  - `diff/wordkit.py`：產生測試文件、轉成標準化結構、逐案比對（`build`／`plan`／`navplan`／`dump`／`diff`）。要用 `/opt/homebrew/bin/python3.12`（系統的 `python3` 是 3.9，會語法錯誤）。
  - `diff/cases_w1..w6.py`、`cases_chk1/2.py`：Word 測試案例。`select.scpt`、`selread.scpt`、`saveword.scpt`、`openword.scpt`、`paras.scpt`、`navstep.sh`：在 Word 設定選取、讀選取、存檔。
  - `e2e/_probe/word-diff.probe.spec.ts`：同一批案例在 GenOffice 執行（`DIFF_DIR`、`DIFF_MODE=edit|nav`、`DIFF_ONLY`）。`docs-eval.probe.spec.ts`／`sheets-eval.probe.spec.ts`：在編輯器裡執行一段程式（`DOCS_JS`／`SHEETS_JS`）。
  - `fx/`：447 個公式案例（`corpus.py`）、`excel_run.scpt`（Excel 的結果在 `excel-out.txt`）、`compare_fx.py`；GenOffice 側用 `e2e/_probe/fx.probe.spec.ts`（`FX_DIR`）。重新比對：`cd test/_probe/fx && /opt/homebrew/bin/python3.12 compare_fx.py`。
- **限制：** 背景模式只能送進 Word 的按鍵是 Enter、Backspace、Delete、Tab、方向鍵、Home／End、空白，以及方向鍵搭配 Shift／Option／⌘（Option＋Backspace／Delete 也可以）。⌘B、⌘Z、⇧F3 這類快速鍵、一般字母鍵送不進去；需要使用者核准全螢幕控制才能測。Word 的檔案存取只對 `test/_probe/` 有授權（子資料夾會跳「授與檔案存取權」）。

## 2. 已測並已對齊 Word（Word 實測 → GenOffice 修正）

| 範圍 | 案例 | 結果 |
|---|---|---|
| 段落結構編輯（Enter／Backspace／Delete、標題、置中、縮排） | 31 | 27 一致；4 個不同只是 Word 自動把句首小寫改成大寫（自動校正，未實作） |
| 游標移動與選取（Option／⌘＋方向鍵、Home／End、Shift 延伸、CJK） | 57 | 55 一致；2 個（`N42`、`N53`）只差一個字元的像素取整 |
| 清單與表格（Enter／Tab／Backspace／Shift＋Tab、表格方向鍵） | 32 | 32 一致（表格第一列按 ↑ 的 `T11` 例外，見第 3 節） |
| 段落開頭按 Backspace 的變體 | 15 | 14 一致；`B01`（靠右對齊）Word 無反應，原因不明 |
| 跨段選取後刪除 | 13 | 12 案一致；`S13`（⌘X 跨段剪下）無法驗證：GenOffice 側用選單加速鍵觸發剪下，在背景模式不會動作 |

已實作（未 commit）：標題 Enter 的下一個樣式、標題開頭 Enter、跨段刪除規則、置中段落 Backspace 先取消置中、懸掛縮排 Backspace、清單 Enter（空項目、子層降級、開頭）與 Tab（中間插入 Tab 字元）、Option＋左右方向鍵（Word 字詞單位，含 CJK）、Option／⌘＋上下方向鍵（段落）、Shift＋End 包含段落標記、段落上移下移改為 ⌃⇧↑↓。檔案：`apps/docs/src/renderer/editor/word-keys.ts`（新）、`extensions.ts`、`decoration-extensions.ts`、`caret-marks.ts`、`shortcuts.ts`、`packages/docx-engine/src/parse-styles.ts`、`types.ts`。

## 3. 已找到、尚未修正

**文件（Word）**

1. Option＋Backspace 刪字詞時，Word 會順手吃掉一個相鄰空格（智慧剪貼）。`K02`：`Hello, world|!` → Word `Hello,!`，GenOffice `Hello, !`。`K13`：`…混合 t|ext` → Word `…混合ext`，GenOffice `…混合 ext`。
2. `K14`（選取 `world` 後按 Option＋Backspace）：Word 結果是 `Helloworld!`——沒刪選取，反而刪掉選取前的 `, `；GenOffice 只刪選取。很像選取先被收起再刪前一個單位，可能是送鍵的假象，需用全螢幕驗證後再決定。
3. `K09`／`K10`（⌘＋Backspace／Delete）：Word 只刪了一個字元，疑似修飾鍵沒有送進去，目前不能當標準答案。GenOffice 的 ⌘＋Backspace 會刪到行首，⌘＋Delete 沒有反應。
4. `S13`：⌘X 跨段剪下要套用和 Backspace 相同的邊界規則，待驗證。
5. `B01`：靠右對齊段落開頭按 Backspace，Word 沒有任何反應。
6. 自動校正（句首大寫等）：Word 有，GenOffice 沒有。
7. `T11`：表格第一列的儲存格內（第 1 字後）按 ↑，Word 落在上一段的第 1 字後（推測是保持水平位置），GenOffice 落在上一段結尾。

**試算表公式（以 Excel 為準；447 案、832 格，37 案不同）**

可修的差異（29 案）：

1. **正負號：** `=1+--1` GenOffice 得 0（應為 2）；`=-"3"`、`--"1e3"`、`ISNUMBER(--"1e3")`、`--" 12 "` 得 #NAME?（應為 −3、1000、TRUE、12）；`--"0x10"` 應為 #VALUE!。手動探測時 `=2*--3`、`--TRUE`、`--"3"` 也出過問題，這三個不在 447 案語料內，修好後補進語料。
2. **數字轉文字：** `1.0&""` 得 "1.0"（應為 "1"）；`1/3&""` 是 16 位（應為 15 位有效數字 "0.333333333333333"）；`123456789012345678&""` 應為 "123456789012345000"；`1E-7&""` 應為 "0.0000001"。
3. **日期時間：** `EDATE(2024/1/31,1)` 與 `EDATE(2024/3/31,-1)` 得 45353（3/2），應為 45351（2/29）；`SECOND(0.7534)` 得 53，應為 54（先四捨五入到秒；`HOUR(0.75)`、`MINUTE(0.7534)` 在語料內是一致的）；`DATE(1900,2,29)` 得 61，應為 60；`DAY(60)` 得 28，應為 29（Excel 把 1900 當閏年）；`DATE(1899,12,31)` 得 0，Excel 把小於 1900 的年份加上 1900（得 693962）。
4. **空字串與空白：** 儲存格內容為 `=""` 時 `ISBLANK` 應為 FALSE（GenOffice TRUE）；`REPT("a",0)`、`T(1)` 應回空字串，GenOffice 回空白。
5. **文字函式：** `SEARCH("b*c",…)`、`SEARCH("~*",…)` 要支援萬用字元 `*`、`?`、`~`；`TEXT(TRUE,"0")` 應為 "TRUE"（GenOffice 是 "true"）；`DOLLAR(1234.567,2)` 結果尾端多一個空格。
6. **邏輯與查找：** `IF("a","y","n")` 應為 #VALUE!（GenOffice 當成真）；`MATCH(35,遞增資料,-1)` 應為 #N/A（GenOffice 回 4）；交集運算子 `SUM(M1:M5 L1:L5)` 沒有交集應為 #NULL!（GenOffice #REF!）。
7. **數學：** `POWER(-8,1/3)` Excel 得 −2，GenOffice #NUM!。
8. **動態陣列：** `TAKE(M1:M5,2)` 得到 50、40（應為 10、20）；`DROP(M1:M5,3)` 得到 20、10（應為 40、50）——順序顛倒。

不用修的（8 案，Excel 端或語系造成，不能當標準答案）：

- `1E+308*10`、`1E+400`、`-1E+400`：Excel 端讀回空字串（Excel 拒收該輸入或 AppleScript 讀不出錯誤值）。
- `LEN("😀a")`：Excel 讀到 2，疑似經 AppleScript 傳入時表情符號被轉壞；Excel 本身以 UTF-16 計算應為 3，GenOffice 的 3 是對的。
- `LENB("你a")`、`CODE("你")`、`ASC("ＡＢＣ")`：行為隨 Excel 的語系（是否雙位元組）而變，這台 Excel 不是雙位元組語系。
- `DATEVALUE("3/5/2024")`：Excel 這台語系的日期格式不吃 月/日/年。

其餘 410 案（含財務函式、SORT／FILTER／UNIQUE／SEQUENCE 等動態陣列）全部一致。另外：新活頁簿只有 1000 列（Excel 為 1,048,576 列），要確認使用者能不能輸入到更下面的列。

## 4. 還沒測的項目

**Word／文件**
- 需要使用者核准全螢幕控制：格式快速鍵（⌘B／I／U、置中靠左靠右、行距、標題 ⌥⌘1～3、分頁符號、不斷行空格、⇧F3 大小寫、上下標、字級）、打字式自動格式（`1. `、`- `、`a) `）、智慧引號、復原分組（連續打字要幾次復原）、剪貼簿與貼上格式、尋找取代。
- 背景可測：雙擊／三擊選取、Shift＋點擊延伸、換行字元（Shift＋Enter）旁的刪除、圖片、頁首頁尾、頁面設定、字型。
- 未寫成固定測試：上面「已對齊」的案例還沒做成 `e2e/word-parity.spec.ts`（用 Word 實測值當標準答案）；`word-keys.ts` 的字詞單位函式（`wordUnits`／`wordRight`／`wordLeft`）還沒有單元測試。

**Excel／試算表**：複製貼上的參照位移、填滿序列、排序與篩選、插入刪除列欄後的公式更新、數字格式顯示（X-12）、編輯鍵、凍結窗格、合併儲存格、工作表操作、XLSX 往返。

**PowerPoint／簡報**：全部尚未開始（物件預設樣式與位置、複製貼上位移、方向鍵微調距離、文字方塊自動調整、圖層、對齊分配、備忘稿、放映）。

**PDF／Markdown／HTML**：全部尚未開始（清單在 `.agent/skills/genoffice-qa/testing/pdf-markdown-html.md`：F-01～F-07、M-01～M-04、H-01～H-04）。

## 5. 接續時的順序建議

1. 重跑 `npm run build:all`、文件單元測試、完整 e2e，確認今天的修改沒有弄壞既有項目。
2. 把「已對齊」的 Word 案例做成 `e2e/word-parity.spec.ts`，補 `word-keys.ts` 單元測試。
3. 修試算表公式差異（先處理第 3 節的 1～3 類），每項補單元測試，並把 `=2*--3` 等案例補進語料。
4. 向使用者要全螢幕控制，一次跑完文件的快速鍵與打字類測試。
5. 依序做 Excel UI、PowerPoint、PDF／Markdown／HTML。

## 6. 環境提醒

- Homebrew 在背景升級期間，`node` 短暫無法啟動（缺 `libsimdjson.33.dylib`），升級完成後恢復（node v26.10.0）。
- Excel 在背景（隱藏）時處理 AppleScript 很慢，447 案跑了約 15 分鐘。
- 依使用者規則：用電池且電量低於 50% 時暫停所有開發與測試。

## 7. 2026-10-08 接續記錄（電池 49%、用電池，依規則停手）

- 核對並更正了上面的數字與描述：重跑 `compare_fx.py` 確認 447 案中 37 案不同，其中 29 個是真差異、8 個是 Excel 端或語系造成。
- 新增不用開 Electron 的公式測試台 `apps/sheets/tests/formula-engine.ts`：`createFormulaEngine()` → `set({ A1: '=…', M1: 10 })` → `value('A1')`，一批約 3 秒。目前**還沒有任何修正**，也沒有用它寫正式測試（探索用的暫存檔已刪除）。
- 讀 Univer 0.25.1 引擎（`node_modules/@univerjs/engine-formula/lib/es/index.js`）得到的線索，**都還沒用測試台驗證，修之前先驗證**：
  - 一元負號 `-x` 由 `PrefixNode` 呼叫 `MINUS(0, x)`（約第 10099 行）；`PrefixNodeFactory.checkAndCreateNodeType` 對以引號包住的 token 直接略過（約第 10143 行），`-"3"`、`--"1e3"` 可能因此變成 #NAME?。
  - `&` 與 `CONCATENATE` 共用 `Concatenate` 執行器（約第 37781 行），數字直接用 JS 內建轉字串（17 位、`1e-7`），要改成 Excel 的 15 位有效數字規則；`numfmt-fix.ts` 已實作 General 顯示規則，可參考重用。
  - 字面值 `1.0` 走 `ValueObjectFactory.create`（約第 8112 行）：`Number("1.0").toString()` 不等於 `"1.0"` 時會變成字串物件，所以 `1.0&""` 得 "1.0"。**若屬實，`ISNUMBER(1.0)`、`3>=2.0` 這類比較也會出錯，影響會比其他項目大**，是第一個要用測試台確認的點。
  - `TAKE`（約第 29970 行）的 `_getResultArray` 看起來沒有問題，順序顛倒可能來自 `ArrayValueObject.slice` 或語料的資料排列，需要實測。
- 建議做法：用測試台為第 3 節那 29 案建立 `excel-formula-parity.test.ts`（期望值取自 `test/_probe/fx/excel-out.txt`），逐項用「包裝原函式執行器」的方式修（同 `cell-function.ts`、`ifs-empty-set.ts`，在 `App.tsx` 安裝）；全部做完後用 `fx.probe.spec.ts` 在真 app 重跑 447 案，目標是 37 案降到只剩 8 個 Excel 端假差異。
