# Docs 非文字物件：選取／操控／格式設定 與 Word for Mac 差異評估

- 觀察日期：2026-10-20
- 基準：Microsoft Word for Mac 16.113.3（Microsoft 365，繁體中文介面）
- 受測：GenOffice Docs，分支 `qa/local-fixes`，`apps/shell/out` 為 2026-10-16 17:05 建置（Docs 原始碼無晚於建置的修改）
- 性質：**只評估，未修改任何產品程式碼**。探針放在 `e2e/_probe/`（未 commit）。

## 0. 方法與限制（請先讀）

1. **GenOffice 側**：以 Playwright Electron 探針驅動真實滑鼠／鍵盤事件，讀取 ProseMirror 選取狀態、`document.activeElement` 與 DOM。視窗 1280x760、AI 側欄開啟（部分補測用 2400x960）。截圖在 `/private/tmp/claude-501/obj-docs/`，原始結果為同資料夾 `obj*.jsonl`。
2. **Word 側限制**：computer use 的整螢幕控制連續被中斷（"user interrupt"），背景模式又無法點擊 Word 畫布，因此**無法在 Word 畫布上做滑鼠並排操作**。Word 基準的來源分三級，報告內以標籤註明：
   - 【Word 實測】本機 Word 16.113.3 以 AppleScript 驗證（例如跨格範圍 `selection` 回報 cells=4 rows=2 cols=2）。
   - 【Word 辭典】Word 的 AppleScript 辭典明文（`insert rows`：「插入列數等於所選列數」；`insert columns` 同理）。
   - 【Word 既有行為】Word for Mac 長期穩定的行為，**未並排實測**，交給後續驗證（Codex）在 Word 上確認。
3. 同時有其他 agent 在跑 Slides／Sheets 的 Electron；本評估一次只開一個 Electron 並於結束關閉。
4. 未使用私人檔案；探針文件皆為全新空白內容。

## 1. 結論摘要

- 差異總數 **27 項：P0 = 6、P1 = 6、P2 = 15**（另有「一致」項目與「未驗證」項目，見第 5、6 節）。
- 使用者在 Slides 發現的兩類問題在 Docs 的狀況：
  - (1)「改格式時選取被取消」：**本文文字不會**（功能區 17 種操作選取全保留、焦點留在文件），但有三個例外：①**字型／段落對話框關閉後焦點掉到 BODY，鍵盤輸入進不了文件（D-01）**；②**文字方塊／形狀內選字後按功能區，編輯被結束、選取消失、格式沒套用（D-05）**；③多格選取時段落格式只套到一格（D-02）。
  - (2)「表格選到文字、無法跨格範圍選取」：**Docs 有範圍選取**（拖曳、Shift+點擊、Shift+方向鍵都會產生 CellSelection，反白、底紋／粗體／字級／項目符號／刪除／合併／刪列欄都對整個範圍生效），這點比 Slides 好；但缺整列／整欄點選、右鍵會破壞範圍（D-03）、插入列欄不依選取數量、整表選取沒有反白且把手位置錯位。
- 前五項最影響使用者（細節與重現見第 7 節）：D-01、D-04、D-02、D-03、D-06（D-05 同屬文字方塊問題，與 D-06 並列）。

## 2. 優先序定義

- P0：日常編輯會卡住、破壞內容或使輸入進不了文件。
- P1：常用功能與 Word 行為明顯不同。
- P2：少用或有替代路徑。

## 3. 差異清單（逐項：Word 預期 → GenOffice 實測）

### 3.1 文字格式與焦點

| ID | 優先 | 差異 | Word 預期 | GenOffice 實測 |
|---|---|---|---|---|
| D-01 | P0 | 字型／段落對話框關閉後焦點掉到 BODY | 【Word 既有行為】對話框確定或取消後選取保留、焦點回文件，直接打字會取代選取 | 選字 → 右鍵「Font…」或「Paragraph…」→ OK 或 Esc：選取範圍保留（7-20），`activeElement` 為 BODY，之後打字「Z」**沒有進文件**（3/3 重現，`obj12.jsonl` a、a2、a3） |

### 3.2 表格

| ID | 優先 | 差異 | Word 預期 | GenOffice 實測 |
|---|---|---|---|---|
| D-02 | P0 | 多格選取時段落層級格式只套到最後一格 | 【Word 既有行為】2x2 選取後按「置中」，四格都置中；靠右、行距同理 | 2x2 拖曳選取後 Home「置中」、「靠右」、「行距」：只有右下格（最後一格）改變（`obj10.jsonl` range center/right/line spacing）。**同一範圍上粗體、字級、底紋、項目符號是 4/4 格都生效**，所以問題只在段落對齊／行距路徑 |
| D-03 | P0 | 右鍵點在儲存格範圍選取內，範圍被折疊成游標 | 【Word 既有行為】右鍵不破壞選取；選單的合併儲存格、插入、刪除作用於整個範圍 | 拖曳選取 r0c0–r1c1（4 格）→ 在範圍內右鍵：選取變 TextSelection、`.selectedCell` 由 4 變 0，選單仍列出 Merge Cells／Insert›／Delete›，但已無範圍可作用（`obj4.jsonl` 5b）。右鍵在已選取的**文字**上則保留選取（正常，`obj11.jsonl` a） |
| D-04 | P0 | 儲存格內「插入圖片」會把表格拆成兩張 | 【Word 既有行為】圖片以行內物件插入游標所在儲存格 | 游標在表格內某格 → Insert > Picture：頂層結構變為 `docTable, docProtected(image), docTable`，圖片夾在兩張表之間，原表被切斷（`obj4.jsonl` 8a、截圖 `t4-pic-in-cell.png`）。圖片是區塊節點，沒有走 `docInlineImage` 行內路徑 |
| D-07 | P1 | 整表選取（移動把手）：把手位置錯位、且選取後沒有反白 | 【Word 既有行為】把手在表格左上外側；整表選取時所有儲存格反白 | 懸停表格後把手出現在表格**下方偏右**（表格左上 (511,241)，把手 (909,339)，偏移約等於頁面容器原點 (415,116)；`obj6.jsonl` 三個懸停點位置相同）。程式碼推測（未驗證）：`table-handle.ts` 的 `container` 在外掛建立時取得，之後被搬移，座標以 viewport 當容器偏移。點把手後為 NodeSelection(docTable)，**畫面沒有任何反白**（`t4-handle-selected.png`），僅功能區顯示 Table 分頁。整表套粗體／置中仍對全部 12 格生效 |
| D-08 | P1 | 無法點列左緣選整列、點欄上緣選整欄 | 【Word 既有行為】滑鼠移到列左外側（白箭頭）點一下選整列，可上下拖曳選多列；欄上緣（黑箭頭）選整欄 | 點列左緣 10px 外：游標落在 r0c0（TextSelection 12）；拖曳左緣：同；點欄上緣 5px 外：游標跑到表格後的段落（`obj3.jsonl` E、E2、F）。右鍵 Select› 子選單可替代，但不是 Word 的主要路徑（子選單內容未逐項驗證） |
| D-09 | P1 | 插入列／欄不依選取數量 | 【Word 辭典】插入數量等於選取的列／欄數 | 選兩列按 Insert Above：列數 4→5（只插一列）；選兩欄按 Insert Right：欄數 4→5（`obj3.jsonl` J1、J2） |
| D-15 | P2 | 欄線拖曳行為與列高拖曳 | 【Word 既有行為】拖欄線只改相鄰兩欄；拖列下緣可調列高；表格右下有縮放控制點 | 拖第 1 欄右邊界 +80px：欄寬 150,150,150,150 → 199,134,134,134（其他欄被均分重配，且位移 49px 不跟手）；拖列下緣不調高，反而選到儲存格；無右下角表格縮放把手（`obj5.jsonl` 2、3；`obj4.jsonl` 4） |
| D-16 | P2 | 表格右鍵選單缺項 | 【Word 既有行為】含表格內容、框線與底紋、平均分配、儲存格對齊、文字方向等 | 只有 Cut/Copy/Paste/Paste as Plain Text/Font…/Paragraph…/Insert›/Delete›/Select›/Merge/Split/Synonyms/Translate›/Hyperlink…/New Comment（`obj4.jsonl` 5） |
| D-17 | P2 | Table Layout 缺功能 | 【Word 既有行為】平均分配列／欄、排序、公式、轉換成文字、檢視格線、繪製表格／清除、文字方向、儲存格邊界 | 分頁只有刪除表格、插入四向、合併／分割、刪列／欄、對齊（上中下、左中右）、AutoFit（僅 3 項）、重複標題列、表格內容；Table Properties 對話框分頁為 Table/Row/Cell/Alt Text，無獨立 Column 分頁 |
| D-18 | P2 | 儲存格內三擊與格尾 Shift+→ | 【Word 既有行為】三擊選段落（不選整格）；格尾第一次 Shift+→ 選取格尾標記，第二次才延伸到下一格 | 三擊選取整格（CellSelection 1 格）；格尾按一次 Shift+→ 就變成 2 格選取（`obj3.jsonl` H、`obj4.jsonl` 6） |
| D-14 | P2 | 整表選取後 Delete／Backspace 直接刪表 | 【Word 既有行為，待驗證】Delete 清除內容、表格留下 | 兩個鍵都整張表刪除（`obj4.jsonl` 1b、1c） |
| D-19 | P2 | 範圍選取後輸入與 Backspace | 【待驗證】Word 的輸入／Backspace 對選取儲存格的處理（內容取代位置、是否出現刪除儲存格對話框） | 輸入「X」：四格內容清空，X 出現在最後一格；Backspace：清空內容、保持範圍（`obj3.jsonl` I6、I7）。Word 差異未實測 |

### 3.3 圖片

| ID | 優先 | 差異 | Word 預期 | GenOffice 實測 |
|---|---|---|---|---|
| D-10 | P1 | 圖片是區塊物件，不是行內圖片 | 【Word 既有行為】預設為「與文字排列」的行內圖片，是段落中的一個字元；圖片左右可同行打字、可用 Shift+方向鍵選取 | 游標在「abc\|def」中間插入圖片：段落被拆成 `Paragraph(abc), P:image, Paragraph(def)`（`obj7.jsonl` 1、`i7-mid.png`）。圖片獨佔一個區塊，無法與文字同行（僅 docx 匯入的儲存格內圖片走 `docInlineImage`） |
| D-19b | P2 | 拖曳行內圖片的結果 | 【Word 既有行為】拖曳行內圖片是在文字間移動，仍為行內 | 拖曳 +121/+13 px 後變為浮動 `imageWrap=square-left`，帶位移（`obj7.jsonl` 6）。文繞圖選項只有 Inline、Square 左／右、Top and Bottom、Behind、In Front，沒有 Tight／Through |
| D-20 | P2 | 無多選、群組、選取窗格；方向鍵不微調浮動物件 | 【Word 既有行為】Shift／Cmd＋點選多個浮動物件、群組、方向鍵微調 | Shift＋點兩張圖片仍是單一 NodeSelection；形狀同（`obj7.jsonl` 7、`obj8.jsonl` 5）；方向鍵移動游標而非物件；Layout > Arrange 只有 Position／Bring Forward／Send Backward／Align／Wrap Text |
| D-21 | P2 | 雙擊與右鍵選單 | 【Word 既有行為】雙擊開啟 Picture Format；右鍵含 Format Picture、Edit Alt Text、Size and Position、Crop、Caption | 雙擊開啟全螢幕影像檢視器（`gs-imgview-mask`）；右鍵含 View Image／Save Image As…，另有與圖片無關的 Font…／Paragraph…／Synonyms／Translate；沒有 Alt text 入口（Picture Format 分頁也沒有） |
| D-22 | P2 | Picture Format 分頁內容 | 【Word 既有行為】圖片樣式、框線、效果、校正、色彩、壓縮、替代文字、位置、文繞圖、大小與裁剪 | 只有 Remove Background、Crop、Replace Picture、Rotate／Flip、Wrap 按鈕、Align L/C/R、兩個尺寸輸入、Reset Size |

### 3.4 文字方塊、形狀、WordArt、圖表、方程式等

| ID | 優先 | 差異 | Word 預期 | GenOffice 實測 |
|---|---|---|---|---|
| D-05 | P0 | 文字方塊／形狀內：功能區格式化會結束編輯、丟失選取、格式未套用 | 【Word 既有行為】框內選字後按 Home 的粗體等，格式套在選取文字、選取保留、仍在編輯 | 雙擊進入框內編輯並選取「hello world」→ Home 粗體：`activeElement` 變 BODY、DOM 選取清空、`.doc-content-editing` 數量 1→0、粗體未套用（`obj9.jsonl` C）。根因線索：`extensions.ts` 的 `onDocumentMouseDown` 對框外 mousedown（含功能區）一律 `setEditing(false)` |
| D-06 | P0 | 文字方塊／形狀要「雙擊＋再單擊」才有游標；雙擊後直接打字會刪除整個物件 | 【Word 既有行為】單擊框內即出現游標，直接打字 | 單擊框內只是選取物件（NodeSelection）；此時打字「Hello」→ 物件被取代，文件只剩一個「Hello」段落（`obj8.jsonl` 2）；雙擊後直接打字同樣取代（`obj9.jsonl` A、`obj8.jsonl` 3）；雙擊後**再單擊框內**才能輸入（`obj9.jsonl` B）。Esc 可回到物件選取（與 Word 一致） |
| D-11 | P1 | 縮放與旋轉控制點 | 【Word 既有行為】8 個縮放點＋旋轉把手；圖片、形狀皆可自由旋轉 | 圖片：8 個縮放點、無旋轉把手（僅 Rotate Right/Left 90°、翻轉按鈕）。文字方塊／形狀／圖表／WordArt：**只有右下一個縮放點**、無旋轉（`obj7.jsonl` 2、`obj8.jsonl` 1、4）。圖片縮放行為一致：角點等比、邊點變形、Shift 等比 |
| D-12 | P1 | Shape Format 情境分頁過度精簡 | 【Word 既有行為】形狀樣式、填滿、外框、效果、文字方向／對齊／邊界、尺寸、旋轉、文繞圖、位置 | 只有 Shape Fill、Shape Outline、B/I/U、字色、四個對齊（`obj8.jsonl` 1、6）；沒有尺寸輸入、旋轉、效果、文字方塊選項。繪製形狀為 Word 式十字游標拖曳（與 Word 一致） |
| D-23 | P2 | 移動方式與右鍵 | 【Word 既有行為】拖曳框線移動；右鍵有 Format Shape 等 | 只能拖左上「↕」把手移動（把手拖曳有效，+146/-84）；從框內主體拖曳不移動（`obj9.jsonl` E）；右鍵形狀無任何選單（`obj8.jsonl` 7）；文字方塊從儲存格內插入時放到表格之後 |
| D-24 | P2 | 圖表 | 【Word 既有行為】Chart Design／Format 分頁、編輯資料 | 單擊選取（NodeSelection）、單一右下縮放點；無圖表情境分頁；雙擊不開資料視窗（資料格在選取時顯示於圖表下方）（`obj8.jsonl` 11） |
| D-25 | P2 | 方程式 | 【Word 既有行為】行內方程式，雙擊／點選進入方程式工具 | 從 Insert > Equation 圖庫選取後為**區塊**物件（`P:passthrough`）；單擊選取 NodeSelection 並出現「Equation」情境分頁；雙擊沒有編輯視窗（`obj13.jsonl` c） |
| D-26 | P2 | 缺少：SmartArt、內容控制項／表單欄位（無 Developer 分頁）、圖示／3D 模型 | 【Word 既有行為】皆有 | Insert／ribbon 全部按鈕清單中沒有這些功能（`explore.json`、`obj12.jsonl` 0） |
| D-27 | P2 | 頁首／頁尾經由浮動 popover 的純文字 textarea 編輯 | 【Word 既有行為】頁內就地編輯、可含豐富格式與物件 | 從 Insert > Header 進入：浮動 popover 含 textarea 與 Picture／Page／Total pages／Date 等按鈕和 OK／Cancel；Esc 後焦點為 BODY（`obj12.jsonl` g、`m12-hdr.png`）。**直接雙擊頁面頁首區的就地編輯未測** |

（編號 D-13 未使用；D-19b 為 D-19 之外另列的圖片項目。）

## 4. 一致項目（已實測，與 Word 相同或可接受）

- **文字格式保留選取與焦點**（`obj10.jsonl`，選取 7-20）：粗體、底線、字型（輸入框＋Enter／下拉）、字級（輸入框＋Enter／Tab／下拉）、字色色盤、螢光筆色盤、大小寫、置中、項目符號、行距選單、段落底紋、樣式 Heading 1、格式刷，共 17 項全部「選取不變、焦點在文件、之後打字會取代選取」。字級輸入框只點一下不改值時焦點留在輸入框（與 Word 相同）。字級輸入框聚焦時選取以灰底保留（`inactive-selection.ts`）。
- 在已選取文字上右鍵：選取保留，Font…／Paragraph… 對話框能開（焦點問題見 D-01）。
- **表格範圍選取**：拖曳跨格（含同列）、Shift＋點擊（2x2 選取 9 格正確）、Shift＋↓ 延伸成整格，反白為半透明藍。
- 範圍上：底紋、粗體、字級、項目符號（4/4 格）、Delete 清空內容並保持範圍、合併儲存格、刪除多列／多欄（刪 2 列、2 欄正確）、Tab／Shift＋Tab 移動、末格 Tab 新增列、複製貼上範圍（貼回同形狀）。
- 整表（把手）選取後套粗體、置中對全表有效；表格改字型等功能區操作後選取與焦點保留（`obj3.jsonl` L1）。
- 圖片：單擊選取、方向鍵由物件選取回到游標、選取時打字會取代圖片、Rotate 與對齊後仍保持選取、Crop 模式（Enter 確認／Esc 離開，選取保留）、複製貼上圖片、右鍵 Wrap Text› 子選單。
- 區段分隔符號：單擊選取（NodeSelection）、Delete 刪除；Columns 選單（One／Two／Three／Left／Right／More）；Design 分頁有 Page Color、Watermark、Page Borders。
- 超連結：單擊只放游標（Word for Mac 相同）；Cmd＋點擊的開啟邏輯在 `App.tsx:5882`（`window.open`），實際導向結果未觀察。

## 5. 尚未驗證項目

1. **Word 並排實測**：所有標【Word 既有行為】者都未在 Word 畫布上實測（見第 0 節限制），特別是 D-14、D-19（Delete／Backspace／輸入對儲存格範圍的行為）。
2. 欄位（PAGE／DATE 等）的選取與編輯：探針插入 Page number 後在頂層節點看不到欄位，樣式需用頁面實際渲染判讀，結果不可用。
3. 註腳／章末註：插入後找不到 `.doc-note-ref`，且焦點落在 AI 側欄的 textarea（`obj13.jsonl` f，疑似探針選擇器或焦點問題，待確認）。點選／雙擊參照標記、註腳區編輯未驗證。
4. 目錄（TOC）：點選／雙擊／更新表格分頁未驗證（探針選擇器逾時）。
5. 頁首頁尾：直接雙擊頁面頁首區的就地編輯、頁首內物件選取，未驗證。
6. 交互參照、書籤、分欄與頁面框線對話框的物件層級行為。
7. 圖表建立後的資料編輯完整流程；WordArt 單擊／雙擊文字編輯（與 D-06 同機制，僅確認雙擊後打字會取代，細節未逐一確認）。
8. 浮動圖片錨點標記、環繞文字拖曳定位精度、Format Picture 窗格（Docs 無此窗格）。
9. 跨頁表格列重複標題在畫面上的行為；表格巢狀。
10. 同一探針在較小視窗或功能區折疊狀態的表現。

## 6. 差異統計

| 優先 | 項目數 | 編號 |
|---|---|---|
| P0 | 6 | D-01、D-02、D-03、D-04、D-05、D-06 |
| P1 | 6 | D-07、D-08、D-09、D-10、D-11、D-12 |
| P2 | 15 | 表格：D-14、D-15、D-16、D-17、D-18、D-19；圖片：D-19b、D-20、D-21、D-22；物件：D-23、D-24、D-25、D-26、D-27 |
| 合計 | 27 | 其中 D-14、D-19 的 Word 行為為【待驗證】，列入差異但屬低信心 |

## 7. 建議修正順序

1. **D-01 焦點**：所有對話框（字型、段落、符號、超連結等）關閉後把焦點還給編輯器並還原選取。影響所有文字編輯，改動最小、收益最大。
2. **D-02 + D-03**：CellSelection 上的段落格式套用到所有格；右鍵在 CellSelection 內不要重設選取。兩者同屬「表格範圍選取」一致性，修完才算追平 Slides 使用者要求的「整列改底色」那類流程。
3. **D-04 + D-10**：圖片改為行內節點（至少在儲存格內與段落中插入時不拆段、不拆表）。需評估 docx 儲存模型（`docInlineImage` 已存在）。
4. **D-05 + D-06**：文字方塊／形狀：單擊框內即進入編輯並放游標；功能區 mousedown 不結束編輯；選取物件時打字不得刪除物件（至少先確認編輯意圖）。
5. **D-07、D-08、D-09**：表格把手定位與整表反白；列左緣／欄上緣選取；插入列／欄依選取數量。
6. **D-11、D-12**：形狀／文字方塊 8 點縮放與旋轉把手；Shape Format 補尺寸、旋轉、文繞圖。
7. P2 依使用頻率：右鍵選單補項（表格內容、圖片 Alt text）、Table Layout 缺項（平均分配）、欄線拖曳語意、浮動物件多選／微調、SmartArt 與內容控制項等新功能。

---

# 附錄：重現手冊（P0／P1，供 Codex 驗證）

通用前置：
1. `npm run build:all` 確認 `apps/shell/out` 為最新；本評估使用 2026-10-16 建置。
2. 啟動 GenOffice，從首頁快速卡片開新 Docs 空白文件。
3. Word 側：用 Word for Mac 16.113.3 新增空白文件，建立同樣內容。
4. 自動化重現：`GENOFFICE_E2E_NO_VIDEO=1 npx playwright test --config e2e/playwright.config.ts e2e/_probe/<探針>`（需前景；背景會有焦點誤差）。探針與結果檔對照見每項「證據」。共用工具在 `e2e/_probe/obj-common.ts`。
5. 在手動重現時，內容文字均可自訂，下列步驟只列必要動作。

## D-01 對話框關閉後焦點掉到 BODY（P0）

- 步驟：
  1. 輸入「Alpha bravo charlie delta」，選取「bravo charlie」。
  2. 在選取文字上按右鍵 → 選「Font…」→ 按 OK（或按 Esc）。
  3. 直接打字「Z」。
  4. 重複：右鍵「Paragraph…」→ OK。
- Word 預期：對話框關閉後選取仍在、焦點在文件，打字會取代選取。
- GenOffice 實際：選取範圍保持 7-20，但 `document.activeElement` 是 `BODY`，「Z」沒有進入文件，文字不變。
- 證據：`e2e/_probe/docs-objsel12.probe.spec.ts` 測項 a／a2／a3，`obj12.jsonl`。

## D-02 多格選取的段落格式只套一格（P0）

- 步驟：
  1. Insert > Table 建 4x4，每格填入文字。
  2. 滑鼠從 r0c0 拖到 r1c1（共 2x2，四格反白）。
  3. Home 按「置中」（再試「靠右」、「行距 2.0」）。
  4. 對照：同一選取按粗體、字級 20、Table Design 底紋。
- Word 預期：四格全部置中（靠右、行距亦同）。
- GenOffice 實際：只有右下格（r1c1）置中；粗體、字級、底紋、項目符號四格全部生效。
- 證據：`docs-objsel10.probe.spec.ts`（range center／right／line spacing／bullets）、`docs-objsel3.probe.spec.ts` I1–I3。

## D-03 右鍵在範圍選取內折疊選取（P0）

- 步驟：
  1. 4x4 表格，拖曳選 r0c0–r1c1。
  2. 在選取範圍內任一格按右鍵。
  3. 選單出現後觀察反白；點「Merge Cells」。
- Word 預期：四格反白保持，Merge Cells 合併這四格。
- GenOffice 實際：右鍵瞬間範圍折疊為游標（TextSelection，`.selectedCell` 由 4 變 0），選單仍有 Merge Cells，但已無範圍可合併。
- 證據：`docs-objsel4.probe.spec.ts` 5b，`obj4.jsonl`。

## D-04 儲存格內插入圖片拆表（P0）

- 步驟：
  1. 建 4x3 表格並填字；游標放在第 2 列第 1 格。
  2. Insert > Picture，選任一 PNG。
  3. 觀察表格。
- Word 預期：圖片以行內圖片出現在該儲存格內，表格仍為一張。
- GenOffice 實際：表格被拆為上下兩張，圖片以區塊夾在中間（頂層：`docTable, docProtected(image), docTable`）。
- 證據：`docs-objsel4.probe.spec.ts` 8a、截圖 `t4-pic-in-cell.png`。

## D-05 文字方塊內功能區格式化中斷編輯（P0）

- 步驟：
  1. Insert > Text Box；雙擊框、再單擊框內，輸入「hello world」。
  2. Shift+Home 選取整行文字。
  3. Home > Bold。
- Word 預期：選取文字變粗體，選取保留，游標仍在框內可繼續輸入。
- GenOffice 實際：`activeElement` 變 BODY，選取消失，編輯模式結束，沒有粗體。
- 證據：`docs-objsel9.probe.spec.ts` C。

## D-06 文字方塊進入編輯的方式與誤刪（P0）

- 步驟：
  1. Insert > Text Box（或 Shapes 畫矩形）。
  2. 單擊框內，直接打字「Hello」。
  3. 復原，改為雙擊框內後立刻打字「abc」。
  4. 再次復原，改為雙擊後再單擊框內，打字「xyz」。
- Word 預期：單擊框內就有游標，打的字進框內，物件不會消失。
- GenOffice 實際：步驟 2、3：物件被刪除，文件只剩「Hello」／「abc」一個段落；步驟 4 才進框（框內文字 xyz）。Esc 後回到物件選取（符合 Word）。
- 證據：`docs-objsel8.probe.spec.ts` 2、3；`docs-objsel9.probe.spec.ts` A、B。

## D-07 整表把手錯位且整表選取無反白（P1）

- 步驟：
  1. 建 3x4 表格；滑鼠移到表格上，觀察把手位置。
  2. 點把手；觀察表格有沒有反白。
  3. 按 Delete。
- Word 預期：把手在表格左上外側；整表選取時所有儲存格反白；Delete 清除內容（表格保留，【待驗證】）。
- GenOffice 實際：把手出現在表格下方偏右（以表格左上 (511,241) 為例，把手在 (909,339)，三個懸停點位置相同）；點擊後僅功能區出現 Table 分頁、表格沒有反白；Delete／Backspace 整張刪除。
- 證據：`docs-objsel6.probe.spec.ts`（style 位置 top 223／left 493 與容器原點 (415,116) 相加得 (909,339)）、`docs-objsel4.probe.spec.ts` 1／1b／1c、截圖 `t4-handle-hover.png`、`t4-handle-selected.png`。

## D-08 整列／整欄點選（P1）

- 步驟：
  1. 4x4 表格。滑鼠移到第 2 列左緣外 10px 點一下；再試由第 2 列拖到第 3 列左緣外。
  2. 滑鼠移到第 2 欄上緣外 5px 點一下。
- Word 預期：第 2 列整列反白（拖曳為兩列）；第 2 欄整欄反白。
- GenOffice 實際：游標放到 r0c0（列）；欄上緣點擊讓游標跑到表格後的段落；均無整列／整欄反白。
- 證據：`docs-objsel3.probe.spec.ts` E、E2、F、截圖 `t-E-rowleft.png`。

## D-09 插入列／欄不依選取數量（P1）

- 步驟：
  1. 4x4 表格，選取前兩列任意兩個儲存格（跨列拖曳）。
  2. Table Layout > Insert Above；再另選兩欄，按 Insert Right。
- Word 預期：【Word 辭典】插入與選取數量相同（2 列、2 欄）。
- GenOffice 實際：列 4→5、欄 4→5，只各插入 1。
- 證據：`docs-objsel3.probe.spec.ts` J1、J2。

## D-10 圖片為區塊物件（P1）

- 步驟：
  1. 輸入「abcdef」，游標移到 abc 與 def 之間。
  2. Insert > Picture。
- Word 預期：圖片行內出現在段落 abc 與 def 之間，三者在同一段落。
- GenOffice 實際：段落被拆成三個區塊：段落(abc)、圖片、段落(def)。
- 證據：`docs-objsel7.probe.spec.ts` 1、截圖 `i7-mid.png`。

## D-11 縮放與旋轉控制點（P1）

- 步驟：
  1. 插入圖片、矩形形狀、文字方塊、圖表各一，分別選取。
  2. 數控制點並嘗試旋轉；圖片再依序拖右下角、右邊中點、Shift＋左上角。
- Word 預期：每個物件 8 個縮放控制點加一個旋轉把手。
- GenOffice 實際：圖片 8 點但無旋轉把手（縮放比例行為與 Word 一致）；形狀／文字方塊／圖表只有右下 1 點、無旋轉。
- 證據：`docs-objsel7.probe.spec.ts` 2、4；`docs-objsel8.probe.spec.ts` 1、4。

## D-12 Shape Format 情境分頁（P1）

- 步驟：選取矩形形狀 → 看 Shape Format 分頁內容。
- Word 預期：形狀樣式、填滿、外框、效果、文字、排列、大小。
- GenOffice 實際：只有 Shape Fill、Shape Outline、B/I/U、字色、四個對齊。
- 證據：`docs-objsel8.probe.spec.ts` 1、6，`obj8.jsonl`。

---

*附：本報告引用的探針檔 `e2e/_probe/docs-objsel{0,1,2,3,4,5,6,7,8,9,10,11,12,13}.probe.spec.ts` 與 `obj-common.ts` 均未 commit；Word 測試文件已關閉不存檔。*
