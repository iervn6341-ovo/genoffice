# Sheets：儲存格範圍以外元件的選取／操控／格式行為 — 與 Excel for Mac 的差異評估

- 觀察日期：2026-10-20
- 範圍：只評估與整理，未修改任何產品程式碼、未 commit、未動 `.claude-flow/`、`test/`、PDF、pptx、`setup-genoffice-qa-skill.sh`；未開啟私人活頁簿。
- GenOffice 版本：工作樹 `qa/local-fixes`（`apps/shell/out` 現有建置，未重建）。
- Excel 基準：本機 Microsoft Excel for Mac 16.x（Microsoft 365，深色主題）。**版本號未讀取**。

## 方法與限制（先看這段）

1. GenOffice 端：用 Playwright Electron 啟動 shell（背景模式、獨立 userData），再用一個暫時的 REPL（在 scratchpad，未放入 repo）對編輯器頁面送真實滑鼠／鍵盤事件、讀 DOM 與 Univer API 的選取狀態。所有「GenOffice 行為」欄皆為實測。截圖放在 `/private/tmp/claude-501/obj-sheets/`（g01～g40）。
2. **Excel 端並排實測未能完成**：Excel 的格線是畫布，背景模式（app_* 工具）無法點擊／輸入；改用全螢幕控制時每次都被「user interrupt」中止（共 5 次嘗試）。只成功：開啟 Excel、建立空白活頁簿 Book1（內無資料，仍開著未存檔）、讀取功能區的輔助功能樹。因此下表「Excel 基準」欄是依 Microsoft 官方行為與既有 Office-Mac 稽核累積的已知行為，**標示為「已知行為，未並排實測」**。凡是我對 Excel 行為不確定的項目另行註明。建議之後在使用者可操作 Excel 時補一輪並排驗證（見文末「未驗證項目」）。
3. 圖表、樞紐分析表與表格均為 GenOffice 自訂的 DOM 疊層（`WorkbookVisuals.tsx`），不是 Univer 原生物件；格線選取則是 Univer 畫布。兩者行為分界不同，差異也集中在疊層物件。
4. 測試解析度 1280x800 CSS px、縮放 100%。

## 總覽

差異共 **27 項**：P0 = 2、P1 = 14、P2 = 11。另列出一致項目供對照。

優先序定義：P0 日常編輯會卡住；P1 常用；P2 少用。

## 一致（不需處理）

| 項目 | 實測結果 |
|---|---|
| 拖曳選範圍、Shift+點擊延伸、Shift+方向鍵延伸 | A2:C5、C5:D6 等皆正確 |
| 欄／列標頭點擊、Shift+標頭連選、標頭拖曳、左上角全選 | B:D、3:5 等正確；角落全選＝整張工作表 |
| ⌘A 兩段式 | 第一次選連續資料區 A1:D6，第二次全選 |
| 表格內 ⌘A | 資料 → 含標題 → 全表，三段式與 Excel 相同 |
| Ctrl+點擊多重選取 | 可加入（但 Mac Excel 的 Ctrl+點擊是右鍵，見 B1） |
| 名稱方塊輸入單一範圍／整欄／整列 | B2:C4、A:A、2:3 正確 |
| 拖曳邊框移動、填滿控點填序列 | 3,5 → 3,5,7,9 正確（Ask AI 小膠囊遮擋見 B3） |
| 欄寬拖曳、多欄同拖、雙擊自動調整、列高拖曳 | 與 Excel 一致 |
| 圖表／圖形／圖片：點選出現 8 個控制點、拖曳移動、控制點縮放、Esc 取消、點儲存格取消選取、Delete 刪除、功能區 Undo 可復原 | 一致 |
| 圖形隨儲存格移動並調整大小（預設） | 欄寬變動時圖形跟著變，與 Excel 預設一致 |
| 表格：篩選按鈕、Table Design 情境分頁、插入／刪除表格欄 | 有 |
| 資料驗證下拉箭頭 | 有 |
| 註解標記（儲存格角落小三角） | 有 |
| 圖表情境分頁 Chart Design | 選取圖表時出現（不自動切換，與 Excel 相同） |

## 差異清單

### P0

**A1　功能區按鈕點擊後鍵盤焦點留在按鈕上（格線收不到鍵盤）**
- GenOffice：選 A2:C5，點 Home 的粗體或填色（或其下拉箭頭）後，`document.activeElement` 是該 BUTTON，選取範圍保留；此時按 ↓ 不會移動儲存格、直接打字不會輸入（打 `Z` 後 A2 仍是 `Apple`）、Esc 也不作用在格線。要先用滑鼠點回格線才行。
- Excel（已知行為，未並排實測）：功能區按鈕不搶焦點，點完格式後直接打字／按方向鍵仍作用在格線。
- 影響：改完格式就要再點一次儲存格才能繼續鍵盤操作，日常編輯高頻觸發；與使用者在 Slides 回報的「改格式後選取被取消」同源。
- 重現：Home → 選 A2:C5 → 點 B → 按 ↓ 或輸入文字。
- 證據：g01-sel.png、g02-fillpopup.png。

**A2　儲存格內編輯時用功能區改格式，文字選取與焦點消失**
- GenOffice：雙擊 A2（內容 `Hello World`）進入編輯，⌘A 全選文字（藍色反白），點 Home 粗體：粗體有套上，但反白消失、游標回到開頭、焦點移到按鈕，編輯器仍開著；接著輸入文字沒有反應。
- Excel（已知行為）：選取保留，可連續疊加格式，輸入會取代選取文字。
- 影響：儲存格內部分文字的格式設定（粗體、顏色、字級）幾乎無法連續操作。這就是使用者在 Slides 看到的同一類問題。
- 重現：雙擊含文字的儲存格 → ⌘A → 點粗體 → 輸入字。
- 證據：g06-a-selectall.png、g06-b-afterbold.png、g06-c-aftertype.png。

### P1

**B1　⌘+點擊無法加選（Mac 慣用），標頭 ⌘+點擊變成範圍延伸**
- GenOffice：A2 → ⌘+點 C5 → 選取只剩 C5。⌘+點欄標頭 B 再 ⌘+點 D → 得到 B:D（等於 Shift 連選）。只有 Ctrl+點擊能加選（A2:C5 + C5；欄標頭 B:B + D:D）。
- Excel（已知行為）：⌘+點擊加選非連續儲存格／欄／列；Mac 上 Ctrl+點擊是右鍵選單。
- 影響：Mac 使用者無法用慣用手勢多重選取；且 Ctrl+點擊在 Mac 會觸發系統右鍵，實務上等於沒有多重選取。
- 重現：點 A2 → ⌘+點 C5，觀察名稱方塊。

**B3　「Ask AI」浮動膠囊蓋住填滿控點**
- GenOffice：用滑鼠拖出範圍（例如 B2:B3）後，膠囊（`button.ai-ask-trigger`，約 82x27 px，位於選取範圍右下方）出現在填滿控點正下方；在控點中心點擊命中的是膠囊而不是畫布，只剩控點最上緣幾像素可抓。（用 API 設定選取則沒有膠囊，控點正常。）
- Excel：沒有此類浮層，填滿控點永遠可抓。
- 重現：滑鼠拖選 B2:B3 → 在右下角小方塊上按下拖曳。
- 證據：g21-sel.png。

**C1　圖表內部元素無法用真實點擊選取（標題、圖例、座標軸、數列、資料點）**
- GenOffice：點圖表上的標題、圖例、座標軸、長條，Format 窗格標頭永遠是 `Format — Chart Area`，沒有元素被選取。原因：`WorkbookVisuals.tsx` 外層容器在 pointerdown 時 `setPointerCapture`，之後的 click 被重新導向到容器（事件記錄：click 目標為 `div.shape-editable`），圖表內各元素的 onClick（`selectElement`、`selectCategoryAxis` 等）永遠收不到。以程式直接呼叫 `figcaption.click()` 則標題會被選取，標頭變成 `Format — Chart Title`，證明元素選取邏輯存在但被這個點擊重導擋住。
- Excel（已知行為）：單擊選圖表，再單擊選標題／圖例／座標軸／整組數列，再單擊選單一資料點；Format 窗格跟著切換。
- 影響：圖表的細部格式（標題、圖例、座標軸、數列顏色）無從開始。
- 重現：Insert → Column chart → 點圖表標題 → Chart Design → Format Pane，標頭仍是 Chart Area。
- 證據：g10-*.png、g13-formatpane.png；程式位置 `WorkbookVisuals.tsx` 約 1111–1136 行。

**C2　方向鍵不能微調物件位置（圖表、圖形、圖片）**
- GenOffice：選取圖表／圖形後按 →、↓，位置不變。
- Excel（已知行為）：方向鍵移動選取物件，Option+方向鍵更細微（⌘ 無此行為）。
- 重現：選圖形 → 按方向鍵。

**C3　縮放修飾鍵與圖片長寬比鎖定缺失**
- GenOffice：圖表以 Shift 拖曳 SE 控點，寬度 545 → 605、高度不變（沒有保持比例）；沒有 Option 從中心縮放。圖片（Icons 插入，139x120）拖曳角控點變成 199x130，比例走樣。
- Excel（已知行為）：圖片預設鎖定長寬比（角控點等比縮放）；Shift 角控點保持比例；Option 從中心縮放。
- 重現：Insert → Icons → 任選 → 拖角落控點。

**D1　繪製圖形後，格線選取被誤改、新圖形沒被選取**
- GenOffice：Shapes → 矩形 → 在格線拖出框，結束後名稱方塊顯示 `17:21`，A17:D21 整列被選取（拖曳手勢同時被格線當作列選取）；新圖形沒有控點，必須再點一次才選取；沒有 Shape Format 分頁。
- Excel（已知行為）：繪製完成後該圖形立即被選取（並顯示 Shape Format 分頁），格線選取不變。
- 重現：Insert → Shapes → 矩形 → 在 A18 附近拖出矩形。
- 證據：g17-shape.png、g18-shapesel.png。

**D2　圖形／圖片／圖表不能多選，也沒有框選與選取窗格**
- GenOffice：Shift+點擊另一物件時，原選取被取代（圖形取消、圖表被選）；沒有框選手勢；`WorkbookVisuals.tsx` 中沒有 shift／meta 加選邏輯；無 Selection Pane。
- Excel（已知行為）：Shift／⌘ 點擊加選；Home → Find & Select → Select Objects 框選；Selection Pane 可逐一選取與隱藏。
- 影響：無法一次移動、刪除、對齊多個物件。

**D3　圖形／圖片缺少旋轉、裁剪、黃色調整控點，也沒有 Shape/Picture Format 分頁、Format Shape 窗格、群組、前後順序、對齊、替代文字**
- GenOffice：選取的圖形只有 8 個縮放控點；情境分頁只有 Chart Design 與 Table Design（`ExcelShell.tsx` 的 RibbonTab 定義）；程式碼中找不到群組、對齊、前後順序、替代文字的 UI；Insert 沒有 SmartArt。
- Excel（已知行為）：旋轉控點、圖片裁剪、黃色調整控點、Shape/Picture Format 分頁、Format Shape/Picture 窗格、Group／Arrange／Align、Alt Text。
- 備註：匯入的已旋轉圖形可正確顯示（有旋轉渲染），但不能由使用者操作旋轉。

**D4　圖形與圖片沒有右鍵選單**
- GenOffice：在選取的圖形與圖片上按右鍵，沒有任何選單出現（圖表有自己的選單）。
- Excel（已知行為）：Cut／Copy／Paste、Group、Bring to Front、Format Shape、Edit Alt Text 等。

**D5　功能區文字格式不作用於選取的圖形文字**
- GenOffice：選取圖形（含文字）後點 Home 粗體，狀態列顯示 `Applied to the current selection.`，實際套在格線的目前儲存格，不是圖形文字；雙擊進入圖形文字編輯後點功能區任何東西（含切換分頁），編輯器因失焦立即結束並提交，格式也沒套上。
- Excel（已知行為）：圖形／文字方塊選取時，Home 的字型、字級、顏色、對齊套用到圖形文字，編輯中保持選取。
- 這也是使用者在 Slides 回報問題的 Sheets 版本。
- 證據：g19-b.png、g20-a.png。

**D6　Text Box 沒有繪製模式**
- GenOffice：Insert → Text → Text Box 立即在目前選取位置插入，沒有選取、沒有進入文字編輯。
- Excel（已知行為）：游標變十字，拖曳繪製後立即進入文字輸入。
- 備註：Shapes 圖庫則有正確的繪製模式（單擊預設 1x1 英寸、拖曳自訂、Shift 正方形、Esc 取消，見 `shape-draw.ts`）。

**D7　刪除物件後焦點掉到 BODY，⌘Z 暫時失效**
- GenOffice：選圖表 → Delete，物件刪除成功，但 `document.activeElement` 變 BODY；此時按 ⌘Z 沒有反應，必須先點一次格線，⌘Z 才能復原（功能區 Undo 按鈕則可立即復原）。
- Excel（已知行為）：⌘Z 立即復原。
- 重現：點圖表 → Delete → ⌘Z。

**E1　樞紐分析表缺少欄位清單窗格與情境分頁，切片器／時間表要先存檔**
- GenOffice：建立與編輯都是置中的表單對話框（來源範圍、列、欄、值下拉），不是可拖放欄位的 PivotTable Fields 窗格；選取樞紐表輸出時沒有 PivotTable Analyze／Design 分頁；Slicer／Timeline 按鈕在未存檔時停用，狀態列提示「save to enable slicers, timelines, and refresh」；輸出範圍與來源範圍重疊時以紅字拒絕。另外建立樞紐表後再做工作表管理會出現 `Couldn't create the automatic recovery copy: A new pivot cannot be saved together with sheet management changes` 的錯誤提示（資料安全相關，建議另案追）。
- Excel（已知行為）：右側欄位窗格拖放、PivotTable Analyze／Design 分頁、插入切片器即時可用。

**E2　工作表分頁不能 Shift／⌘ 點擊建立群組工作表，右鍵選單缺項**
- GenOffice：Sheet1 → Shift+點 Sheet3，只是切換到 Sheet3（沒有群組）；⌘+點擊同樣。右鍵選單有 Delete、Copy、Rename、Change color、Hide、Unhide、Protect…，沒有 Insert、Move or Copy、Select All Sheets。
- Excel（已知行為）：Shift／⌘ 點擊群組工作表（標題顯示 [Group]），群組內編輯同步；選單有 Insert／Move or Copy／Select All Sheets 等。
- 備註：已有另一位 agent（sheets-tabbar）在處理分頁列，此項僅記錄。

### P2

| 編號 | 項目 | GenOffice | Excel 基準（已知行為） |
|---|---|---|---|
| B2 | Option+拖曳邊框複製 | Alt／Ctrl／⌘ 加拖曳都只是移動（來源被清空，值移到目的） | Option+拖曳為複製 |
| B4 | 名稱方塊多重區域 | 輸入 `A1,C3:D4` 後選取維持原狀（未解析） | 可接受逗號分隔多區域 |
| B5 | F5 開啟 Go To | F5 沒有反應；Ctrl+G 會開啟 Go To 對話框 | F5／Ctrl+G 皆可（F5 在 Mac 未實測，不確定） |
| C4 | 圖表介面形式 | 無 Chart Format 分頁、無圖表旁的「＋／筆刷／篩選」元素按鈕；Format 窗格是浮動面板直接蓋在格線與圖表上；雙擊圖表打開的小編輯器蓋住圖表右半 | 圖表元素按鈕、Chart Format 分頁、窗格停靠在右側 |
| C5 | 物件選取時的格線狀態 | 選取物件時目前儲存格仍高亮、名稱方塊仍顯示儲存格位址 | 選取物件後儲存格選取消失，名稱方塊顯示物件名稱 |
| D8 | 錨定選項 | 圖形固定為隨儲存格移動並調整大小（與預設一致），找不到 Move and size／Move but don't size／Don't move 的設定 UI | Format Shape → Properties 三選一 |
| E3 | Insert → Table | 直接以目前範圍建立表格，沒有 Create Table 對話框（範圍、含標題核取） | 先顯示 Create Table 對話框 |
| E4 | 表格選取輔助 | 滑鼠移到欄標題上緣／左上角沒有向下／斜向箭頭（游標是 move），沒有右下角調整大小把手（只有 Resize Table 對話框） | 向下箭頭選欄、左上角斜箭頭選表、右下角調整把手 |
| F1 | 註解 | 有角落標記；滑鼠停在標記上沒有出現註解卡片；Comments 面板是浮動覆蓋層，作者名稱欄可編輯 | 停留顯示註解卡片，窗格停靠 |
| F2 | 分頁預覽 | `page-break-preview.tsx` 為唯讀疊層（無指標處理），藍色分頁線不可拖曳、無法插入／移除分頁 | 可拖曳分頁線、右鍵插入／重設 |
| F3 | Ask AI 膠囊的層級 | 開啟 Icons 對話框時，膠囊仍浮在對話框邊緣之上（g23-icons2.png） | 無此類元件 |

## 建議修正順序

1. **A1 + A2 + D5（焦點與選取保持）**：功能區按鈕 `onMouseDown` 阻止預設焦點轉移（或操作後把焦點還給格線／儲存格編輯器／圖形文字編輯器），並在格內編輯與圖形文字編輯期間不讓點擊功能區造成失焦。三項同根，修一次解掉使用者回報的主因，Slides 同樣適用。
2. **C1（圖表元素選取）**：圖表內部事件在 pointer capture 之後被重導，需改為 pointerup 時依 `elementFromPoint` 判斷元素，或僅在拖曳超過門檻才 capture。順便驗證 Format 窗格各元素分頁。
3. **B1 + D2 + C2 + C3（選取與操控基礎）**：⌘+點擊加選（格線與物件）、物件 Shift／⌘ 加選與框選、方向鍵微調、Shift／Option 縮放與圖片鎖定比例。
4. **D1 + D6 + D7**：繪製圖形後立即選取並避免格線選取被改；Text Box 繪製模式；刪除物件後把焦點還給格線。
5. **B3**：Ask AI 膠囊避開填滿控點（例如改放選取範圍右上外側，或延遲／僅在懸停時出現）。
6. **D3 / D4 / E1**：Shape/Picture Format 分頁與右鍵選單、旋轉／裁剪、樞紐欄位窗格——屬功能補齊，工作量大，可分批。
7. P2 項目依需要排入，E3、E4、B2 成本低。

## 尚未驗證項目與限制

- **所有 Excel 基準都沒有並排實測**（原因見「方法與限制」），建議補驗證的重點：A1／A2（確認 Excel 點功能區後焦點行為）、B1（⌘+點擊）、B5（F5）、D1（繪製後立即選取）、D2（Shift 多選）、C3（圖片鎖定比例預設）、E3（Create Table 對話框）。
- 未驗證：SmartArt（Insert 沒有入口，視為缺少）、圖案群組、Select Objects 框選、Selection Pane、Alt Text（程式碼中找不到，未在 UI 窮舉）、迷你圖的選取與情境分頁、超連結點擊與編輯、凍結窗格分隔線拖曳（Univer 原生，未測）、切片器與時間表（需先存檔才可建立，本輪未存檔）、圖表資料範圍的藍紫框拖曳、圖表更換類型對話框細節、條件式格式圖示、F5 在 Excel 的實際行為、隱藏／取消隱藏列欄的選取細節（只確認顯示小箭頭）。
- 圖表選取後再點格線在乾淨情境可取消選取；曾在 Format 窗格開著時觀察到物件選取殘留，判斷為窗格開啟所致，未列為差異。
- 圖片只用 Icons 產生的點陣圖測試，未測貼上圖片與匯入既有 xlsx 圖片。
- 測試過程 Electron 曾被外部中止一次（額度中斷），已重啟後補測 D7，其餘結果不受影響。
- 測試用資料皆為合成資料；Excel 的 Book1 空白活頁簿仍開啟、未存檔。
