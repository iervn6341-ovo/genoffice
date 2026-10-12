# Microsoft 365 Mac 工具細項比對：暫停紀錄

記錄日期：2026-10-12（Asia/Taipei）。狀態：使用者要求暫停；尚未完成所有工具細項的完整比對，也尚未進行功能修正。

## 使用者要求與比對基準

- 比對 GenOffice 的 Docs、Sheets、Slides 與 **Microsoft 365 Mac 桌面版**。
- 重點是每個工具的下拉選單、對話框、情境工具列與實際功能，不只比較主工具列有沒有按鈕。
- 使用者特別指出：頁首頁尾功能不同、方程式旁缺少「符號」、方程式功能不足。
- 使用者提供了 PowerPoint 原生頁首頁尾、GenOffice 頁首頁尾及原生方程式工具列三張截圖，並授權使用 computer use 檢查 Office。
- 本機 Word、Excel、PowerPoint 的版本均為 **16.113.3，build 16.113.26092714**。程式碼檢查基準為 HEAD `3ae13422`。

## 已經做的事情

1. 閱讀目前工具列、對話框、動作處理與文件引擎程式碼，追查按鈕最後插入的資料類型。
2. 透過 computer use 實際開啟 Mac Office 主功能區與以下選單／對話框，取得原生介面內容。
3. 查閱 Microsoft 官方說明，補充 Mac 平台與功能語意；沒有把先前 Windows 參考清單當作 Mac 實機證據。
4. 對 Docs 共用的 LaTeX → OMML 轉換執行小型語意檢查，結果記錄於下方。
5. 目前**未修改產品程式碼、未提交 commit、未建立 PR**。本次新增的交付物是這份暫停紀錄；完整細項報告尚未產出。

Office 操作使用空白測試文件／活頁簿／簡報。使用者原本開啟的 23 張投影片簡報僅作讀取參考，沒有儲存修改。部分空白測試視窗可能仍開啟，沒有另存為專案檔案。

### 已成功實際查看的 Office 介面

| 應用程式 | 已查看範圍 |
| --- | --- |
| Word | 常用、插入、繪圖、設計、版面配置、參考資料、郵件、校閱、檢視主功能區 |
| Word | 頁首內建圖庫、頁首及頁尾情境工具列、進階符號與特殊字元分頁 |
| Word | 方程式情境工具列及全部 11 類結構選單 |
| Word | 字型對話框的字型／進階兩分頁；段落對話框的縮排與行距／分行與分頁／中文印刷樣式三分頁 |
| Word | 插入表格的 10 × 8 格選擇圖庫及其選單；未成功取得「插入表格」完整對話框 |
| Excel | 常用、插入、頁面配置、公式、資料、校閱、檢視主功能區；確認有 Automate 分頁 |
| Excel | 儲存格格式的數值、對齊、字型、框線、填滿、保護六分頁 |
| Excel | 頁面設定的頁面、邊界、頁首／頁尾、工作表四分頁；自訂頁首對話框與頁首頁尾情境工具列 |
| Excel | 資料剖析精靈三步驟與進階數字分隔設定；假設分析完整選單 |
| PowerPoint | 常用、插入、繪圖、設計、轉場、動畫、投影片放映、錄製、校閱、檢視主功能區 |
| PowerPoint | 頁首及頁尾的「投影片」「備忘稿與講義」兩分頁；圖案格式情境工具列；方程式情境工具列的結構入口 |

Excel 深入對話框一開始多次沒有成功開啟，**後來已成功取得上述六個格式分頁、四個頁面設定分頁、自訂頁首、資料剖析與假設分析**。接續時不能沿用「Excel 深入對話框全部未驗證」的舊狀態。

## 優先確認的差異

### Slides：頁首及頁尾

Mac PowerPoint 的投影片分頁有日期／時間、自動更新或固定、日期格式清單、投影片編號與起始值、頁尾、標題投影片不顯示、預覽，以及「套用」「全部套用」。備忘稿與講義另有頁首、頁尾與頁碼。

GenOffice 目前有日期、自動更新勾選、編號、頁尾文字，但缺少：

- 投影片／備忘稿與講義分頁及備忘稿／講義頁首設定。
- 日期格式選擇、明確的自動／固定單選模式。
- 編號開始於、標題投影片不顯示、預覽。
- 僅套用目前投影片；目前介面只有全部套用。

另外兩項是行為差異，不能只補按鈕：

- 重新開啟時 `dateAuto` 固定預設為 `true`，讀取介面也沒有回傳日期模式；固定日期設定無法正確回填。
- 套用會刪除既有日期／頁尾／編號 placeholder 再重建，採固定位置、12pt、灰色 `898989`，不能保留原先母片或個別物件的位置與樣式。

引擎已有 `slideIndexes?` 參數，但 UI／IPC 尚未傳遞，可作為「套用目前投影片」的既有接點。自動日期已使用 `datetime1` 原生欄位，不能誤寫成完全沒有自動日期。

來源：[對話框](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/slides/src/renderer/components/InsertDialogs.tsx:121>)、[套用動作](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/slides/src/renderer/insert-actions.ts:366>)、[IPC](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/slides/src/main/slides-main.ts:3727>)、[引擎](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/packages/pptx-engine/src/headerfooter.ts:58>)、[讀取設定](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/packages/pptx-engine/src/headerfooter.ts:160>)。Mac 基準亦見 [Microsoft 頁首頁尾說明](https://support.microsoft.com/en-us/powerpoint/add-page-numbers-date-and-time-or-footer-text-to-all-slides-in-powerpoint-for-mac)。

### Docs：頁首及頁尾

- 原生 Word 有 21 個內建頁首樣式，以及編輯／移除命令；專案下拉面板目前是單一文字輸入、確定／取消。
- 原生編輯模式會出現頁首頁尾情境工具列：日期、功能變數、圖片、移至頁首／頁尾、前／後節、連結到前一節、顯示文件文字、上下距離及關閉。專案缺少這套完整操作介面。
- **第一頁不同、奇偶頁不同已實作**，不應重複列為缺項。
- 頁碼已有上／下各左中右共六位置、移除、七種編號格式、起始值與續前節；章節編號等更多設定尚未提供，原生頁碼詳細對話框仍需補驗。
- 頁首頁尾區域已有直接文字編輯，不能說完全無法編輯；但尚未等同 Word 能混排圖片、欄位及多段富文字的完整編輯流程。

來源：[頁首頁尾輸入面板](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/docs/src/renderer/components/ribbon-insert-tab.tsx:111>)、[工具群組](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/docs/src/renderer/components/ribbon-insert-tab.tsx:1185>)、[頁碼格式](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/docs/src/renderer/App.tsx:6707>)。

### Sheets：頁首及頁尾

實機 Excel 已確認頁首頁尾工具包含頁碼、總頁數、日期、時間、檔案路徑、檔名、工作表名、圖片與圖片格式，以及移至頁首／頁尾。頁面設定亦有預設頁首／頁尾清單、預覽與自訂編輯入口。

專案已有頁首／頁尾各左中右三欄，共六個文字輸入，並有 `&P`、`&N`、`&D`、`&T`、`&F`、`&A` 六個快捷欄位碼。尚缺：

- 檔案路徑快捷插入、圖片及圖片格式、頁首文字格式工具。
- 第一頁不同、奇偶頁不同、隨文件縮放、與頁面邊界對齊四項設定。
- 預設頁首頁尾清單及原生預覽／頁面配置編輯流程；專案目前沒有 Page Layout 檢視。

輸入框可以保留使用者手打的欄位碼，因此應描述成「缺少對應 UI／編輯流程」，不能直接推論所有手打碼都無法儲存。

來源：[六欄編輯與快捷碼](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/sheets/src/renderer/HeaderFooterDialog.tsx:45>)。

### 三套方程式實作不同

| 應用程式 | 目前實作 | 與 Office 的關鍵差異 |
| --- | --- | --- |
| Docs | 8 個內建公式、LaTeX 輸入與 MathML 預覽；可行內／獨立插入、再次編輯，儲存 OMML 數學結構 | 缺少 11 類結構選單、直接操作佔位區的情境編輯、UnicodeMath 與專業／線性轉換介面；轉換支援仍是子集 |
| Sheets | 6 個公式範本、LaTeX 輸入與預覽 | 插入前轉成 PNG，最後是圖片；無法按 Office 數學物件的方式再次編輯結構 |
| Slides | 12 個 Unicode 文字範本與單一文字輸入 | 插入 Cambria Math、28pt、斜體文字方塊；不是原生 OMML 數學物件，沒有結構編輯器 |

來源：[Docs 方程式](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/docs/src/renderer/components/EquationModal.tsx:77>)、[Sheets 方程式](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/sheets/src/renderer/EquationDialog.tsx:7>)、[Slides 插入結果](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/slides/src/renderer/insert-actions.ts:389>)。

Word 實機逐一開啟的結構選單如下；計數為圖庫範本項目，包含常用公式，**不是唯一數學符號數量，也不能直接套用為 Excel／PowerPoint 的相同數量**。

| 結構 | 範本數 | 已看到的細項 |
| --- | ---: | --- |
| 分數 | 9 | 直式、斜式、橫式、小型、微分及常用分數 |
| 上下標 | 8 | 上標、下標、上下標、左側上下標、常用範本 |
| 根號 | 6 | 平方根、指定次方根、二／三次方及常用範本 |
| 積分 | 21 | 單／雙／三重積分、上下限配置、封閉積分、微分 |
| 大型運算子 | 40 | 總和、積、同乘積、聯集／交集、邏輯運算與上下限配置 |
| 括弧 | 41 | 成對、分隔、單邊、分段案例、堆疊、常用範本 |
| 函數 | 27 | 三角、反三角、雙曲、反雙曲及常用範本 |
| 加強記號 | 27 | 點／帽／橫線／箭頭等記號、公式方框、上下橫線 |
| 極限與對數 | 8 | log、指定底數、ln、lim、min、max 及常用範本 |
| 運算子 | 21 | 基本運算子、帶說明的箭頭與常用範本 |
| 矩陣 | 22 | 空矩陣、點線、單位矩陣、括弧矩陣、稀疏矩陣 |
| 合計 | 230 | Word 此版本的圖庫項目 |

已執行 LaTeX → OMML 小型檢查：`\frac`、`\dfrac`、`\tfrac` 產出相同結構；指定次方根 `\sqrt[3]{x}`、`bmatrix` 可轉換；`\boxed`、`\overset`、`\underset` 不支援。這不是完整 UI／儲存／重開驗證。來源：[數學轉換器](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/packages/docx-engine/src/math.ts>)。

### 符號

- Docs 有固定 50 字元面板；原生 Word 的進階符號有字型、符號／特殊字元分頁、字元資訊、捷徑資訊、自動校正與鍵盤捷徑設定。
- Sheets 有五類、123 個面板項目（包含重複字元）；插入動作附加到儲存格內容，不是儲存格文字游標位置的插入。原生 Excel 符號對話框細項尚未取得。
- **Slides 符號工具群組只有方程式按鈕，確實缺少旁邊的「符號」**。原生 PowerPoint 插入功能區已確認兩者相鄰，符號需適當文字編輯情境才啟用；符號對話框本身尚未成功取得。

來源：[Docs 字元表](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/docs/src/renderer/components/ribbon-insert-tab.tsx:57>)、[Sheets 符號](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/sheets/src/renderer/SymbolDialog.tsx>)、[Slides 符號群組](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/slides/src/renderer/components/RibbonInsertTab.tsx:474>)。

## 其他已查到的細項

### Docs

- 字型對話框：Word 有中／英文字型分開選擇、底線樣式／顏色、強調標記、雙刪除線、小型大寫、隱藏，以及進階字距／位置／比例／OpenType 設定。專案目前較簡化，缺少這些完整欄位與進階分頁。[來源](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/docs/src/renderer/components/ContextMenu.tsx:545>)
- 段落對話框：Word 有分頁控制、與下段同頁、段落中不分頁、中文首尾禁則、標點壓縮、中英／中數字間距、格線與定位點等。專案提供主要對齊、行距、縮排與段前後間距，缺少完整三分頁。部分底層屬性已存在，應列為缺少可編輯 UI，不能一律說引擎不支援。[來源](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/docs/src/renderer/components/ContextMenu.tsx:792>)
- 常用「排序」明確 disabled；原生 Word 有排序、亞洲方式配置、注音標示、圍繞字元等入口。[來源](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/docs/src/renderer/components/Ribbon.tsx:3497>)
- 參考資料：來源管理已有簡化欄位與五類來源，但內建引文插入一般文字，書目使用固定生成格式，沒有原生引用樣式切換／欄位更新完整流程。**Zotero 整合已存在**，不能寫成沒有引用功能。[來源](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/docs/src/renderer/components/ribbon-references-tab.tsx:410>)
- 標號有 Figure／Table／Equation 三標籤及描述、SEQ 欄位，但自訂標籤、位置、章節編號等完整對話框未提供；目錄、註腳／章末註、索引相關更多設定還需逐一核對。
- 樣式已有套用面板；建立／修改／管理樣式流程尚不足。郵件合併主分頁缺少。
- 表格設計／版面配置已有合併、分割儲存格、對齊、AutoFit、重複標題列、屬性等，不能概括成沒有表格工具；完整原生表格情境工具細項尚需補驗。
- 保護文件已有四模式及密碼介面，不应重複列為完全缺少。

### Sheets

- 儲存格格式六分頁已存在，但原生 Number 有 **Special** 類別，專案目前 11 類沒有它。
- 對齊：原生可輸入角度、縮小字型以配合、合併儲存格及分散對齊調整；專案旋轉只有 0／±45／±90／直書六預設，對話框沒有完整文字控制項。水平／垂直對齊下拉的所有原生選項仍需展開補驗。
- 字型：原生有上標／下標、Normal font、底線樣式選單；專案較簡化。原生底線選單的全部項目尚未展開。
- 框線：原生有對角線、各邊／內部線及可點選預覽，線條樣式含 Dash Dot、Dash Dot Dot、Medium 變體等；專案對話框只有七種線型、框線預設，沒有完整逐邊預覽與對角線編輯。
- 填滿：原生有背景色、圖樣色、圖樣樣式；專案只有背景色／無填滿。[Microsoft Mac 說明](https://support.microsoft.com/en-us/excel/apply-or-remove-a-cell-shading-in-excel-for-mac)
- 保護：原生回填 Locked／Hidden 當前狀態；專案草稿預設為「不變更」，尚未回填兩項選取狀態。格式對話框來源：[FormatCellsDialog](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/sheets/src/renderer/FormatCellsDialog.tsx>)、[草稿／命令](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/sheets/src/renderer/format-cells.ts>)。
- 資料剖析：原生三步驟有分隔／固定寬度、其他分隔符、多分隔符合併、文字辨識符、逐欄一般／文字／日期／略過、目的範圍與小數／千位分隔設定。專案只有逗號、分號、空白、Tab 四個快速命令。[來源](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/sheets/src/renderer/ExcelShell.tsx:2613>)
- 假設分析：原生有情境管理員、目標搜尋、資料表；專案只有目標搜尋。[來源](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/sheets/src/renderer/ExcelShell.tsx:2653>)、[Microsoft 說明](https://support.microsoft.com/en-us/excel/introduction-to-what-if-analysis)
- 公式「最近使用」目前實際打開 All 目錄；從選取範圍建立名稱只提供頂列／左欄。移除箭號與錯誤檢查的原生選單細項尚需展開。
- 頁面設定：專案功能區有三種邊界預設、七種紙張、設定／清除列印範圍、重複頂列與選取列；原生還有自由邊界、頁首頁尾距離、水平／垂直置中、起始頁碼、重複左欄、列印順序、黑白／草稿／註解等。必須再核對專案 PrintDialog，避免把移到列印視窗的功能誤列成全應用程式缺項。
- 新增註解與備註都接到 `note-open`，目前不是原生 threaded comments 的回覆／解決討論流程。[來源](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/sheets/src/renderer/ExcelShell.tsx:2830>)
- 匯入檔案的表格名稱、調整大小、轉為範圍、標題列、合計列、篩選按鈕受限制；樣式及條紋等已有可編輯功能。[來源](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/sheets/src/renderer/ExcelShell.tsx:1550>)
- 插入圖表目前九種主要按鈕，沒有對應全部原生階層、統計、瀑布、地圖等圖庫。實際圖表編輯情境工具與各子類型尚需補驗。
- 選擇性貼上已有 12 類貼上、五種運算、略過空白、轉置；切片器、時間表、新表格樣式也已存在，不能沿用舊報告誤列為缺少。

### Slides

- SmartArt 插入一般圖形群組，沒有原生 diagram 資料／版面／色彩／樣式結構，所以後續 SmartArt 階層及樣式編輯語意不同。[來源](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/packages/pptx-engine/src/smartart.ts:1>)
- Zoom 目前是投影片跳轉超連結圖形，完整原生 Zoom 編輯流程尚不等同。
- 轉場清單為 **12 項（含無），即 11 種效果**，不能寫成 12 種效果；持續時間、按滑鼠／延遲換頁、效果選項與全部套用已有實作，缺少原生聲音選擇等。[來源](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/slides/src/renderer/components/Ribbon.tsx:284>)
- 動畫有 19 個效果預設、五種路徑，開始／持續／延遲／觸發入口已有；動畫複製格式不足，動畫窗格深層選項尚需核對。
- 投影片放映中的設定放映、錄製按鈕目前 disabled；原生 Mac 此版另有「錄製」主分頁。播放旁白、使用時間、媒體控制與字幕設定尚不足。[來源](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/slides/src/renderer/components/Ribbon.tsx:2956>)
- 設計目前以主題、背景及 16:9／4:3 為主；原生變體的色彩、字型與背景樣式入口及自訂尺寸流程需補齊／進一步核對。
- 原生檢視有投影片母片、講義母片、備忘稿母片；專案沒有相應完整母片編輯流程。
- 圖案格式情境功能區目前主要是變更圖案、快速樣式、填滿、外框、效果；原生有合併圖案、文字填滿／外框／效果、替代文字、重排物件、鎖定比例等。部分排列／尺寸功能位於專案其他位置，須區分入口位置差異與真正缺少功能。[來源](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/apps/slides/src/renderer/components/Ribbon.tsx:3674>)
- 繪圖缺少原生套索、筆跡轉文字／圖案／數學等流程。
- **文字方向四種模式、垂直對齊、表格情境工具、自訂放映、排練時間已實作**。專案也已有螢幕錄製入口。
- 此次原生 Mac Insert 未看到 Windows 清單中的 Photo Album／Screen Recording，Draw 未看到 Ink Replay；不能單憑 Windows 參考清單把它們列為此 Mac 基準的缺項。

## 尚未完成與接續順序

以下尚未驗證完整一致性，沒有把「按鈕存在」視為功能全數通過：

1. 先整理正式細項對照表：每項標示原生實機／官方 Mac 說明／程式碼判讀／待驗證，以及 UI、資料語意、儲存重開四層結果。
2. 補看 Word 頁碼格式、目錄、註腳設定、表格屬性／表格情境工具、框線網底、圖片與圖表深層對話框。
3. 補看 Excel 對齊下拉、底線選單、符號、選擇性貼上、資料驗證、排序、條件格式、樞紐分析與圖表情境工具；比對專案 PrintDialog 與原生頁面設定。
4. 補看 PowerPoint 符號對話框、完整字型／段落、表格／圖片／圖表／SmartArt 情境工具、動畫窗格、放映設定及母片。
5. 最後補測可編輯資料結構與存檔重開，尤其固定日期回填、目前投影片／全部套用、保留頁尾樣式、方程式重新編輯、符號在游標處插入。
6. 使用者目前只要求檢查，尚未要求實作修正；接續不應自行把全部差異開始改寫。

最後嘗試開啟 Excel Data Validation 時，computer use 回報 **`Sky Computer Use native pipe closed before response`**。此項沒有取得內容，須重建／重新選取 CUA app 並取得新介面狀態再操作。PowerPoint 符號對話框與 Word 插入表格完整對話框先前亦未成功取得，不能標成已實際驗證。

CUA 工作階段原有 `uiAudit` 陣列保存原生 AX 觀察；這是工具工作階段記憶，未匯出成檔案，不能假設下次仍可使用。若仍存在可用於整理；若不存在，以本紀錄接續並補取必要證據。任何舊 AX 索引都不能直接重用。

## 專案狀態與既有資料

- 舊的工具列／版面檢查位於 [ribbon-dropdown-audit-2026-10-11.md](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/ribbon-dropdown-audit-2026-10-11.md>) 與 [handoff-ribbon-audit-2026-10-11.md](</Volumes/Kingston NV2/Program_workspace/swift_workspace/genoffice/docs/qa/handoff-ribbon-audit-2026-10-11.md>)。其中部分缺項已有後續修正，且部分參考源為 Windows，接續需以目前程式碼與 Mac 實機覆核。
- 先前通過的工具列 E2E 驗證主要針對布局、折疊与常見操作，不能當作本次細項一致性的證明。本次沒有新增測試檔案或重跑整套 E2E。
- 暫停前沒有任何既有 tracked 檔案變更。工作區原有 `.claude-flow/`、`apps/docs/.claude-flow/`、`e2e/_probe/`、`test/`、`setup-genoffice-qa-skill.sh`、一份 PDF 及 `第四次估咪簡報.pptx` 等未追蹤內容；已保留，不能視為本次新增或刪除。

建議後續實作優先次序（僅建議，尚未執行）：頁首頁尾狀態／套用範圍與樣式保留 → Slides 符號及三套原生可編輯方程式 → Docs 字型／段落與 Sheets 格式／資料剖析 → SmartArt、母片及其他深層工具。

## 暫停後討論的 agent 分工與模型設定

使用者希望配置三位檢查員、三位實作者與一位監督員，分別涵蓋 Docs、Sheets、Slides。尚未啟動 agent；目前仍維持暫停。

| 角色 | 使用者指定模型 | 思考設定 |
| --- | --- | --- |
| Docs／Sheets／Slides 檢查員 | GPT-6 Luna（`gpt-6-luna`） | 使用者稱「Extra」；目前解讀為 Extra High（`xhigh`） |
| Docs／Sheets／Slides 實作者 | GPT-6 Sol（`gpt-6-sol`） | Medium（`medium`） |
| 監督員 | 使用者尚未指定 | 使用者尚未指定 |

目前工具限制同時最多四個 active agents，包含主 agent。已向使用者建議主 agent 擔任監督，分階段安排三位檢查員、三位實作者與獨立複驗；此執行方式尚未正式啟動。建立指定模型／思考設定的子 agent 時，需要使用 `fork_turns: "none"`，並在任務訊息明確提供本紀錄、比對基準、檔案範圍及交付要求。

監督責任還需包含共用模組的單一負責人、修改衝突協調、優先順序與整合驗收；檢查員需回頭複驗，並覆蓋檔案儲存重開、Office 往返、Undo／Redo、窄視窗與鍵盤操作。
