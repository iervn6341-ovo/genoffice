# GenOffice Swift 與 Rust 原生重構實作計畫

建立日期：2026 年 10 月 10 日  
產品方向：先完成 macOS，再延伸到 iPadOS 與 iOS。  
優先順序：功能正確與檔案保真、低資源佔用、互動效能、原生外觀、可跨平台的核心。開發時間不作為選擇架構的限制。  
程式碼定位基準：qa/local-fixes，HEAD 58cf824，另有尚未提交的異動；開始實作時重新記錄基準，不得覆蓋現有工作。

以 Rust 建立共用文件與編輯核心，以 Swift 建立 Apple 原生應用。Rust 保存文件狀態，Swift 持有畫面需要的局部資料；兩側透過嵌入式函式庫、版本化接口與批次操作合作。核心不依賴 Node、Electron、DOM、桌面子程序或 Apple UI 框架。macOS、iPad 與 iPhone 使用同一份核心程式碼，分別建置平台二進位。

此架構是實作方向，不代表混合語言已被證明比純 Swift 更快。效能決策必須用相同文件、相同硬體與 release build 驗證；不得用刪除功能、降低文件品質或省略重算來取得較好的數字。

## 1 實作者開始工作的方式

1. 讀取本文件、儲存庫適用的 AGENTS.md、CLAUDE.md 與 CONTRIBUTING.md。先確認工作目錄和未提交異動。
2. 從 N01 開始，以第 10 節的相依關係選擇可實作任務。未完成的前置驗證不能標記成已完成。
3. 每次只交付可驗收的範圍，在本文件第 12 節記錄結果、產物、測試、效能證據及下一項任務。
4. 先建立功能基準與 Swift Rust 接口原型，再大量搬移模組。原生應用建立於新的 native 目錄，既有 JS 版保留作為對照與過渡產品。
5. 資料結構、公開接口、相容性或平台範圍改變時，先更新對應架構決策紀錄與本文件，再繼續實作。一般技術選擇可按本文件自主處理；改變產品範圍時另取得使用者決定。

任務狀態統一使用 TODO、IN_PROGRESS、DONE、BLOCKED。DONE 必須具有產物、驗收結果與可重現證據。實作存在、能編譯、能顯示畫面，都不等於功能完成。

## 2 範圍與平台承諾

### 2.1 保留的產品範圍

- Docs、Sheets、Slides、PDF、Markdown、HTML 六種編輯器。
- 新建、開啟、儲存、另存、外部修改處理、復原重做、搜尋取代、剪貼簿、列印、匯出與大檔案處理。
- AI provider、串流、工具呼叫、附件、網路與媒體工具、選取範圍操作、批次編輯、變更檢視、取消與回復。
- CLI、MCP、無介面轉換與匯出、桌面定位控制、技能整合、設定、專案與聊天資料遷移。
- 語系、淺深色主題、無障礙、快捷鍵、CJK 與 RTL、多螢幕簡報、簽章及發行。

功能清單以目前程式碼、測試及 QA 文件合併建立；不能只採用 README 宣稱。每個功能分別記錄讀取、保留、顯示、編輯、復原、儲存重開與匯出能力。原本未完成的功能與既有缺陷另列，不冒充新版本已完成，也不默認所有 Office 規格都已受支援。

### 2.2 平台邊界

| 平台                  | 本次策略                                        | 驗收要求                                                          |
| --------------------- | ----------------------------------------------- | ----------------------------------------------------------------- |
| macOS                 | 第一個完整原生產品                              | 六種編輯器與桌面整合通過功能矩陣                                  |
| iPadOS                | 核心從第一階段即能建置，後續完成原生 UI         | 同一套文件語意、儲存與重算；補觸控、鍵盤、Files、生命週期測試     |
| iOS                   | 核心從第一階段即能建置，後續完成適合小螢幕的 UI | 不默認只讀；先逐項決定各功能的入口與能力，不能把 Mac 介面直接縮小 |
| Windows Linux Android | 為核心保留移植邊界，不納入第一階段 UI 承諾      | Rust 核心能在 Windows Linux 做非 UI 測試；平台服務接口不綁 Apple  |

CLI、技能安裝、外部 provider 子程序與本機桌面控制是桌面能力；手機的文件與 AI 編輯仍須保留。需要桌面程序的 provider 在手機端應明確顯示能力限制，不能默默換模型或宣稱完全相同。未來採用遠端代理須另訂連線與隱私需求。

跨裝置可執行與跨裝置同步分開驗收。核心先具備穩定 ID、版本、可持久化命令與衝突偵測；iCloud 文件同步、聊天同步與同時編輯合併由 N34 處理，不能把檔案同步視為即時協作。

Rust 官方支援 iOS iPadOS 真機與模擬器 target；本專案的依賴能否使用，仍以 N03 N13 N22 的實際 build 與真機驗收決定。[官方 target 說明](https://doc.rust-lang.org/rustc/platform-support/apple-ios.html)

## 3 Swift Rust 與原生函式庫的分工

| 責任                                          | 主要實作                                              | 邊界與原因                                                                |
| --------------------------------------------- | ----------------------------------------------------- | ------------------------------------------------------------------------- |
| 檔案 package、ZIP XML、來源保留、patch writer | Rust                                                  | 批次處理與記憶體控制集中，格式核心可跨平台                                |
| 文件模型、穩定 ID、版本與編輯交易             | Rust                                                  | 每份文件只有一個權威狀態，UI AI CLI 使用相同命令                          |
| XLSX 解析與公式計算                           | Rust                                                  | 遷移現有引擎，逐項替換 Univer 語意與 TS gateway；不保留兩套常駐完整活頁簿 |
| DOCX PPTX 幾何排版、分頁與圖形語意            | Rust                                                  | 輸出可共用的 layout snapshot，螢幕與匯出採用相同結果                      |
| 字型解析、字形塑形與度量                      | 可替換平台服務，Apple 初始實作用 Swift 呼叫 Core Text | Rust 透過批次接口取得字形與來源映射；不把 Core Text 物件存進核心模型      |
| 原生介面、焦點、選取呈現、觸控與 IME          | SwiftUI 與 AppKit UIKit                               | 系統輸入與輔助使用能力留在平台端                                          |
| 畫布繪製、文字、圖片與 PDF 輸出               | Swift 平台 renderer                                   | Core Graphics 與 Core Animation 為基準，Metal 只有量測證明收益後才導入    |
| PDF 解析、內容修改、抽取與頁面 raster         | Rust wrapper 與 PDFium 原生函式庫                     | 重用成熟引擎，去除 WASM／JS 包裝；逐平台驗證原生建置及現有 API            |
| OCR                                           | Swift Vision adapter                                  | 可重用既有 vision-ocr.swift 的演算法，改為程序內服務                      |
| AI 代理狀態機、工具 schema、交易與歷程        | Rust                                                  | 工具操作共用文件核心；狀態與上下文有明確容量限制                          |
| AI 網路傳輸、憑證、背景與連線生命週期         | Swift URLSession、Keychain 與平台 adapter             | provider 協定解析與狀態機可攜；平台 transport 不出現在文件 crate          |
| 專案與聊天儲存、命令歷程                      | Rust store，Swift 提供儲存位置與檔案授權              | 遷移現有 JSON JSONL，保留穩定 ID；存儲技術需按 N29 量測選定               |
| HTML 內容預覽                                 | Swift WKWebView                                       | 瀏覽器引擎只用於 HTML 功能及明確隔離的圖表／公式 fallback                 |
| CLI MCP                                       | Rust executable 與可替換網路服務                      | 桌面與無介面模式共用核心，不啟動 Electron                                 |

PDFium 是 C C++ 原生依賴，不是 Rust 重寫的 PDF 引擎。使用它的理由是功能完整性與已知行為；建置、效能及相容性仍要驗證。Rust 與 Swift 混合不表示所有第三方依賴都必須重寫。

既有 Rust 引擎只是遷移來源。現有 Univer 的公式、資料驗證、互動和渲染能力，以及 xlsx-gateway 的 TS 寫入功能，不能視為已由 Rust 完成。

### 3.1 目標資料流

```text
Swift 原生 Shell 與六種編輯器
  ├─ AppKit UIKit 輸入與無障礙 adapter
  ├─ Core Text 字型與 shaping adapter
  ├─ Core Graphics renderer 與原生列印
  ├─ URLSession Keychain Vision 文件授權
  └─ HTML 專用 WKWebView
                   ↕ 批次 C ABI 與 Swift wrapper
Rust session 與 transaction 核心
  ├─ DOCX XLSX PPTX Markdown HTML source engines
  ├─ layout snapshots 公式依賴圖 資源管理
  ├─ AI 命令 Undo journal 專案儲存
  └─ PDFium wrapper 與檔案轉換 IR
                   ↕
原檔 原始 XML 媒體 命令 journal 可重建快取

Rust CLI MCP → 相同的格式與交易接口
Apple CLI 的排版匯出 → 共用 Swift renderer，無可見視窗
```

Rust 是文件內容的權威來源；Swift 的可見頁、選區、工具列值及輸入 composition 是有版本的局部投影。不得把完整文件轉成 Swift Observable 模型，又在 Rust 保存完整副本。

## 4 原生接口與資料所有權

### 4.1 預設採用 C ABI

以 Rust staticlib、cbindgen 產生的 C header、XCFramework 與手寫 Swift wrapper 為初始接口。Rust repr C 的固定佈局記錄、帶長度字串、opaque handle 與 buffer 構成穩定邊界。UniFFI 可在 N04 對照，但只有實測不增加關鍵路徑配置與複製成本時才採用；不得同時維護兩個對外產品接口。

每個 Apple 平台與 simulator 變體分別編譯，按 Apple 的 XCFramework 流程包裝；Rust 與 Swift 資料所有權遵守跨語言接口合約。[XCFramework 說明](https://developer.apple.com/documentation/xcode/creating-a-multi-platform-binary-framework-bundle)、[Rust FFI 說明](https://doc.rust-lang.org/stable/nomicon/ffi.html)

N03 固定的初始 build matrix 為 aarch64-apple-darwin、aarch64-apple-ios、aarch64-apple-ios-sim；若保留 Intel Mac，增加 x86_64-apple-darwin 及需要的 simulator 變體。各 build 使用對應 Xcode SDK 與一致 deployment target，iPad 使用相同 iOS target。N04 的 ffi crate 輸出 staticlib，C header 在 CI 重新產生與比對，不手動複製 Rust 內部結構。[Rust linkage](https://doc.rust-lang.org/reference/linkage.html)、[cbindgen](https://github.com/mozilla/cbindgen)

以操作意圖呼叫核心，例如 open_document、apply_transaction、query_visible_range、request_layout、read_resource、save_document。避免每個儲存格、字元、glyph 或 AI token 一次跨語言呼叫。JSON 可用於設定、外部 MCP 與除錯；畫面更新、layout、cell range 與媒體不得依賴整份文件 JSON 或 base64 傳輸。

### 4.2 必須明確定義的合約

- SessionHandle、SnapshotHandle、ResourceHandle 使用帶 generation 的 ID；關閉後的舊 handle 返回明確錯誤，不能錯誤指向新 session。
- session 擁有模型、來源、編輯 journal 與快取；snapshot 不可變，引用的 buffer 在 snapshot 釋放前有效。釋放 snapshot 不會關閉 session。
- Rust 配置的 buffer 只能透過對應 Rust release 函式釋放。Swift 以 RAII wrapper 管理，不可自行 free。
- Swift 輸入 buffer 的有效期與複製規則明文定義。非同步工作不得保留 withUnsafeBytes scope 的指標；需要持有時由核心接收所有權或執行一次明確複製。
- 零複製只用於生命週期可證明的不可變資料；不為宣稱零複製而暴露可變懸空指標。記錄每個熱路徑的複製次數與 bytes。
- FFI schema 與 layout version 分開版本化。驗證 size、offset、alignment、endianness、可選值、status 與指標長度；布林與 enum 使用明確整數表示。
- 錯誤包含 code、operation、recoverability 與必要上下文；文件內容與憑證不進效能紀錄。
- 所有 exported function 不得讓 Rust panic 越過外語言邊界。Result 處理一般錯誤；panic unwind build 在入口捕捉並使受影響 session 失效。panic abort、OOM、native crash 無法靠 catch_unwind 恢復，必須從持久化 journal 復原。

### 4.3 版本與文字位置

命令帶 document ID、base revision、operation ID 與穩定 block／element／sheet ID。舊版本操作拒絕或按明確規則重定位；不允許 AI 使用過期座標直接寫入。批次操作全成或全退，且合併成合理的復原單位。

存檔先固定 revision 的不可變 save snapshot，背景只寫該版本；使用者繼續編輯時，新 revision 仍然 dirty。成功寫入 snapshot 不得清除之後的編輯。比較外部檔案 fingerprint，再原子提交；external change 與同步衝突不能用最後寫入者直接覆蓋。字型、公式與版面快取帶 revision，舊工作完成時不得覆蓋新畫面。

文字核心用 UTF 8 的邏輯文字 offset 加穩定段落 ID；Swift text input 使用 UTF 16。邏輯文字位置與原始 XML／HTML source offset 分開，透過來源映射處理 entities 與標記。建立每個受影響段落的雙向映射、grapheme 邊界、glyph cluster 與 caret affinity；不能把 UTF 8 byte、UTF 16 code unit、Unicode scalar、grapheme 與 glyph index 當成相同單位。中英混排、emoji、組合字、RTL 及 surrogate pair 必須有驗收案例。

### 4.4 排程與取消

每個 document session 的 mutation 採單一序列執行，read snapshot 可供 renderer 讀取。Swift actor 只協調工作；阻塞 Rust 呼叫放在專用背景 executor，不佔 MainActor 或 Swift cooperative executor 的長工作時段。Core Text layout 物件遵守單一 work queue 的使用要求。

使用有容量上限的共用工作池，不能每個分頁建立無限制執行緒。互動優先於預取、縮圖與批次匯出。取消 token 以 operation ID 管理，解析、重算、轉換在分塊邊界檢查；取消儲存不能遺留半成品或宣稱成功。回呼不得在持有 document lock 時重入核心或同步等待主執行緒。

## 5 記憶體與效能規則

### 5.1 檔案與模型

- 原檔以唯讀 backing store 保存，按 ZIP part 讀取；大圖片與嵌入媒體不在開啟時全部解碼。
- memory map 只能減少適合資料的複製，不代表壓縮 ZIP entries 可直接映射成解壓資料。解壓採串流、容量限制與可驅逐快取。
- 未修改 XML 保留來源切片與未知節點；被修改部分用 patch writer 更新，其他 part 保留內容。完全未修改儲存的原檔一致性依各格式現有合約驗收。
- 稀疏活頁簿按 tile／range 載入；不配置完整最大行列網格。文字內容採經量測選定的 rope／piece table／arena 組合，不默認每個字元與儲存格都是 heap object。
- 用 style interning、共享不可變資源與緊湊 ID 降低重複配置。避免 Arc、String clone 與 Swift copy on write 在熱路徑造成隱性放大。
- Undo 採 inverse operations 加必要的結構共享，不保存每次操作的整份模型或 ZIP。checkpoint／journal 採磁碟與有上限的記憶體策略，依檔案大小調整。

Swift HostFileStore 管理 security scoped URL、Files provider 與 coordinated read write，將已授權的讀取能力／穩定來源 snapshot 提供給 Rust；不能只傳字串路徑就結束授權 scope。雲端 provider 檔案的下載、搬移與授權失效是可恢復狀態。先完成暫存輸出與驗證，再透過平台協調原子提交；權限或同步失敗不得抹去原檔。

### 5.2 排版與畫面

- Docs 只對變動段落與受影響後續區段重新排版，直到結果穩定；欄位、總頁數與全文件操作允許必要的全域 pass，但不得錯誤漏更新。
- Sheets 只重算 dirty dependency closure；volatile、spill、結構修改及跨表依賴另有明確 invalidation。
- Slides 的物件修改僅重建受影響的 layout、hit testing 與資源，不能每次拖曳重建整份 deck。
- 頁面、儲存格與圖形用可見區域及有限前後預取繪製，不能逐項建立 SwiftUI View；selection 與 caret 是輕量 overlay。
- 拖曳只更新暫時 transform，提交後產生單一交易；必要的 live preview 與公式拖曳規則另測。
- 使用者打字只同步處理必需的輸入與最近內容，背景完成其餘分頁。顯示暫時排版不得改寫文件內容。
- 圖片按照繪製尺寸 downsample，thumbnail、screen 與 print 分開快取；原始媒體保留。頁面 raster 依 zoom 與 backing scale 限制大小。
- GUI、縮圖、PDF 匯出與列印共享同一版 layout snapshot；不能透過不同文字重新換行路徑產生不同分頁。

### 5.3 快取與生命週期

建立全 App 的 MemoryBudgetCoordinator。每個 cache 同時有 bytes 上限與 item 上限，追蹤 CPU buffer、解碼圖像、glyph、PDF page、GPU texture、Undo、AI 上下文與 document sessions。CPU 與 GPU 的預算不得重複當成同一份記憶體。

背景分頁移除非必要 raster、縮圖與 layout snapshot；保留少量有上限的鄰近頁快取。memory pressure 時先停預取，再釋放可重建資料。未儲存內容必須先寫 recovery journal 才能卸載，不能為了低記憶體丟掉編輯。

iOS 進背景時提交可恢復 checkpoint、取消可重建工作並降低快取；不能依賴 App 永遠持續執行。快取預算以實機、文件種類與狀態調整，不宣稱存在所有 iPhone 都適用的固定安全上限。

### 5.4 需要量測才能採用的最佳化

SIMD、Metal renderer、跨程序隔離、額外 worker、自訂 allocator、壓縮快取與 release LTO 等，先提出對照實驗，再以瓶頸證據採用。分別量測 CPU latency、峰值 footprint、GPU／IO／energy；不能只優化單項，卻讓另外一項明顯惡化。Rust 初始化、PDFium 與 WebKit 在使用時才載入，不為預熱而常駐完整引擎。

## 6 文字輸入與渲染架構

### 6.1 不建立第二個完整編輯模型

AppKit NSTextInputClient 與 UIKit UITextInput adapter 處理 IME、marked text、selection、鍵盤與平台文字服務；可以用 TextKit 協助局部輸入與幾何，但不得讓整份 NSTextStorage 成為與 Rust 並存的第二個權威文件。初始原型要驗證相鄰段落的上下文、拼字、VoiceOver、撤銷與輸入法候選窗。

若局部文字 adapter 不能滿足系統文字服務，N08 應比較替代設計並記錄記憶體代價；不能以關閉 IME、無障礙或貼上格式來通過效能門檻。

### 6.2 Core Text 與跨平台 shaping

Rust layout 定義 TextShaper 接口。Apple adapter 以完整文字 run 與樣式批次傳入 Core Text，返回 font identity、glyph、advance、offset、cluster、bidi 與 fallback 結果。快取 key 包含文字、樣式、語言、字型檔 fingerprint、feature、variation、版本及 scale 需求，不能只用 family name。

Rust 控制 Office 的行距、換行、縮排、tab、繞圖、表格與分頁規則；Core Text 返回字形及度量，不直接替代 Word 分頁。匯出使用相同度量、字型版本與 glyph placement。字型嵌入授權、CJK 標點、RTL、ligature、vertical text、Office 私有字型與 EMF WMF 的已知能力納入矩陣。

非 Apple 平台可另做 HarfBuzz FreeType 等 adapter。此處只承諾接口可替換，不承諾不同 shaping backend 天然逐像素相同；需使用一致字型及語料驗證。

Core Text 的字型、glyph 與 layout API 是 Apple shaping adapter 的依據，layout 物件的工作序列限制須列入實作合約。[Core Text 官方文件](https://developer.apple.com/documentation/CoreText)

### 6.3 原生外觀

SwiftUI 管理首頁、工具列、側邊欄、設定與對話框；AppKit UIKit custom canvas 管理文件呈現與高頻互動。主題分為 UI chrome 與 document content；深色頁面只做顯示映射，儲存／列印／匯出不改作者顏色。沿用現有語系 key 與語意色彩概念，針對 macOS、iPad 和 iPhone 分別設計入口。

## 7 HTML PDF 與外部能力的界線

HTML preview 保留 WebKit，文件腳本是內容能力；移除 React Electron UI 不等於禁止 HTML 執行腳本。只在 HTML 功能需要時建立 WKWebView，離開後按預算回收。不得把 Docs Sheets Slides 的核心重新包回 WebView。

HTML visual inspector 允許小型、隔離、版本化的 JavaScript bridge，操作結果轉成 Rust source patch；它不是完整 JS 應用狀態。需比較 Chromium 與 WebKit 的 computed style、字型、print CSS 與截圖差異，將 HTML → DOCX 抽取改接 WebKit adapter。不能以此宣稱所有自有程式碼已完全移除 JS。

Markdown Mermaid、WaveDrom、KaTeX 等能力逐項決定原生實作、Rust library 或隔離 renderer；保留作者原始語法及可重開編輯。fallback 延遲載入且快取有上限。N26 對功能與記憶體做決策，不默默刪除圖表與公式。

PDF 預設以原生 PDFium 維持抽取、render、內容修改的權威路徑；Swift 畫布接收有限頁面 raster 與文字幾何。避免同時用 PDFKit 和 PDFium 長期保存兩份完整 PDF 狀態。PDFKit 如用於獨立預覽或系統服務，需明確生命週期及資源證據。

PDFium 原生 build 必須驗證目前 WASM 包裝實際使用的 API、自訂能力、font subset、HarfBuzz 與 pdf-lib 補充邏輯。PDFKit 註記不能替代原文內容流修改；不存在宣稱一個 library 完全覆蓋所有現有功能的捷徑。

macOS 核心預設嵌入函式庫。若 PDF 或轉換工作實測需要 crash isolation，可選擇 macOS helper，但不得把子程序變成共用核心必要條件；iOS iPadOS 必須使用程序內接口及恢復 journal。桌面外部 AI CLI integration 另列平台 adapter。

上述平台限制依 Apple 的子程序說明；PDF 內容修改依 PDFium 原生 API 驗證，HTML 預覽則依 WebKit 平台能力驗證。[Apple 子程序說明](https://developer.apple.com/forums/thread/747499)、[PDFium 編輯 API](https://pdfium.googlesource.com/pdfium/+/refs/heads/main/public/fpdf_edit.h)、[WKWebView](https://developer.apple.com/documentation/webkit/wkwebview)

## 8 可重現的效能驗收

### 8.1 基準資料與實機

N01 產出 benchmark manifest，包含文件 hash、格式、頁數／sheet／slide／非空 cell／formula／image 數、壓縮與解壓大小、字型及來源授權。至少涵蓋以下案例，並加入現有回歸 fixtures；表中的大小是待建立的測試規格，不是目前已量測的文件。

| 格式     | 一般案例                              | 壓力案例                                                |
| -------- | ------------------------------------- | ------------------------------------------------------- |
| DOCX     | 20 頁 CJK 英文，含表格圖片與頁首頁尾  | 500 頁、多節、跨頁表格、修訂、註解與大量媒體            |
| XLSX     | 5 表、100000 非空 cell、10000 formula | 稀疏 100 萬行，跨表依賴、spill、圖表與大 shared strings |
| PPTX     | 30 頁、文字圖片圖表、母片與動畫       | 300 頁、群組與自訂 geometry、大圖片與媒體               |
| PDF      | 100 頁 selectable text 與掃描混合     | 1000 頁、大 raster、subset fonts、旋轉與透明度          |
| Markdown | 10000 行，表格程式碼與圖表公式        | 長段落、大 code block、相鄰大量 diagrams                |
| HTML     | 原生單檔頁面、內嵌媒體與 scripts      | 大 DOM、複雜 CSS、字型與長頁面轉換                      |

硬體至少選定一台入門 Apple Silicon Mac、一台最低支援等級的 iPhone、一台最低支援等級的 iPad；型號、RAM、OS、解析度、refresh rate、熱狀態與供電寫入 manifest。模擬器只驗證建置與功能，不作實機記憶體／耗電結論。最低 OS 與架構由 N02 定案；macOS Intel 如需保留，單列 release 與效能 profile，不默認可放棄。

macOS 同機比較 Electron 基準版與 native release 版。iPhone iPad 沒有現有 Electron 版本，採 device budget 與 native 版本間回歸比較，不能捏造改善百分比。若歷史 QA 的 build／fixture 不同，重新建立基準，不挪用歷史 PASS。

分別量測冷啟動、暖啟動、首次可互動、完整載入／重算／資源就緒、編輯、滾動、存檔、匯出、取消、關閉與重開。首次可互動不得混同全部完成。cold run 指 App process 與 App cache 的狀態；OS filesystem cache 是否清除另標示，不宣稱開關 App 已清除系統 cache。

### 8.2 初始目標

以下是設計目標，N01 N08 N16 的實機數據取得後，依裝置 profile 固定 gate。不得因某實作較慢就自行放寬；調整必須附瓶頸、品質與資源取捨證據，更新 decision record。任何平台的 OS 記憶體終止門檻不由此表決定。

| 指標                   | 初始目標                                          | 測量定義                                                                             |
| ---------------------- | ------------------------------------------------- | ------------------------------------------------------------------------------------ |
| 打字至最近內容顯示     | p95 不超過 50 ms                                  | 含輸入 adapter、交易、可見內容 layout 與 paint；IME composition 另測                 |
| 滾動與拖曳             | 60 Hz profile 的 p95 frame time 不超過 16.7 ms    | 高刷新裝置另設 profile；記錄掉幀與最差長停頓                                         |
| 一般案例冷啟動至可互動 | p95 不超過 2 秒                                   | 固定硬體、離線內容、明確 cache 狀態；完整載入另報                                    |
| 可取消長工作           | p95 在 100 ms 內確認取消                          | 從要求取消到核心確認停止／進入安全停止點，儲存 commit 分界另列                       |
| 空 App footprint       | macOS 目標 180 MiB，iPhone iPad 目標 120 MiB 以內 | 無文件、無 AI、未建立 WebKit PDFium；是初始預算而非成果                              |
| 一般案例總峰值記憶體   | macOS 目標 400 MiB，iPhone iPad 目標 250 MiB 以內 | 各格式逐測，不把所有文件同時開啟混入一般單文件數據                                   |
| 反覆開啟關閉 30 次     | 不得持續線性成長                                  | 預熱後所有 session resource worker 釋放；allocator retention 與 bounded cache 另解釋 |
| 版本間資源回歸         | 穩定 benchmark 惡化超過 10% 需阻止自動放行並調查  | 同時看中位數、p95、噪音與功能差異，不能靠單次測量判斷                                |

大型文件不套用一般文件的固定總峰值數字。為來源、索引、dirty state、undo、render cache 各自建立實測成本模型及 profile budget；要求 cache 不隨已瀏覽頁數無界成長，壓力下能釋放可重建資料，超出能力時明確回報，不能耗盡記憶體或破壞文件。

每組正常 latency 至少 30 次獨立樣本，分開冷暖與文件狀態，報 median、p95 及樣本數。能耗用固定腳本持續執行，控制熱狀態與供電；工具不提供可比較數值時保存 Instruments trace 與操作條件，不捏造耗電百分比。不同語言的核心 microbenchmark 與端到端 App benchmark 分開報告。

macOS 舊版記錄主程序、所有 renderer、sidecar 與 GPU process 的使用情形；新版也要記錄 WebKit helper 與 GPU，不能只比較主程序 RSS。Apple physical footprint、RSS、Rust heap allocation 與 Metal texture 各自標示，避免當成可互換數字或直接相加成單一總量。release measurements 與 sanitizer runs 分開。

Apple 的記憶體指南將 footprint、峰值、背景狀態與裝置限制分開處理；本計畫的初始預算不能代替這些量測。[Apple 記憶體指南](https://developer.apple.com/documentation/xcode/reducing-your-app-s-memory-use)

### 8.3 驗證工具與產物

Rust 使用 cargo test、clippy、criterion 或等效微基準、allocation counters；Swift 使用 XCTest／Swift Testing、XCUITest 與 Instruments。對可測的 FFI helper 使用 Miri，實際混合 App 使用 Address／Thread Sanitizer；不宣稱 Miri 覆蓋 Core Text PDFium 等外部原生程式。

必要產物為 benchmark manifest、腳本、JSON CSV 結果、trace、build hash、工具鏈與字型 fingerprint。效能 gate 以固定 dedicated runner／實機結果為準，普通 CI 執行 smoke regression，不在共享 runner 上強制精細 timing。

## 9 目前程式碼與重構來源

以下相對路徑以 repository root 為基準。現有實作是移植與驗證來源，不要求逐行翻譯。

| 能力                   | 主要來源                                                                                                       |
| ---------------------- | -------------------------------------------------------------------------------------------------------------- |
| Shell 與分頁           | apps/shell/src/main/tab-manager.ts、apps/shell/src/main/index.ts、apps/shell/src/renderer/src                  |
| DOCX 格式與原檔保留    | packages/docx-engine/src/parse.ts、patch.ts、zip-splice.ts、types.ts、tests                                    |
| Docs 編輯與分頁        | apps/docs/src/renderer/editor、pagination 系列、apps/docs/tests、e2e/docs 與 word 類案例                       |
| XLSX Rust 引擎         | apps/sheets/native/xlsx-engine/src/lib.rs、main.rs、archive.rs、recalc.rs、workbook.rs、worksheet.rs           |
| XLSX 修改與結構        | packages/xlsx-gateway/src/gateway、apps/sheets/src/renderer/save-actions.ts、shared-formula-journal.ts         |
| Formula 與 Univer 補償 | apps/sheets/src/renderer/excel-formula-compat.ts、formula 系列、apps/sheets/tests/excel-formula-parity.test.ts |
| PPTX 解析與修改        | packages/pptx-engine、packages/pptx-ops、docs/slides-op-journal.md                                             |
| PPTX 排版與互動        | packages/pptx-render、apps/slides/src/renderer/konva-adapter.ts、TextEditOverlay.tsx、animation-play.ts        |
| PDF 編輯               | apps/pdf/src/main/text-edit.ts、font 系列、apps/pdf/src/renderer、apps/pdf/tests                               |
| PDF 轉換與 OCR         | packages/pdf2docx/src/ir.ts、extract、analyze、rebuild 系列、ocr-helper/vision-ocr.swift                       |
| HTML 與轉換            | apps/html/src/renderer/document、preview、packages/html2docx/src、packages/html2docx/ARCHITECTURE.md           |
| Markdown               | apps/markdown/src/renderer/editor、export、apps/markdown/tests                                                 |
| AI 與 provider         | packages/agent-core、ai-provider、ai-search、各 app 的 ai 工具與 prompts                                       |
| 資料與外部協定         | packages/project-store、packages/cli、apps/shell/src/main/mcp、control-server.ts                               |
| 無介面匯出             | docs/headless-pdf-export.md、apps/shell/src/main/headless-export.ts、各 app 的 exporter                        |
| 功能與 QA 基準         | docs/qa/skill-feature-audit.md、project-completeness-review.md、office-parity-recheck-2026-10-07.md、e2e       |

新目錄建議如下，N03 建立最小骨架；不要在規劃階段建出沒有職責的空 crate。既有 JS workspace 繼續存在，待 N35 驗收後按模組移除 runtime 依賴。

```text
native/
  Cargo.toml
  Cargo.lock
  crates/
    core/            # sessions IDs transactions undo journals
    package/         # ZIP XML source slices patch writers
    docx/ xlsx/ pptx/ # document semantics and format IO
    layout/          # pagination geometry render snapshots shaping contract
    pdf/ convert/    # PDFium wrapper and conversion IR
    text/            # Markdown HTML source models and position mappings
    agent/ store/    # portable agent state and persistence
    ffi/ cli/        # C ABI and desktop tools
  apple/
    Packages/
      EngineBridge/  # Swift wrappers and binary target
      PlatformKit/   # fonts files transport OCR render lifecycle
      EditorKit/     # input selection accessibility UI adapters
      AppUI/         # reusable UI tokens strings components
    GenOffice/       # macOS iOS iPadOS Xcode targets
  tests/fixtures/
  benchmarks/
  scripts/
docs/architecture/
  swift-rust-refactor-plan.md
  decisions/
```

外部 binary target 不提交未追蹤的大型生成物。Rust 工具鏈、Cargo.lock、Swift dependencies、PDFium revision、Xcode SDK 與 header schema 固定版本，可重現生成 XCFramework。iOS 真機與 simulator 變體不能只因同為 ARM64 就合併成同一 slice。

## 10 任務與相依關係

所有任務初始為 TODO。下表的驗收是最低條件，仍須符合第 8 節與第 11 節。狀態更新記在第 12 節。

### 10.1 基礎與風險原型

| ID  | 任務與交付                                                                         | 前置    | 驗收                                                                                                     |
| --- | ---------------------------------------------------------------------------------- | ------- | -------------------------------------------------------------------------------------------------------- |
| N01 | 建立功能矩陣、fixture manifest、原生與舊版 benchmark harness、已知缺陷清單         | 無      | 每一功能可追到來源與案例；現有基準有 build hash，不覆蓋未提交異動；保留實際失敗與略過                    |
| N02 | 固定平台、最低 OS、裝置 profile、架構／renderer／font／FFI 決策紀錄                | N01     | macOS 首發與行動後續範圍明確；Intel、離線字型、同步與桌面專屬能力有決定；核心不依賴 Apple UI             |
| N03 | 建立 Rust workspace、Apple package、macOS／iOS simulator／真機 build smoke、CI     | N02     | 同一 core 在 macOS Windows Linux 測試，Apple wrapper 在三種 Apple 環境可建置；真機實際執行基本接口       |
| N04 | C ABI、Swift wrapper、batch／buffer／callback／錯誤／取消接口及對照 microbenchmark | N03     | stale handle、重複 release、非同步 buffer、panic、取消與並發案例通過；bulk route 不用 JSON；記錄複製成本 |
| N05 | Session、穩定 ID、revision、operation、原子交易、undo redo、recovery journal       | N04     | 手動 AI CLI 相同命令；stale 操作拒絕；跨段結構復原正確；關閉／取消／crash recovery 不丟內容              |
| N06 | ZIP XML package、原始來源、resource store、patch writer、HostFileStore             | N05     | 未知 part 保留；大媒體延遲讀取；寫入原子性與外部變更衝突通過；文件授權期限覆蓋所有讀寫                   |
| N07 | MemoryBudgetCoordinator、工作池、分頁回收、memory pressure 與背景生命週期          | N05 N06 | cache 有 byte/item 上限；30 次開關無線性洩漏；低記憶體與恢復 journal 驗證；不丟未儲存狀態                |
| N08 | Core Text shaping、UTF 8 UTF 16 映射、局部輸入 adapter 與無障礙原型                | N04 N05 | macOS iPhone iPad CJK IME emoji RTL 選取與 VoiceOver；輸入與 layout snapshot 一致；無整份重複文字模型    |

### 10.2 六種編輯器與轉換

| ID  | 任務與交付                                                                               | 前置            | 驗收                                                                                                        |
| --- | ---------------------------------------------------------------------------------------- | --------------- | ----------------------------------------------------------------------------------------------------------- |
| N09 | Rust DOCX parse、模型、原始 XML anchor、save、styles rels theme、encryption              | N06             | docx-engine corpus 往返；只改指定區域；未修改檔合約與大型 ZIP case 通過；Word 開啟無修復提示                |
| N10 | Docs section、paragraph、table、float、notes、header footer 分頁與 display list          | N08 N09         | CJK 多欄跨頁表格與圖片繞排對照；局部修改更新到穩定位置；GUI export 同一佈局                                 |
| N11 | 原生 Docs 編輯、游標選取、剪貼簿、快捷鍵、尋找取代、undo 分組                            | N05 N10         | 既有 word-keys 與 paste corpus；IME 候選與跨段操作；操作復原儲存重開完整流程                                |
| N12 | Docs 註解修訂、樣式、表格進階、欄位目錄、公式圖表與完整 ribbon                           | N11             | 按 N01 功能矩陣逐項保留顯示編輯；對既有部分功能明確標示；Office 與列印對照                                  |
| N13 | 現有 XLSX Rust library 化，移除對 main.rs stdio 子程序的核心依賴                         | N04 N06         | 程序內 open range close convert；真機可用；handle cache cancel 不依賴桌面路徑與子程序                       |
| N14 | 將 xlsx-gateway TS 寫入、結構、資源、shared formula 與 rels 修改遷入 Rust                | N13             | cell style row col sheet chart validation CF pivot 等寫入按矩陣；未修改部分保留；XLS XLSM CSV 路徑另驗證    |
| N15 | 單一權威 formula engine，先驗證 IronCalc，加完語意與缺少函式                             | N13             | 用 Excel 已驗證 corpus 與現有補償案例；spill、日期空白強制轉型循環與跨表參照；不再常駐 Univer 第二模型      |
| N16 | 原生稀疏 grid、tile rendering、selection、cell edit、freeze、fill、formula bar           | N08 N13 N15     | 100 萬行虛擬化；可見範圍批次 FFI；觸控鍵盤貼上與公式重定位；資料不按整個網格配置                            |
| N17 | Sheets filter sort tables charts pivot slicers CF validation tracing 與完整 ribbon       | N14 N16         | 每項操作 undo redo save reopen；結構改動後 formula 正確；AI 手動相同交易；Excel 無修復提示                  |
| N18 | Rust PPTX parse patch、masters layouts themes assets、element identity 與 ops            | N06             | pptx-engine pptx-ops corpus；未知 part 保留；copy group ungroup reorder IDs；PowerPoint 無修復提示          |
| N19 | Rust PPTX render tree、text autofit、geometry、charts SmartArt effects 與 Apple renderer | N08 N18         | 原有 render corpus；群組座標、私有字型、縮圖與 PDF 一致；未支援的 effect 有明確保留策略                     |
| N20 | 原生 Slides 編輯、hit testing、拖曳、對齊、裁切、文字與表格編輯                          | N05 N19         | pointer touch keyboard；preview 只改 transform，commit 可 undo；儲存重開不漂移                              |
| N21 | 動畫 timeline、放映、notes、presenter view、多螢幕與媒體                                 | N20             | 用現有 animation 與 show 案例；真雙螢幕；iPad 外接輸出按平台能力測；idle 不常駐高頻 timer                   |
| N22 | PDFium 原生 build、API coverage、Rust wrapper、lazy page text raster resource            | N03 N04         | macOS iOS simulator 與真機；實際使用 API 可用；大 PDF 按頁；無重複完整 PDF model                            |
| N23 | PDF 原文編輯、字型 fallback subset、插入、annotations、搜尋頁面操作與加密                | N05 N08 N22     | 現有 text-edit 與 font corpus；內容流修改再抽取；選取與 render 一致；PDF reader 重開與 undo 測試            |
| N24 | Rust conversion IR、PDF 抽取分析重建、Vision OCR、DOCX XLSX PPTX 轉換                    | N09 N14 N18 N22 | goldens CJK RTL text scan 混合；多欄表格圖片；保存為可編輯結構；取消與背景恢復測試                          |
| N25 | Rust Markdown source AST patch 與原生區塊／source 編輯、outline search export            | N05 N08 N09     | 原始語法保留、最小 diff、IME 表格 code lists 圖片；save reopen 與 DOCX 匯出                                 |
| N26 | Markdown math Mermaid WaveDrom 等完整能力與有預算的 renderer                             | N25             | 支援清單明確；原文與可編輯性保留；每種 renderer 尺寸與 30 次切換洩漏測試                                    |
| N27 | Rust HTML source patches、原生 source editor、WKWebView inspector 與 HTML 轉 DOCX        | N05 N06 N09     | stale patch 拒絕；選取雙擊文字 move restyle；WebKit 與原 Chromium 對照；DOCX pipeline 與 JS bridge 生命週期 |

N15 不把 IronCalc 直接視為 Univer 等價替代。先產出函式、錯誤值、參照、locale、array 與 recalc coverage；若有缺口，優先在同一 Rust engine 補齊，或在同一權威模型內整合執行器。不得讓 unsupported formula 默默變成 0、空白或永遠使用 cached value。

N09 N18 的未知節點保留、N23 的內容修改與 N24 的轉換保真是不同合約。每種輸入／功能明確記錄可編輯性與 lossy 轉換；不能用 raster 整頁替代所有可編輯文件驗收。

N24 同時移植 file-parse 的 AI 附件抽取與舊 DOC PPT XLS 格式路徑。純抽取、轉成新格式與原格式可編輯是不同能力；按現況逐項對齊。OLE／CFB、巨集保留與 Office 加密另設 fixtures；保留巨集不代表執行巨集。

### 10.3 AI 桌面整合與跨裝置

| ID  | 任務與交付                                                                                        | 前置                                                    | 驗收                                                                                                          |
| --- | ------------------------------------------------------------------------------------------------- | ------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------- |
| N28 | Rust agent loop 與 protocol parser、Swift transport／Keychain、media search 工具與六 app 工具接口 | N05 N04                                                 | provider 矩陣、SSE 截斷重試取消與工具錯誤；上下文有上限；AI edit 全批回復；桌面外部 CLI 走 platform adapter   |
| N29 | project/chat/settings 資料遷移、版本、journal 與可持久化 timeline                                 | N05 N06                                                 | 既有 JSON JSONL 容錯匯入；原資料備份可回復；ID 與文件關聯保留；crash 與並行保存不損壞                         |
| N30 | Rust CLI MCP、GUI 定位控制、skills integration 與原有命令協定                                     | N09 N14 N18 N22 N24 N25 N27 N28 N29 N31 N38             | 命令名 schema JSON exit code 相容；headless 與 open document 工具皆測；Mac control token 與 selection 含版本  |
| N31 | 共用 screen thumbnail PDF print renderer 與無可見視窗 Apple exporter                              | N12 N17 N21 N23 N26 N27                                 | GUI CLI 同文件同 layout；export 不啟動 Electron；Word 分頁／sheet print area／slides notes handouts／實體列印 |
| N32 | iPad 原生 UI、Files、keyboard pointer touch、multiwindow 與生命週期                               | N07 N12 N17 N21 N23 N24 N26 N27 N28 N29 N31             | 真機全部格式開編存；低記憶體切換 background recovery；無子程序；各入口按能力矩陣驗收                          |
| N33 | iPhone 原生 UI、small screen、selection、software keyboard 與可達功能                             | N32                                                     | 不靠縮小 Mac ribbon；六格式與 AI 實機流程；小螢幕無障礙；載入取消與背景恢復                                   |
| N34 | 跨裝置文件與 chat 同步、版本衝突、離線恢復                                                        | N29 N32 N33                                             | 雙裝置離線編輯與重連、不同版本文件與 chat；不把 OOXML 二進位最後覆蓋視為合併；保留衝突副本                    |
| N36 | 每階段效能與資源最佳化，固定 device profile 與回歸 gate                                           | N07 及該階段功能                                        | 第 8 節的量測腳本 trace 與 gate；無混淆主程序／總程序；品質一致且 cache 有上限                                |
| N37 | 功能全矩陣、Office 互通、fidelity、golden UI 生命週期與平台差異驗收                               | N12 N17 N21 N23 N24 N26 N27 N30 N31                     | 逐功能 evidence；Office 開啟檢查、fixtures 與已知缺口；不能只搬移舊測試的 PASS 文字                           |
| N38 | 完整 macOS Shell、首頁與分頁、最近檔案、設定與原生選單、主題語系、視窗和快捷鍵                    | N05 N07 N29                                             | 不同格式分頁／未存關閉／外部改名與修改；設定遷移；原生主題與無障礙；單文件 dirty 狀態一致                     |
| N35 | macOS 原生發行、簽章公證、更新與功能切換，移除完成模組的 JS runtime                               | N12 N17 N21 N23 N24 N26 N27 N28 N29 N30 N31 N36 N37 N38 | macOS 驗收報告、升級與回退、檔案關聯、離線啟動、合法依賴清單；六 app 達到已定功能矩陣                         |

N36 是每階段 gate，首次執行不依賴後續全部功能；N37 可逐模組累積。N35 只要求 macOS 範圍完成，不依賴 N32 N33 N34；行動發行各自再次執行對應 N36 N37。桌面 release 與行動 release 不能共用一個未區分平台的 PASS。

### 10.4 建議推進順序

1. N01 N02 N03 N04，固定來源與接口，驗證 Rust 真機 build。
2. N05 N06 N07 N08，先證明交易、來源保留、輸入與記憶體生命週期。
3. 同步檢查三個高風險原型：N09 N10 的分頁、N13 N15 N16 的公式與 grid、N18 N19 的文字與圖形排版。原型結果決定成熟實作，不能略過。
4. 交付第一條完整原生流程：Shell 加 Markdown 或 PDF viewer 的 open edit save undo close reopen。PDF viewer 僅作第一個垂直流程，不算 N23 完成。
5. 逐模組完成六種編輯器，同時推進 N28 N29 N38；每個模組接入共用命令、匯出與效能驗收。
6. N30 N31 N36 N37 N35，完成 Mac 整合與發行；再完成 N32 N33 N34 及行動 gate。

## 11 功能與發行驗收規則

### 11.1 每個功能的完成條件

- 同一來源文件開啟後內容與佈局符合已定合約，原本無法理解的資訊不因儲存被抹除。
- 手動操作與 AI 工具採同一 command；完成後 undo redo save close reopen 仍一致。
- session 切換、取消、外部檔案修改、memory pressure、App 終止與恢復不錯用版本、不丟未儲存內容。
- 原生 UI keyboard mouse touch、IME、VoiceOver 與高 DPI 按平台矩陣驗證。
- 列印、PDF、圖片、格式轉換與 headless 共享 layout 與 resource 規則。
- 所有失敗、缺少函式、保留但不能編輯與有損轉換明確呈現；禁止假成功或靜默降級。

### 11.2 舊版與 Office 的對照

舊 JS 版用於現有行為與回歸對照；Word Excel PowerPoint 用於已確認的格式與語意。既有 bug 不能只因舊版相同就視為新版本正確。語系、字型、Office 版本與測試自動化限制記錄在 evidence。

OOXML fidelity 分成：未修改整檔 bytes、未修改 part 解壓後 bytes、XML 片段 bytes、格式／語意與 visual layout。各格式按既有可成立的合約驗證，不能把 ZIP metadata／compression 不同直接當作內容損壞，也不能因能開檔就忽略未知 XML 被改寫。

視覺比較固定字型、版面、scale 與來源；文字 glyph raster 差異與幾何差異分別判斷。頁數、換行、表格 bounds、object 位置、formula 結果與來源 anchor 應有語意 assertions，不能只靠 pixel snapshot。

### 11.3 切換與回退

各格式 native 能力先使用明確 feature flag 或內部測試入口。功能達到矩陣後切換，舊 JS 產品與 fixtures 仍可作對照；同檔禁止兩個引擎同時修改。新 store 遷移先備份與記錄版本，升級失敗可回復。到最終 release 才移除已替換 runtime，保留歷史測試資料與授權 NOTICE。

新原生功能只執行相關 Rust Swift UI 與效能檢查；如共用來源或舊 JS 程式碼也有修改，額外執行 CONTRIBUTING.md 要求的原有檢查。純規劃文件修改不代表所有 App 測試需要重跑，文件的驗證結果也不能當成產品驗收。

## 12 任務追蹤與後續交接

初始狀態：N01 至 N38 全部 TODO；本次交付為實作計畫，不代表 native 骨架或重構程式碼已完成。N35 N36 N37 N38 的編號為追蹤 ID，執行順序依第 10 節，不按數字強行排序。

每次完成工作，追加以下紀錄；未完成的工作不能改成 DONE。路徑均使用 repository 相對路徑，方便換裝置與 checkout。

```text
日期與 commit／工作目錄狀態：
任務 ID 與狀態：
完成範圍與產物路徑：
測試命令、實機與 fixture：
功能驗收結果與未覆蓋範圍：
效能／footprint／複製／cache 證據：
已知問題與架構決策紀錄：
下一個可執行任務與前置條件：
```

未來實作提示可直接使用：

> 請先讀取 docs/architecture/swift-rust-refactor-plan.md 與適用的儲存庫規則，檢查現有未提交異動，依相依關係選擇下一項 TODO 任務。以 Swift 原生 UI 與 Rust 共用核心實作，遵守單一文件狀態、批次 FFI、容量上限、檔案保真及驗收規則。先完成該任務的必要測試與實機／效能證據，再更新第 12 節。若 N01 尚未完成，先做 N01，不直接大量搬移 UI 或建立全部空 crate。

## 13 官方技術依據

- [Rust iOS 與 iPadOS 編譯目標](https://doc.rust-lang.org/rustc/platform-support/apple-ios.html)：Apple 真機與 simulator 支援；支援 target 不等於本專案所有依賴已驗證。
- [Apple XCFramework 建置方式](https://developer.apple.com/documentation/xcode/creating-a-multi-platform-binary-framework-bundle)：不同平台與變體的原生函式庫包裝。
- [Apple 對 iOS 子程序的限制說明](https://developer.apple.com/forums/thread/747499)：現有 sidecar 模式不能直接作為行動 App 架構。
- [Core Text](https://developer.apple.com/documentation/CoreText)：字形佈局與字型度量，以及 layout 物件的 queue 使用限制。
- [SwiftUI 與 AppKit 整合](https://developer.apple.com/documentation/swiftui/nsviewrepresentable)：原生 canvas 可嵌入 SwiftUI 介面。
- [Apple 記憶體量測與降低使用量](https://developer.apple.com/documentation/xcode/reducing-your-app-s-memory-use)：footprint 與裝置相關限制。
- [Instruments 效能分析](https://developer.apple.com/tutorials/instruments)：互動阻塞與資源量測。
- [Rust FFI 與 unwinding](https://doc.rust-lang.org/stable/nomicon/ffi.html)：跨語言接口、所有權與 panic 邊界。
- [PDFium 原生編輯 API](https://pdfium.googlesource.com/pdfium/+/refs/heads/main/public/fpdf_edit.h)：內容修改與 native wrapper 能力驗證。
- [WKWebView](https://developer.apple.com/documentation/webkit/wkwebview)：HTML 內容與專用 browser adapter。
- [Apple 文件衝突處理](https://developer.apple.com/documentation/uikit/uidocument)：文件同步後仍須處理版本衝突。
