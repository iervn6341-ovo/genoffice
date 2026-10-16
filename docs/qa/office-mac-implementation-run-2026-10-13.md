# Microsoft 365 Mac 工具細項：實作與獨立複驗

日期：2026-10-13（Asia/Taipei）。起始產品 HEAD：`6578a439`。承接 2026-10-12 的三份細項稽核；歷史證據保留不改寫。

使用者明確授權依修正優先順序開始實作，三位實作者使用 GPT-6 Sol／High，完成後由 GPT-6 Luna／Extra（`xhigh`）檢查員獨立複驗，root 負責監督與整合。

## 排程與 ownership

最多同時四個 active agents（包含 root）。實作、檢查分階段；Electron／原生 Office 共用桌面與 build 互斥。任何 build 由 root 統一執行，檢查員在相同產物上逐一進行 UI 測試。

| 工作     | 負責人／設定                                                      | 修改範圍                                                                           |
| -------- | ----------------------------------------------------------------- | ---------------------------------------------------------------------------------- |
| Docs     | `docs_implementer`／GPT-6 Sol High                                | `apps/docs`、`packages/docx-engine`；共享 `RibbonFoldGroup.tsx` 及其測試單一負責人 |
| Sheets   | `sheets_implementer`／GPT-6 Sol High                              | `apps/sheets`、`packages/xlsx-gateway`                                             |
| Slides   | `slides_implementer`／GPT-6 Sol High                              | `apps/slides`、`packages/pptx-engine`；`pptx-ops` 頁尾 op 接線及測試               |
| 獨立複驗 | docs_verifier、sheets_verifier、slides_verifier／GPT-6 Luna xhigh | 核對修改及驗收案例、建立本輪測試證據；修正意見交還實作者                           |
| 整合     | root                                                              | ownership、來源審查、整合 build、測試排程、狀態紀錄                                |

## 第一批驗收範圍

按原修正優先順序先處理頁首頁尾與確定的資料正確性缺陷。之後的方程式／符號及進階格式工具維持後續優先項目；不將任何未實作選項當作已完成。

- Docs：避免單一純文字操作摧毀既有 rich header/footer；節與 variant 導覽、連結和欄位插入需對接真實資料模型。保留 first/odd/even 與其他未改節。共享折疊群組內巢狀 picker 可持續操作，完成命令、外部點擊和 Escape 仍關閉。
- Sheets：現行 Univer JSON list 轉合法 Excel literal，重開轉回 JSON。保留範圍來源與 legacy CSV，提示／錯誤／範圍不變。不能用改 test seed 遮掩問題；含逗號／引號／換行、空字串或超過 inline 限制的清單不能 silently 改值。頁首頁尾 UI 開放有保存模型支持的 first/even、欄位及文字格式。
- Slides：固定／自動模式和日期格式回填；套用目前頁／全部及標題頁排除。保留既有 placeholder 幾何、格式與動畫。經 main／op／engine 真正保存範圍與語意；備忘稿／講義僅在安全保存支援完成後提供可用 UI。

## 檢查要求

先由實作者跑針對改動的單元／資料往返／型別檢查。完成後檢查員獨立審查及實際 UI → 保存 → 關閉重開 → 再編輯驗證；共享選單需窄／寬視窗與跨應用回歸。不把 engine pass 等同全部 Office parity，也不把 locator 錯誤當產品 bug。

新 UI 使用語意色彩 tokens；新增字串完成各 locale shards。main／preload 改動需包含 shell 重建。保留既有 untracked 文件、原始使用者檔案與歷史證據，不提交／推送。

## 當前狀態

**第一批已實作、整合並完成獨立複驗。** 三位 Sol High 實作者與三位 Luna xhigh 檢查員分階段工作，由 root 排程與整合。此次完成頁首頁尾、資料驗證保存及折疊群組互動；未宣稱與 Microsoft 365 Mac 全部功能相同。

| 應用   | 本批交付                                                                                                                                                 | 已通過的桌面情境                                                                                                                                              |
| ------ | -------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Docs   | 指定頁首頁尾編輯目標節及 first/default/even 變體；連結前節；多行文字與 PAGE/NUMPAGES；保留表格、圖片關係及未改段落；跨存檔 Undo；巢狀符號／方程式 picker | 3 項：1280px 折疊 Symbols；多節 rich header 保存／重開及 link；QAT Undo/Redo、保存後 Undo 再保存時原 rId／XML bytes 恢復                                      |
| Sheets | UI JSON list 與 Excel inline literal 轉換、空白保留、無法無損匯出時明確報錯；首頁／偶數頁及格式／欄位；保存、重開及多頁 PDF；停用再啟用時保留內容        | 4 項：真實 UI 清單建立／保存／重開編輯；HF 保存／重開／PDF／重新啟用；窄版 Orientation MenuSelect 與 font palette；逗號選項保存失敗時原 bytes／live rule 保留 |
| Slides | 固定／自動日期及格式回填；套用目前／全部；標題頁排除；保留 placeholder 幾何、樣式、動畫；繼承母片欄位讀取、隱藏及關係複製                                | 1 項綜合情境：繼承欄位、目前頁、固定日期重開、自動日期預覽、全部套用、兩張標題頁排除、hyperlink relationship 及第二次重開；日期 radio 實測 14×14px            |

8 項為分次序列執行後各情境的最終通過結果，並非單次完整 test command 的總數。Docs 的節切換控制改變頁首頁尾編輯目標，與本文 caret 分開；不把它列為已完成 Word 式本文頁面導覽。Sheets 重新開啟後檢查實際規則編輯器，未測儲存格 dropdown 選取。此次未在原生 Word／Excel／PowerPoint 重跑新修改；原生差異仍參照歷史稽核。

## 修正與驗證證據

root 審查修正了 Docs 共用 rId 污染、self-closing sectPr、待存 link／文字衝突與跨存檔歷史。Luna 來源審查找出 Slides 繼承母片欄位無法正確讀取／關閉，以及 Sheets 停用 first/even 後回填失去文字。實際 UI 又找出 Sheets 把選項顏色 `formula2` 當公式匯出、PDF `&F` 顯示 UUID snapshot 名稱。以上已交回實作者修正並通過最終驗收。root 視覺檢查另修正 Slides radio 過大。中間 profile／locator／variant／XML escape／非精確文字與非同步保存判斷的探測錯誤保留原始結果，經修正後重新測試；不當成產品缺陷或通過證據。

- `build:all` 成功；後續 Sheets／Slides／shell 重建成功；最後檔名修正再重建 Sheets／shell 成功，包含 main／preload。[建置與測試證據](office-mac-evidence-2026-10-13/README.md)。後續 Prettier 調整只有格式變更。
- DOCX 引擎完整套件：142 個檔案通過，1420 項通過、1 項略過。PPTX 引擎完整套件：103 個檔案通過、1 個檔案略過，1065 項通過、1 項略過。兩個既有 HF 測試契約改為檢查原 part／動畫保留。
- Docs／shared UI 聚焦測試 91 項通過；Sheets 最後 5 個聚焦套件 124 項通過；Luna 獨立清單／色彩 metadata 聚焦測試 37 項通過（與實作者測試有重疊，不加總）。Slides HF 引擎／op／日期測試通過。相關 TypeScript、shell／shared UI 型別檢查通過。
- 修改範圍的 Prettier、lint 無 error、`git diff --check` 及 theme colors 通過；既有 hook lint warnings 保留。全域 `format:check` 亦掃到大量既有 untracked probes，失敗未以改寫使用者檔案繞過。`check:english-comments` 因既有 tracked `.claude/skills/genoffice-qa/SKILL.md` 在工作樹缺失而未能執行；未建立假檔案或安裝依賴。新增程式註解以英文撰寫。

### 個別交付與獨立複驗

- [Docs 實作](docs-office-mac-implementation-2026-10-13.md)／[Luna 複驗](docs-office-mac-verification-2026-10-13.md)。最終 rich run：[證據](office-mac-evidence-2026-10-13/docs-verification/sections-rich-header-2026-10-12T17-07-42-134Z/)。各 run 的實際時間及檔案以複驗報告記錄為準。
- [Sheets 實作](sheets-office-mac-implementation-2026-10-13.md)／[Luna 複驗](sheets-office-mac-verification-2026-10-13.md)。[最終 HTML report](office-mac-evidence-2026-10-13/sheets/accepted-html-report/index.html) 包含截圖／XML；[PDF、活頁簿與 inline JSON](office-mac-evidence-2026-10-13/sheets/accepted-attachments/)。
- [Slides 實作](slides-office-mac-implementation-2026-10-13.md)／[Luna 複驗](slides-office-mac-verification-2026-10-13.md)。[最終視覺與保存證據](office-mac-evidence-2026-10-13/slides/runs/2026-10-12T17-06-37-426Z-JLZ5pQ/)。

## 後續優先工作

1. 方程式／符號：Docs 的 11 類結構與進階符號、Slides 缺少的 Symbol 入口及真正結構數學物件、Sheets 方程式語意。此次僅修正折疊互動，未增加這些數學功能。共用數學模組需先處理套件依賴方向，再實作、保存重開及原生往返驗收。
2. 未完成的頁首頁尾細項：Slides 備忘稿／講義及對話框起始編號；Docs DATE 等欄位與新圖片編輯、本文／HF 導覽對齊及畫布雙擊編輯歷史；Sheets 新 HF 圖片及格式碼 fidelity。Slides 對目前頁的標題排除以 local mask 保存，deck checkbox 只反映全部標題 layout 排除。
3. 其餘進階工具：Format Cells、Text to Columns、Notes／threaded Comments 與稽核列出的深層控制項。仍需原生 Microsoft 365 Mac 實機確認及匯出再開啟，不以新增外觀控件替代實際功能。

原始使用者文件、既有 untracked 資料及 2026-10-12 歷史證據保持原樣。未提交或推送。
