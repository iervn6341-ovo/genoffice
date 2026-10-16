# Microsoft 365 Mac 工具列細項持續驗收（2026-10-14）

本輪三位 GPT-6 Sol／High 實作者在 Docs、Sheets、Slides 平行處理來源，Electron 前景測試依序執行以避免 macOS 共用焦點互搶。三位 GPT-6 Luna／Extra 檢查員分別獨立複驗。依據為實際 Word、Excel、PowerPoint Mac 與其原生檔案，不將局部切片記為整項完成。

| 項目 | 本輪結果 | 驗收與證據 |
| --- | --- | --- |
| Docs D6 | 註腳按實際頁面重啟編號，設定可只套用游標所在節；跨節選取停用單節作用範圍。修正存檔重讀時衍生顯示號讓 Undo 歷史被清空的問題，只在正文等價比對中忽略該顯示號，仍比較原生 reference 的種類、ID、格式。 | Docs+Shell Electron `e2e/docs-note-page-scope-verification.spec.ts` 1/1 通過，涵蓋頁尾與正文標記、Save→Undo/Redo→Save、末節 trailing `sectPr`、新程序重開與再編輯。舊 `docs-note-section-preview-verification.spec.ts` 2/2 回歸通過。聚焦 Docs 10/10、DOCX 引擎 11/11、typecheck、Luna 簽核通過；輸出 `/private/tmp/genoffice-docs-note-page-scope-fixed2-results`。 |
| Sheets S4 | 依 Excel Mac 真實 2×4 單／雙、一般／會計底線夾檔與截圖，為純文字、未旋轉儲存格加畫跨欄寬的會計線；原生 `<u>` 格式維持不變。 | Sheets+Shell Electron `e2e/sheets-accounting-underline-verification.spec.ts` 1/1 通過：一般線約 30–50 px、會計線約 151.5 px；Format Cells、Save→Undo→Save→Redo→Save、原生 XLSX 與重開。實作者聚焦 27/27，Luna 26/26；來源及截圖在 `office-mac-field-evidence-2026-10-13/sheets-accounting-underline/`，結果 `/private/tmp/genoffice-sheets-accounting-underline-results3`。 |
| Slides L5 | PowerPoint Mac 圖表 Quick Layout 11 稀疏修改：移除標題、標記自動標題刪除、兩軸主要刻度改向外，並保留版面所需圖例與欄距；不重建圖表資料。 | Slides+Shell Electron `e2e/slides-chart-native-format-verification.spec.ts` 全部 6/6 通過，涵蓋 Layout 11 的 Undo／Redo、Save As、新程序重開、再編輯及關聯／樣式／內嵌活頁簿保留。聚焦 19/19、typecheck、Luna 簽核通過；輸出 `/private/tmp/genoffice-slides-layout11-final-results`。 |

Sheets 的 Luna 複驗只簽核純文字、未旋轉儲存格：富文字及旋轉儲存格目前走另一條 Univer 文件渲染路徑，會計線仍可能顯示成一般底線。實作者正在補安全防呆及使用者可見說明；這兩類尚未達到畫布對等。Docs 節末章末註仍需新的節末重排模型，自訂符號、轉換及導覽待做。Slides 的其他八個 Mac Quick Layout 涉及軸標題、資料表、series labels 或次要格線等原生物件；更完整的圖表／圖片工具也待做。D6、S4、L5 繼續維持 `in_progress`。

## 本輪後續切片

Word Mac 已透過原生介面建立自訂 ♠ 註腳與 ♥ 章節附註，存成 `office-mac-field-evidence-2026-10-13/docs-native-notes/word-mac-custom-marks-source.docx`。`word/document.xml` 的兩種 reference 都帶 `w:customMarkFollows="1"` 並在同個 run 後接 Symbol 字型 `w:sym`（F0AA / F0A9）；note body 首段也以相同 `w:sym` 作標記。Docs 自訂標記依此真檔實作中。

Sheets S4 已對富文字、旋轉儲存格補使用者可見的會計底線限制與停用選項，避免在同一個格式對話框將新會計底線和旋轉混用；未修改的原生底線格式保存。Sheets+Shell 建置、聚焦 28/28 通過；Excel Mac 真檔 Electron `e2e/sheets-accounting-underline-verification.spec.ts` 1/1 通過，包含普通儲存格畫布像素、Undo/Redo、存檔重開，以及 E1 富文字/F1 旋轉的防呆。結果在 `/private/tmp/genoffice-sheets-accounting-underline-guard-results3`。富文字和旋轉儲存格的完整會計線繪製尚未完成，S4 保持 `in_progress`；Luna 複驗進行中。

Slides L5 再增 PowerPoint Mac Quick Layout 9。根據真 PPTX Layout 1→9，僅新增類別軸/數值軸自動 `c:title` 原生子樹，並將數值軸主要刻度由 `none` 改 `out`，保留原 chart/relations/style/colors/embedded workbook；匯入不相容軸時拒絕此版型。Slides+Shell 建置成功，圖表聚焦 126/126、型別/lint 通過；完整 Electron `e2e/slides-chart-native-format-verification.spec.ts` 7/7 通過（37.5s），新增案例涵蓋 Undo/Redo、Save As、新程序重開及右側數值標籤後再編輯。結果 `/private/tmp/genoffice-slides-layout9-final-results`；Luna 複驗待排程。L5 保持 `in_progress`，剩餘 Mac Layout 2/4/5/6/7/8/10 與其他圖表／圖片工具仍待實作。

Sheets GPT-6 Luna／Extra 已獨立審視會計底線防呆及真檔 E2E，聚焦 28/28，沒有阻擋問題。`selectionProtection` 未知時底線選單保守停用，但不顯示特定提示，列為非阻擋 UX 限制。

Slides GPT-6 Luna／Extra 已簽核 Layout 9：重開後最終 chart XML 與 PowerPoint Mac Layout 9 原生快照逐位元相同，chart 關聯、style/colors、內嵌 XLSX 也未變；聚焦 21/21 及 engine/ops/Slides typecheck 通過。若匯入軸已有空 `<c:title/>`，目前視為已有標題，不會替換為原生自動標題子樹，此邊界情況未覆蓋，但不阻擋真樣本切片。

為下一段 Sheets S4 完整畫布呈現，已在 Microsoft Excel Mac 開啟 `office-mac-field-evidence-2026-10-13/sheets-accounting-underline/excel-mac-rich-rotated-source.xlsx` 並保存 `excel-mac-rich-rotated-source.png`。E1 富文字的會計線水平幾乎跨滿儲存格，F1 45° 旋轉文字的線沿文字基線斜向；因此不能把未旋轉的水平畫法直接套用到旋轉儲存格。

Docs D6 自訂符號來源已依 Word Mac ♠／♥ fixture 接通正文 reference 的 `customMarkFollows + w:sym`、note body 首段 `w:sym`、畫布顯示、編輯與原生寫回；不可安全拆解的混合引用 run 保持受保護。Docs+Shell 建置及 Electron `e2e/docs-note-custom-marks-verification.spec.ts` 2/2 通過（1.2m）：新建、跨 Save Undo/Redo、重開，以及原生 Word Mac 真檔匯入、改 note text 後保存結構。GPT-6 Luna／Extra 獨立聚焦 engine 39/39、Docs 26/26、兩側 typecheck/目標 lint 通過，無產品 blocker。現有 E2E 未直接斷言 Page Preview 的自訂標記，正補驗收斷言；因此暫不宣稱預覽已經實機驗證。結果 `/private/tmp/genoffice-docs-note-custom-marks-results`。

Slides L5 Quick Layout 10 已完成：第三 series 原生數值標籤、gapWidth 75、overlap 40；其他 chart/關聯/樣式/內嵌活頁簿維持原生。Slides+Shell 建置、完整 Electron 圖表套件 8/8（40.1s）通過，含 Layout10 Edit→Undo/Redo→SaveAs→fresh reopen→再編輯；Luna／Extra 獨立核對真 Mac XML diff、聚焦 24/24 與 engine/ops/Slides typecheck，無 blocker。結果 `/private/tmp/genoffice-slides-layout10-final-results`。L5 仍有 Layout 2/4/5/6/7/8 及廣泛圖片／圖表工具未完成。

Sheets S4 富文字/旋轉會計底線的第二次實機畫布 probe 未達 Excel Mac 對等：在 Univer 私有 document-line extension 擴線後，E1 最長可見筆畫只有 4.5 px、F1 只有 3 px；原生 Excel 對照分別約 119 px 與斜向 75 px。線條受 document/cell clip 與旋轉路徑裁切，不能把局部 hook 視為完成。未達標的實驗 patch 已撤回，保留先前 Luna 簽核的停用選項與原生 `<u>` 保存；S4 仍為 `in_progress`，需重新設計安全的 overlay/渲染層並另做真畫面驗收。

Docs Page Preview 補驗收後重新執行 `docs-note-custom-marks-verification.spec.ts`，Electron 2/2 通過（約 1.1m），現在含畫布與 Page Preview 的 ♠/♥ 註腳、章節附註標記；結果 `/private/tmp/genoffice-docs-note-custom-marks-preview-results`。下一個 D6 轉換切片已用 Word Mac 真 UI 建立 before、單筆 footnote→endnote 與單筆 endnote→footnote 三份原生 DOCX，存於 `office-mac-field-evidence-2026-10-13/docs-native-notes/conversion/`。兩方向都會在目標種類 ID 衝突時重編號成 1/2，轉入 body 的 pStyle 在 `ae` 與 `af1` 間切換，自訂 Symbol 標記及兩邊 separator 保留。實作者依三檔設計中。

Sheets S4 撤回未達標實驗後，Sheets+Shell 建置、聚焦 25/25、Electron 原生夾檔回歸 1/1 及 Luna／Extra 68/68 獨立複驗通過；最終 XLSX 保留 E1 rich runs/粗體、F1 45°、原生會計 `<u>`。結果 `/private/tmp/genoffice-sheets-accounting-rollback-results`。該兩類的畫布呈現仍需不同渲染架構，防呆維持開啟。

Slides L5 Quick Layout 7 已依 PowerPoint Mac Layout9→7 真 XML 差異，移除安全 chart title/設 autoTitleDeleted、將 gapWidth 改 300，並新增 value-axis 原生 minor gridlines；自動軸標題保留，其他資料及關聯稀疏保留。Slides+Shell 建置、聚焦 27/27 與 Luna／Extra 獨立複驗無 blocker。完整 Electron 9 案中 Layout7 與其他 7 案通過；既有 Legend None 案在 `e2e/helpers.ts:136` 啟動視窗時發生 promise garbage-collected，未進入產品操作，單獨重跑 1/1 通過。因此 9 案跨兩次執行全部通過，不能表述為單次 9/9。結果 `/private/tmp/genoffice-slides-layout7-final-results`、`/private/tmp/genoffice-slides-layout7-legend-rerun-results`。剩餘原生版型 2/4/5/6/8。

Sheets S5 原生對照已透過 Excel Mac UI 建立 `office-mac-field-evidence-2026-10-13/sheets-text-to-columns/` 的 before/after XLSX、Step2/覆寫警示/結果 PNG 與 manifest。來源 A1=`"A, B",003,4`、A2=`"C,D",005,6`，B1 有 `=10+5` 公式；Delimited/Comma/雙引號 qualifier、第二欄 Text、目的 `$A$1` 後，Excel 在 Finish 先提示「There's already data here. Do you want to replace it?」，接受後 B1 變文字 `003`，A/C 欄正確拆分。GenOffice S5 將採更保守的既有資料/公式覆寫預檢，顯示原因並禁止執行；來源實作中，尚未完成驗收。

Slides L5 Quick Layout 8 已依 PowerPoint Mac Layout 9→8 真 XML 差異接入：`gapWidth` 150→0，移除已知的數值軸主要格線與圖例，其他資料、關聯、樣式及內嵌活頁簿稀疏保留；匯入結構含未知 extension/relation 或不相容格線／圖例時拒絕套用。Slides+Shell 建置、聚焦 30/30、型別及 lint 通過；完整 Electron 圖表套件單次 10/10 通過（約 1.1 分鐘），第 10 案涵蓋 UI→Undo/Redo→Save As→新程序重開→再編輯圖例→再次套用及保存，結果 `/private/tmp/genoffice-slides-layout8-final-results`。Luna／Extra 獨立核對真 Mac XML 與保護條件後簽核無 blocker；剩餘原生 Quick Layout 2/4/5/6 及廣泛圖表／圖片工具，L5 維持 `in_progress`。

Docs D6 下一切片已完成註腳／章末註單筆轉換與下一則導覽，依 Word Mac before/after 三份真 DOCX 處理 ID 碰撞、`pStyle` 映射、Symbol 自訂標記與雙份 note-part；不安全的未知內容拒絕搬移。Docs+Shell 建置、聚焦 34/34 與型別檢查通過。Electron 首跑 footnote→endnote 完整編輯／歷史／重開案通過；endnote→footnote 在測試點擊 hover 後仍隱藏的 note-row 按鈕時逾時，尚未完成該方向實機驗收，需辨別是測試定位或產品可及性，輸出 `/private/tmp/genoffice-docs-note-conversion-results`。D6 維持 `in_progress`。

Sheets S5 Delimited 三步資料剖析精靈已接上 Ribbon：Tab／分號／逗號／空格／單字元自訂、連續分隔、雙引號、預覽、目的地、逐欄 General／Text／Skip；尚未做的 Fixed-width／Date／Advanced 明確停用。Excel Mac 覆寫既有公式會先提示並可接受；GenOffice 採保護公式的拒絕寫入並要求空白目的地，這是刻意保守的功能差異。Sheets+Shell 建置、聚焦 5/5、型別／格式／主題色檢查通過；Luna／Extra 靜態複驗無來源 blocker，但 Skip 尚無 Electron UI 實測。Electron 首跑因測試查找不存在的 Ribbon `aria-label` 失敗，已改以真實 `data-tip`，重跑中；S5 維持 `in_progress`。

Slides L5 Quick Layout 4 已根據 PowerPoint Mac 原生快照接入三個 series-local value labels、plot-level `showVal=1`、底部圖例、gapWidth 75、移除安全圖表標題與已知 value-axis majorGridlines；保留 chart data/關聯/style/colors/extension/內嵌活頁簿，不相容 labels 或未知結構拒絕變更。Slides+Shell 建置、聚焦 33/33、型別/lint 通過，完整 Electron 圖表套件單次 11/11 通過（約 1 分鐘），含 Undo/Redo、Save As、新程序重開及再編輯，結果 `/private/tmp/genoffice-slides-layout4-final-results`；Luna／Extra 獨立簽核無 blocker。既有相容 series-label 子樹保留原格式，外觀可能與 Mac 預設不同；剩餘版型 2/5/6，L5 維持 `in_progress`。

Docs D6 檢查員另發現原生 OOXML 保真缺口：Word Mac 轉換會讓正文 reference run 的 `w:rStyle` 隨 footnote/endnote 在 `af0`/`af3` 間變更，但目前匯入／寫回路徑會丟掉此 char style。即使目前 E2E 的 tag/ID/符號斷言通過，仍不能宣稱原生轉換對等；實作者已接手修復與補原生 XML 斷言。先前 note-row 操作鈕 hover-only 導致 endnote→footnote 案無法點擊的真可及性問題已改為常態可見，待新版本建置與重跑。

Docs D6 轉換切片最終修復：匯入保存正文 reference 的完整 `rawRPr`、寫回保留，轉換時依真 Word Mac 樣本將被移動引用 `af0↔af3`；無法安全映射時保守處理。註記列轉換／編輯／刪除操作鈕常態可見、可用鍵盤聚焦。Docs+Shell 重建後，補正文 `w:rStyle` 原生 XML 斷言的雙向 Electron 完整測試單次 2/2 通過（約 1.5 分鐘），含導覽、ID 碰撞、Undo/Redo、Save、fresh reopen 與再編輯，結果 `/private/tmp/genoffice-docs-note-conversion-style-results`。Luna／Extra 獨立重跑 Docs 5/5、docx-engine 4/4、typecheck/lint 並簽核無 blocker。D6 還有節末重排等範圍，仍為 `in_progress`。

Sheets S5 精靈實機最終測試 1/1 通過：真 Excel Mac before fixture 的 quoted comma、Text `003`/`005`、既有 B1 公式拒絕覆寫、改 D1 輸出、單次 Undo 清空全部六格、Redo、Save 與新程序重開皆驗證；原生 XLSX 公式及 cache 保留。前幾次 E2E 失敗來自 Ribbon 選擇器與 XML 測試讀取器（自閉合 `<c/>`、`xml:space` 的 `#text`），實際 scratch XLSX 證明 Undo 和 inlineStr 正確，測試改用 XML parser 後通過。輸出 `/private/tmp/genoffice-sheets-text-to-columns-style-results`。Luna／Extra 最終簽核待回報；Excel 可確認後覆寫既有值，GenOffice 刻意要求安全的空白目的地。Fixed-width/Date/Advanced 尚停用，Skip 尚未由 GUI 操作驗收，因此 S5 維持 `in_progress`。

Sheets S5 Luna／Extra 已依最終 Electron 1/1、聚焦 5/5 簽核 Delimited 切片無 blocker；確認安全拒絕既有公式目的地與 Excel 提示後可覆寫之差異。Fixed-width 原生 Excel Mac Step 2／Step 3、拖曳與雙擊界線、D1:F3 分欄結果和 before/after XLSX 已記錄在 `office-mac-field-evidence-2026-10-13/sheets-fixed-width/`，但 GenOffice Fixed-width 實作未建置或實機驗收。

Slides L5 Quick Layout 2 依 PowerPoint Mac 真檔接入三組 series 數值標籤、頂部圖例、gapWidth 150、overlap −25、隱藏數值軸與安全移除 majorGridlines。Slides+Shell 建置、聚焦 36/36、完整 Electron 圖表套件單次 12/12 通過（約 1 分鐘），Luna／Extra 獨立簽核無 blocker；結果 `/private/tmp/genoffice-slides-layout2-final-results`。與 Mac 快照不同，匯入數值軸的 `spPr`／`txPr` 子樹在隱藏時仍保留，畫面語意相同且可在重新顯示軸時恢復其格式。版型 5、6 及廣泛圖片／圖表工具尚未完成。

依使用者 2026-10-14 要求，後續程式修正與實測在此停止。已中止仍在研究中的 Docs 節末章節附註及 Sheets Fixed-width 實作者；階段狀態、證據和恢復順序請見 [停止點報告](office-mac-progress-stop-2026-10-14.md)。
