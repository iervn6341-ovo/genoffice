# Microsoft 365 Mac 工具列細項修正進度：停止點

時間：2026-10-14 20:43（Asia/Taipei）。依使用者要求，已停止新的程式修正與實測；Docs、Sheets 實作者已中止，沒有仍在執行的建置或 Electron 驗收。這是階段報告，不代表 Docs、Sheets、Slides 的工具列已與 Microsoft 365 Mac 桌面版全面對等。完整逐項記錄見 [本輪續行紀錄](office-mac-continuation-2026-10-14.md) 與 [任務追蹤](office-mac-task-progress-2026-10-13.json)。追蹤表目前為 15 項 completed、4 項 in_progress、9 項 queued、4 項 pending_audit；這些項目是工作包，並非 Office 全功能覆蓋率。

## 已完成且通過本輪桌面驗收的切片

| 領域 | 結果與實際驗收 | 限制 |
| --- | --- | --- |
| Docs D6 | 真 Word Mac 註腳／章節附註自訂 ♠／♥ 標記、實際頁面重啟編號與節作用範圍；單筆雙向轉換保留 note part、ID 碰撞順序、Symbol、自訂標記、正文 `w:rStyle af0↔af3`。註記列操作鈕改為常態可見。雙向 Electron 單次 **2/2**，含導覽、Undo／Redo、存檔、新程序重開與再編輯；Luna／Extra 獨立簽核。證據：`/private/tmp/genoffice-docs-note-conversion-style-results`、[原生 Word 樣本](office-mac-field-evidence-2026-10-13/docs-native-notes/conversion/)。 | 節末章節附註的跨節／跨頁重排仍未實作；D6 保持 `in_progress`。 |
| Sheets S4 | 一般／會計底線在純文字、未旋轉儲存格的畫布與原生 XLSX 寫回已經真 Excel Mac 樣本及 Electron 驗證；富文字／旋轉情況保留原生格式並阻止不支援的會計線操作。 | 富文字及旋轉儲存格的會計線畫布仍不對等。私有 document-line hook 實驗在真畫面只顯示極短線段，已撤回；需新的渲染層。S4 保持 `in_progress`。 |
| Sheets S5 | Delimited 三步精靈支援 Tab、分號、逗號、空格、自訂單字元、連續分隔、雙引號、預覽、目的地與逐欄 General／Text／Skip。真 Excel Mac fixture 的 quoted comma、文字 `003`／`005`、B1 公式保護、D1 安全輸出、單次 Undo／Redo、存檔重開與 OOXML 公式/cache；Electron **1/1**，Luna／Extra 獨立簽核。證據：`/private/tmp/genoffice-sheets-text-to-columns-style-results`、[原生 Excel 樣本](office-mac-field-evidence-2026-10-13/sheets-text-to-columns/)。 | Excel Mac 可提示後覆寫既有資料；目前 GenOffice 保守拒絕既有值／公式目的地。Fixed-width、Date、Advanced 尚未完成；Skip 只有 planner 單元測試，未走 UI。S5 保持 `in_progress`。 |
| Slides L5 | PowerPoint Mac 原生柱狀圖 Quick Layout 1、2、3、4、7、8、9、10、11 的受支援圖表變更使用稀疏 OOXML patch，保留系列資料、關聯、樣式、色彩與內嵌活頁簿；不相容匯入圖表拒絕變更。最新 Layout 2 完整 Electron 套件單次 **12/12**，含 Edit、Undo／Redo、Save As、重開與再編輯；Luna／Extra 獨立簽核。證據：`/private/tmp/genoffice-slides-layout2-final-results`、[原生 PowerPoint 樣本](office-mac-field-evidence-2026-10-13/slides-chart-native/)。 | Layout 5、6 和廣泛圖片／圖表情境工具仍缺。Layout 2 將數值軸設為隱藏時保留匯入的 `spPr`／`txPr`，與 Mac 快照的 XML 並非逐字一致；畫面隱藏語意相同。L5 保持 `in_progress`。 |

## 停止時尚未驗收的工作

- Docs D6 節末章節附註：實作者完成佈局路徑調查，尚未提交可驗收的來源切片。現有 Word Mac 樣本缺少「兩節各有章節附註、第一節註文跨頁」案例；未驗證節界、後續節推移、頁面預覽與儲存重開。先檢查已中止的工作樹變更，再用真 Word Mac 樣本定義行為。
- Sheets S5 Fixed-width：已用真 Excel Mac 建立 [來源檔](office-mac-field-evidence-2026-10-13/sheets-fixed-width/excel-mac-fixed-width-before.xlsx)、[分欄後檔](office-mac-field-evidence-2026-10-13/sheets-fixed-width/excel-mac-fixed-width-after.xlsx) 與 Step 2／Step 3 截圖。原生 Step 2 是點擊新增界線、拖曳移動、雙擊刪除；以第 4、7 字元分欄後，D1:F1=`ABCD`／`001`／`XYZ`，短列 F2 為空，純空白的 E3 為空，來源 B1 公式仍為 `10+5`。來源檔是以先前 Excel Mac 活頁簿的 OOXML 改造後在 Excel 開啟，分欄後檔由 Excel UI 儲存。實作者剛開始 fixed-width parser／尺規研究，未建置、未跑 Electron、未經 Luna 複驗；不能算完成。
- Slides L5 Layout 5、6 未實作；其他圖表與圖片工具未完成 Mac 細項比對。
- Sheets S4 富文字／旋轉會計底線仍需不同的畫布渲染方案，不能以已保存原生 `<u>` 代替可見畫面對等。

## 建議恢復順序

1. 先檢查中止時 Docs D6 與 Sheets S5 的未完成差異，避免將部分編輯誤當成已驗收；補完並以真 Word／Excel Mac 樣本、Edit→Undo／Redo→Save→新程序重開驗證。
2. 修 Sheets S4 富文字／旋轉會計底線的畫布呈現，再完成 Slides L5 Layout 5、6。此階段仍須維持未知 OOXML 結構的拒絕修改保護。
3. 推進尚未開始的 Docs D7、Sheets S6/S7、Slides L6/L7 與跨產品 M0–M3；每個切片以可編輯原生物件及往返保存為驗收門檻。
4. 最後執行 V0–V3 的 Microsoft 365 Mac 工具列細項與跨應用往返稽核，才能評估整體對等程度。

本輪工作保留在目前共用工作樹，尚未提交或建立 PR；倉庫原本已有大量其他未提交變更，恢復時應只依本報告、續行紀錄與實際 diff 界定本輪範圍。
