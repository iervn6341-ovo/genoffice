# Docs Office for Mac 獨立驗證紀錄

日期：2026-10-13（Asia/Taipei）。本紀錄搭配 `e2e/docs-office-mac-verification.spec.ts`；歷史稽核與 2026-10-12 原始證據保持不變。生成的 DOCX 僅存放在系統暫存目錄，桌面擷取及副本寫入 `office-mac-evidence-2026-10-13/docs-verification/`。

## 驗證狀態

- **原始碼檢視：完成。** Header/footer 變體編輯以局部 section reference 對應；每次內容編輯建立新 part，原始 part 保留供跨存檔 Undo 使用，並複製 part relationship。連結操作移除該節的局部 reference；跨存檔復原由 local-ref 快照差異轉成 link/ref edit。Header-only 儲存時，重建的 ProseMirror 文件沿用即時 `layout` 屬性，避免將編輯器歷史誤當本文變更而重設。
- **DOCX 引擎與單元測試：** 實作者紀錄的 Docs、UI 聚焦測試為 91 項通過，三個 TypeScript 專案檢查通過；root 後續回報 DOCX 引擎完整套件 1,420 項通過、1 項略過，整合 `build:all` 成功。`sections.test.ts` 將既有 reference 的測試改為要求原 part bytes 不變、建立新 rId 並保留唯一 reference，符合 append-only Undo 契約。這些是共享工作樹中的實作者／root 執行結果，非本驗證員獨立重跑。
- **探測檔靜態檢查：完成。** Prettier check、ESLint 均通過；Playwright `--list` 成功收集 3 個測試。這些結果只代表格式、靜態 lint 與測試可收集，沒有啟動桌面應用。
- **GenOffice 桌面 UI／存檔重開：三個案例通過。** 1280px Symbols 折疊、插入 ≤、Equation gallery 與自訂 fraction 插入通過；截圖與 JSON 在 `office-mac-evidence-2026-10-13/docs-verification/folded-symbols-2026-10-12T16-49-36-129Z/`。三節案例通過 Odd/Even/First header 變體選取、第一頁 header 編輯、工具列 Undo/Redo 的精確文字與連結狀態往返、section 3 連結操作、存檔及新 Electron 重開檢查。DOCX package 驗證 cloned header 保留樣式、table、picture relationship/bytes、PAGE/NUMPAGES；footer1.xml bytes 與原始 fixture 相同，且 first-footer 未建立多餘 reference。證據在 `office-mac-evidence-2026-10-13/docs-verification/sections-rich-header-2026-10-12T17-07-42-134Z/`。跨存檔 Undo 案例也通過：工具列 Undo/Redo 在保存前及首次保存後恢復原 header；第二次保存確認 section 2 reference 回到原始 rId 且 `header3.xml` bytes 完全相同。證據在 `office-mac-evidence-2026-10-13/docs-verification/section-header-undo-2026-10-12T17-03-27-725Z/`。這裡驗證的是工具列命令，不把曾失焦的鍵盤快捷鍵當作通過。

  **截圖界線：**rich-header 測例的保存及重開截圖顯示正文 `Section one body`，儘管頁首頁尾控制處於其他節的選取狀態，頁碼顯示 `Page 2 of 3`。原碼說明這些 Previous/Next 控制切換的是頁首頁尾編輯目標節，與正文 caret 分開；因此本輪通過的是頁首頁尾目標節／變體切換，不代表正文畫布跟隨捲動或完整 Word 的節導航。本測試未驗正文 caret／頁面對齊，該項仍待驗；截圖中的 `0 words` 也不能據此判為正文計數回歸。這項界線與觀察均不代表已確認產品缺陷。
- **Microsoft Word for Mac 原生驗收：未執行。** 請參閱既有 2026-10-12 原生 Office 證據；本輪沒有重新開啟使用者文件。

## 來源與單元檢視

最新引擎寫檔路徑可處理 `sectPr` self-closing 展開、末節 trailing `sectPr`、各節／變體的局部 reference、連結移除及舊 rId 恢復；同一 variant 同時有 pending link 與文字編輯時，內容編輯優先。新增的引擎測試涵蓋共享 rId 不污染其他節、繼承 first header 建立獨立 part 並保留表格／圖片關係、PAGE/NUMPAGES 往返和多次保存後的 ref 恢復。這些是來源與測試層證據，不能替代畫布實際預覽或 Word 原生結果。

跨存檔 Undo 的實際接受條件已通過：在節 2 的 Odd header 編輯後先 Undo/Redo，保存並確認新內容寫入，再 Undo 並保存第二次；DOCX 回到原 header reference 和原 XML bytes，而非只在畫布上看似復原。另一案例在 first header 編輯、連結 section 3、保存後確認 section 2 first 的新 part 與 section 3 reference 的移除，再於新 Electron launch 檢查 UI 與 ZIP package。E2E 中以 package bytes 和 UI 分別斷言兩層結果。

## 已知範圍

- DATE 等未建模欄位在編輯相同段落時可能被扁平化；目前新增探測只覆蓋 PAGE 與 NUMPAGES。
- 未保存的 link 編輯期間，圖片預覽仍可能來自舊 local ref；本測試保存並重開後才檢查文件 package 和可見狀態。
- 部分畫布頁首頁尾雙擊編輯尚未納入 layout history；本驗證透過 Insert → Header 編輯器操作。
- 桌面測試需在建置完成後以 `GENOFFICE_DOCS_MAC_VERIFICATION=1` 執行；執行產物會落在日期 13 的新 evidence 子目錄，不會改寫歷史證據。
