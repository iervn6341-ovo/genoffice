# 本地逐項測試紀錄（2026-09-26，macOS）

依 `genoffice-qa` skill 的 `testing/*.md` 清單，在本機 Mac 上**實際操作 GenOffice** 測試。重點放在雲端報告（`project-completeness-review.md`）標為 NOT TESTED / BLOCKED 的項目：Mac 鍵盤、長序列復原、跨功能流程、第二螢幕、與 Office 的比對。

- **基準：** `qa/local-fixes` @ `67334ff`，已執行 `npm run build:all`。
- **環境：** macOS、Apple M1 Pro，接了 3 個螢幕（內建、ASUS VS229、Sidecar）。已安裝 Word／Excel／PowerPoint for Mac。
- **方法：** 用 Playwright 驅動實際編譯出來的 app，探測腳本放在 `e2e/_probe/`（沒有進 git）。每一項都讀取狀態或存出的檔案來判定。
- **Office 比對：** 每一項都註明「直接觀察」或「推定」。

## 1. 失敗項目（產品問題）

| # | 服務 | 項目 | 嚴重度 | 重現步驟 | 預期（Office） | 實際 | 證據 |
|---|---|---|---|---|---|---|---|
| 1 | Slides | P-08／P-12：Shift+點擊加選（**已撤回**，是探測腳本的錯誤，見 §6） | ~~Medium~~ | 畫兩個空白矩形 → 點 A 的中央 → 按住 Shift 點 B 的**中央** → ⌘G | B 加入選取，⌘G 群組成功（PowerPoint 行為；**推定**） | B 沒有加入選取，只剩 A 被選取，⌘G 不起作用。改點形狀**邊緣**附近就能加選（P-08b 通過）；⌘A 全選後 ⌘G 也正常（P-12b） | 截圖 `p12-shift.png`；每次都能重現 |
| 2 | Slides | P-02：複製投影片快速鍵 | Low | 選取縮圖 → 按 ⌘⇧D | 複製該投影片（在 PowerPoint for Mac **直接觀察**：⌘⇧D 把 2 張變成 3 張；⌘D 沒有反應） | 沒有反應。Slides 的選單沒有這個快速鍵，只能從縮圖的右鍵選單複製 | P-02f |
| 3 | Slides | P-02：縮圖多選 | Medium（功能缺口） | 點縮圖 1 → Shift+點縮圖 3 | 選取 1 到 3（PowerPoint；推定） | 只選到 1 張。雲端的程式碼稽核也標為「未找到」 | P-02b |
| 4 | Sheets | X-12：數值格式下拉選單 | **Medium**（工具列狀態不符） | 存出一個貨幣格式（`$#,##0.00`）的 xlsx → 重新開啟 → 選取那個儲存格 | 下拉選單顯示貨幣或自訂格式 | 顯示 **General**。在同一個工作階段內設定的格式則正確顯示 Currency（X-12a）。資料沒有受影響：公式列顯示 1200，改成百分比後顯示 120000%，再存檔仍是數字 | X-12b |
| 5 | Sheets | X-09：循環參照提示 | Low | 在 C8 輸入 `=C8` | Excel 的狀態列顯示「Circular References: A1」（**直接觀察**）。Excel 的警告對話框本次沒有觀察到 | 儲存格顯示 0，沒有任何提示。App 不會卡住，仍可繼續輸入 | X-09 |

## 2. 本次通過的項目（雲端原本標為 NOT TESTED / BLOCKED）

| 服務 | 項目 | 內容 | 狀態 |
|---|---|---|---|
| Word | W-02 | emoji（含膚色修飾）、CJK、長段落輸入；Backspace 一次刪除整個 emoji | PASS |
| Word | W-13 | 復原壓力：輸入 → 全選加粗 → Enter → 輸入 → 刪除，全部復原回到起點，再全部取消復原回到最終狀態 | PASS |
| Word | W-13b | A → B → 復原 → C 之後，取消復原不會把 B 帶回來 | PASS |
| Word | W-18、W-14 | 報告流程（Heading 1、粗體／斜體、`* ` 清單、2×2 表格）→ 另存 → 檢查 XML → 重開 → 繼續編輯 → 存檔 | PASS |
| Sheets | X-04 | 輸入後按 Enter 會確認並下移；按 Esc 取消編輯 | PASS |
| Sheets | X-06 | Tab、Tab、Enter 之後回到起始欄的下一列（Excel 行為） | PASS |
| Sheets | X-11 | 從 ribbon 合併 A10:B10，復原後取消合併，取消復原後再次合併；存檔的 XML 有 mergeCell | PASS |
| Sheets | X-14 | 用 ⌘D 向下填滿 | PASS |
| Sheets | X-18 | 5 次輸入逐一復原，再逐一取消復原 | PASS |
| Sheets | X-23、X-19 | 費用表：SUM、⌃⇧$ → 另存 → 公式和快取值 `<v>1550.5</v>` → 重開 → 改數值後重算 → 再存檔仍是數字 | PASS |
| Slides | P-17 | 簡報者檢視與雙螢幕放映，雲端標為 BLOCKED 的 5 項，連同相關 spec 共 9 項 | **9/9 PASS**（0 項跳過） |
| Slides | P-02 | 用「新投影片」新增；在縮圖按 Delete 刪除，按 ⌘Z 還原 | PASS |
| Slides | P-12b、P-08b | ⌘A 全選或點邊緣加選 → ⌘G 群組，⌘⇧G 取消群組 | PASS |
| Slides | P-19 | 連續 3 次復原、3 次取消復原 | PASS |
| Slides | P-24、P-20 | 插入文字方塊 → 從選單另存 → 重開：投影片數和順序（依 `sldIdLst`）與文字都正確 | PASS |
| Markdown | M-01 | Mac 上 Home/End 移到行首／行尾 | PASS |
| Markdown | M-02 | 改一段後存檔，其他行（front matter、表格、參考式連結）逐位元組相同 | PASS |
| PDF | F-06 | 用測試 PDF 裁剪第 1 頁，⌘Z 後檔案的 CropBox 還原 | PASS |
| 全部 | 執行期錯誤 | Docs、Sheets、Slides、Markdown、PDF 各工作階段都沒有 console error 或 pageerror | PASS |

## 3. 無法判定或尚未測試

| 項目 | 狀態 | 原因 |
|---|---|---|
| P-14 物件複製貼上 | 大致通過（3/4） | 有 1 次失敗，發生在 P-12 失敗之後、選取狀態不乾淨的時候。要在乾淨狀態下重測 |
| P-02e 刪除後復原，投影片回到原位置 | NOT TESTED | 渲染出來的投影片資料沒有 id，無法分辨順序。張數有正確還原 |
| CJK IME 組字 | NOT TESTED | Playwright 無法模擬真正的 IME |
| 跨文件、跨活頁簿貼上 | NOT TESTED | 需要同時開兩份文件 |
| 效能（W-17、X-21、P-22） | NOT TESTED | 沒有量測 |
| Excel 循環參照的警告對話框 | NOT TESTED | 背景輸入時 `=` 被吃掉，沒能用鍵盤輸入重現。只觀察到狀態列的提示 |

## 4. 測試環境陷阱（已補進 `testing/README.md`）

以下幾點一開始都被誤判成失敗，查證後都不是產品問題：

1. Playwright 的 `click({ modifiers: ['Shift'] })` 在這個 Electron 環境送出的 `shiftKey` 是 false（用 DOM 監聽確認過）。要改用 `keyboard.down('Shift')`。
2. 串流載入的活頁簿，用除錯介面 `getValue()` 讀到的是顯示用的文字（例如 `"$1,200.00"`）；公式列、公式和存檔用的仍然是數字。
3. Slides 的 ⌘S／⌘⇧S 和 Markdown 的 ⌘S 是原生選單快捷鍵，Mac 上 Slides 也沒有 ribbon 的 File 分頁。要從主程序呼叫選單項目。
4. pptx 裡 `slideN.xml` 的檔名順序不等於投影片順序，順序要看 `presentation.xml` 的 `sldIdLst`。

## 5. 判定（修正後）

| 服務 | 判定 |
|---|---|
| Word（Docs） | PASS WITH KNOWN LIMITATIONS（IME、效能未測） |
| Excel（Sheets） | PASS WITH KNOWN LIMITATIONS（#3、#5 已修正並有回歸測試；Excel 的循環參照警告對話框沒有實作） |
| PowerPoint（Slides） | PASS WITH KNOWN LIMITATIONS（#2、#4 已修正；多選只支援 Delete） |
| Markdown | PASS WITH KNOWN LIMITATIONS |
| PDF | PASS WITH KNOWN LIMITATIONS |

## 6. 修正結果（同日，本地修正後重測）

| # | 原問題 | 結論 | 修正 | 驗證 |
|---|---|---|---|---|
| 1 | Shift+點擊形狀中央無法加選 | **撤回，不是產品問題** | 無。探測腳本在前一次改寫時把 Shift 同時套在兩次點擊上，第二次 Shift+點擊把 B 取消選取（PowerPoint 也是這樣切換）。修正腳本後，點中央加選和 ⌘G 都正常 | P-12、P-08b、P-14 各連續 2 次 PASS |
| 2 | 縮圖無法多選 | 已實作 | Shift+點擊選取範圍、⌘+點擊切換；選到的縮圖加上 accent 外框；Delete 用一個交易刪除全部（`slides:delete-slides`，一次 ⌘Z 就全部還原）；全選時保留至少 1 張 | 新增 e2e `slides-thumb-multiselect`（2/2），順便驗證了 ⌘Z 後投影片**回到原本的順序**（原 P-02e 的 NOT TESTED 也補上了）；單元測試 `keyboard-actions` |
| 3 | 重開後數值格式選單顯示 General | 已修正 | 串流載入寫入格子時，不會觸發選單監聽的 mutation。現在每一步串流完成（`afterStream`）都重新讀取目前選取的格式 | 新增 e2e `sheets-numfmt-echo-reopen`：有修正 3/3 PASS；拿掉修正後 2/2 FAIL（顯示 "General"） |
| 4 | 沒有 ⌘⇧D 複製投影片 | 已修正 | 縮圖窗格（沒有選取物件時）按 ⌘⇧D，會在目前投影片後面插入一份複本；⌘D 維持不動作，跟 PowerPoint for Mac 一樣 | 新增 e2e `slides-duplicate-shortcut`（2/2），並驗證 ⌘Z；單元測試 |
| 5 | 循環參照沒有提示 | 已修正 | 被動讀取已載入工作表的公式文字，自己解析參照（A1、範圍、整欄／整列、跨工作表），再用強連通分量找出循環，停止編輯或重算 0.4 秒後才檢查。狀態列顯示「Circular References: D4」（在其他工作表時只顯示「Circular References」），用詞跟 Excel 一樣，20 種語言都有。**第一版**用了 Univer 的 `getAllDependencyTrees()`，但它其實會送出公式引擎的 mutation，和開檔、MCP 切換工作表產生時序競爭（`mcp-sheet-values` 5 次裡失敗 2 次），所以已經換掉 | 單元測試 7 項（參照解析、跨工作表循環、範圍包含自己、10,000 格長鏈）；e2e `sheets-circular-reference` 3/3；`mcp-sheet-values` 改完後 5/5 |

限制：多選目前只支援 Delete，其他操作（複製、⌘⇧D、拖曳）仍然只作用於目前的投影片。循環參照的偵測不追蹤名稱、INDIRECT／OFFSET 和結構化參照；公式超過 20,000 個時不再重新計算；串流模式下還沒載入的部分也不會納入。
