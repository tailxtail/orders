# SPEC — CSV → ODS(PRINT_ALL) → PDF（Windows Server 2022 + LibreOffice headless）

## A) 專案背景與目標
- 作業系統：Windows Server 2022 Datacenter
- Python：3.12（使用 venv）
- 目的：將 CSV 多筆資料批次填入 `templates.ods` 的 `PRINT_ALL` 分頁中，依序追加成「單一 ODS」，最後用 LibreOffice headless 轉成「單一 PDF」（不分段）。
- 模板：`templates.ods` 已準備好，並包含：
  - 分頁名稱：`PRINT_ALL`
  - A～R 欄各欄寬
  - 列印版型：A5、橫向、邊界、縮放（適應頁寬/頁高設定已在模板內完成）
  - 頁首頁尾不使用
- 每筆資料佔用固定區塊高度：31 列
  - 第 n 筆起始列：`start_row = 1 + (n-1)*31`
  - 母版區塊在 `PRINT_ALL` 的第 1 列到第 31 列（Row 1..31）

## B) 固定路徑與目錄結構
- Worker base：`C:\Users\user03\Documents\lo-workers\w01\`
- Input：`C:\Users\user03\Documents\lo-workers\w01\in\`
- Output：`C:\Users\user03\Documents\lo-workers\w01\out\`
- Profile：`C:\Users\user03\Documents\lo-workers\w01\profile\`
- Code：`C:\Users\user03\Documents\lo-workers\w01\code\`
  - venv：`...\code\venv\`
  - scripts：`...\code\scripts\`

## C) 明確需求與限制（最重要）

### C1. 主要流程（必須完全遵守）
1. 複製 `in\templates.ods` → `out\output.ods`（只保留 `PRINT_ALL` 工作表）
2. 逐筆資料 n（n 從 1 開始）：
   - 以母版區塊（Row 1..31）為來源，clone 到第 n 筆區塊（必須保留：文字內容、合併儲存格、列高、儲存格樣式、框線）
   - 寫入前，必須先將以下儲存格/範圍「清空內容（set empty）」，避免殘值：
     - `D5, D6, N6, R5, R7, C13:C28, F13:F28, L13:L28, N13:N28, O13:O28, O29`
     - 清空僅影響儲存格內容（value/text）；不得更動任何樣式、合併、框線、列高、欄寬
   - 將第 n 筆資料的計算結果（值）寫入對應儲存格
   - 若 n >= 2，必須在該筆 `start_row` 實現「Page Break Before」
3. 最後更新列印範圍，確保涵蓋最後一筆（避免 PDF 漏尾端）
4. 由外部 PowerShell 命令呼叫 LibreOffice headless 把 `output.ods` 轉 PDF

### C2. ODS 操作限制
- 必須使用 `odfpy` 操作 ODS（必要時可直接操作 ODS 內部 XML 結構）
- 絕不新增 style：
  - 所有列/格樣式必須重用模板中既有 style name
  - clone 時只能複製節點並保留 style reference，避免 style 膨脹
- 不逐格重建格式：
  - 允許在程式中迴圈處理 row node clone
  - 不允許逐格重新建立樣式/框線/合併等格式

### C3. Page Break（分頁）規格（唯一真實來源）
- `templates.ods` 的 `PRINT_ALL` 第 1000 列已預先插入 `Row Break`（分頁樣板列 Row 1000）
- 程式在 n>=2 時，必須對該筆 `start_row` 實現「Page Break Before」
- 不得新增任何 style
- 必須重用模板內已存在的分頁資訊（來源為 Row 1000），允許方式：
  - 從 Row 1000 對應的 ODF 結構/屬性引用複用到 `start_row`（例如複製 style reference/屬性到目標列）

### C4. 列印範圍更新規格（避免 Codex 猜測）
- 欄固定為 A:R
- 最後一筆的區塊結束列：`end_row = start_row_last + 30`
- 列印範圍必須設定為：`A1:R{end_row}`

### C5. 量級
- 可能 2000+ 筆資料，單一 PDF，不做分段

## D) CSV 與計算邏輯

### D1. CSV 基本資訊
- 檔名：`in\input.csv`
- 第 1 列為 header，第 2 列開始是第一筆資料

### D2. CSV 欄位（header）
```text
"Serial No","Order No","Order Date","Customer Name","Customer Phone","Customer Email","Grand Total","Location",
"Product 1 Name","Product 2 Name","Product 3 Name","Product 4 Name","Product 5 Name","Product 6 Name","Product 7 Name","Product 8 Name","Product 9 Name","Product 10 Name","Product 11 Name","Product 12 Name","Product 13 Name","Product 14 Name","Product 15 Name","Product 16 Name","Product 17 Name","Product 18 Name","Product 19 Name","Product 20 Name","Product 21 Name","Product 22 Name","Product 23 Name","Product 24 Name","Product 25 Name",
"Product 1 SKU","Product 2 SKU","Product 3 SKU","Product 4 SKU","Product 5 SKU","Product 6 SKU","Product 7 SKU","Product 8 SKU","Product 9 SKU","Product 10 SKU","Product 11 SKU","Product 12 SKU","Product 13 SKU","Product 14 SKU","Product 15 SKU","Product 16 SKU","Product 17 SKU","Product 18 SKU","Product 19 SKU","Product 20 SKU","Product 21 SKU","Product 22 SKU","Product 23 SKU","Product 24 SKU","Product 25 SKU",
"Product 1 Quantity","Product 2 Quantity","Product 3 Quantity","Product 4 Quantity","Product 5 Quantity","Product 6 Quantity","Product 7 Quantity","Product 8 Quantity","Product 9 Quantity","Product 10 Quantity","Product 11 Quantity","Product 12 Quantity","Product 13 Quantity","Product 14 Quantity","Product 15 Quantity","Product 16 Quantity","Product 17 Quantity","Product 18 Quantity","Product 19 Quantity","Product 20 Quantity","Product 21 Quantity","Product 22 Quantity","Product 23 Quantity","Product 24 Quantity","Product 25 Quantity",
"Product 1 Price","Product 2 Price","Product 3 Price","Product 4 Price","Product 5 Price","Product 6 Price","Product 7 Price","Product 8 Price","Product 9 Price","Product 10 Price","Product 11 Price","Product 12 Price","Product 13 Price","Product 14 Price","Product 15 Price","Product 16 Price","Product 17 Price","Product 18 Price","Product 19 Price","Product 20 Price","Product 21 Price","Product 22 Price","Product 23 Price","Product 24 Price","Product 25 Price",
"Product 1 Total","Product 2 Total","Product 3 Total","Product 4 Total","Product 5 Total","Product 6 Total","Product 7 Total","Product 8 Total","Product 9 Total","Product 10 Total","Product 11 Total","Product 12 Total","Product 13 Total","Product 14 Total","Product 15 Total","Product 16 Total","Product 17 Total","Product 18 Total","Product 19 Total","Product 20 Total","Product 21 Total","Product 22 Total","Product 23 Total","Product 24 Total","Product 25 Total",
"Additional Products"

### D3. 欄位映射（PRINT_ALL 區塊內座標，皆以“母版座標”表示）
- 出貨日期 → D5：`Order Date` 去除時分秒（原格式 `YYYY-MM-DD HH:MM:SS`，輸出 `YYYY-MM-DD`）
- 出貨客戶 → D6：`Customer Name`
- 出貨單號 → N5：`Serial No`
- 訂單號碼 → R5：`Order No`
- 聯絡電話 → R7：`Customer Phone`
- 產品編號 → C13..C28：`Product 1 SKU`..`Product 16 SKU`
- 產品名稱 → F13..F28：`Product 1 Name`..`Product 16 Name`
- 出貨數量 → L13..L28：`Product 1 Quantity`..`Product 16 Quantity`
- 出貨單價 → N13..N28：`Product 1 Price`..`Product 16 Price`
- 出貨金額 → O13..O28：`Product 1 Total`..`Product 16 Total`
- 出貨總額 → O29：`Grand Total` 金額前面要增加貨幣符號`$`

### D4. 座標位移規則（第 n 筆）
- 第 n 筆相對母版的列位移：`row_offset = (n-1)*31`
- 寫入與清空的所有座標都必須套用 row_offset（欄不變）

### D5. 公式替代（僅供參考，實作以 Python 為準）
（略，照原始清單）

## E) log.txt 規格（統一輸出位置與格式）
- log 檔案：`out\log.txt`
- 一律 append（追加寫入，不覆蓋）
- 每一行建議使用 CSV 格式或 key=value（擇一），必須至少包含：
  - `Reason`（原因代碼）
  - `Serial No`, `Order No`, `Order Date`, `Customer Name`, `Grand Total`
- 觸發條件：
  1) 若 `Product 17~25` 任一欄位有值（Name/SKU/Quantity/Price/Total 任一非空），寫入 log：
     - Reason=`HAS_PRODUCT_17_25`
  2) 金額驗算不相等，寫入 log：
     - Reason=`TOTAL_MISMATCH`
     - 必須額外包含：`Sum(Product Totals)`, `Diff`

## F) Decimal 解析與驗算規格（不得使用 float）
- 所有金額與數量使用 `Decimal` 計算，不得用 float
- 金額欄位可能包含：千分位逗號 `,`、貨幣符號（例如 `$`）、前後空白
  - 解析時需移除 `,`、`$`、空白
  - 空字串視為 0
  - 若解析失敗：寫入 log（Reason=`PARSE_ERROR`）並將該欄位視為 0（或直接 fail 該筆；請在程式內採用「視為 0」）
- 驗算規則：
  - `sum(O13:O28)` 必須精確等於 `O29`
  - `Diff = GrandTotal - SumTotals`

## G) Codex 需要交付的檔案與內容
- `code\requirements.txt`：至少含 `odfpy`
- `code\scripts\build_output_ods.py`：
  - 讀取 `in\input.csv`（檔名可參數化）
  - 複製 `in\templates.ods` → `out\output.ods`
  - 只保留 `PRINT_ALL`
  - 每筆資料追加區塊：以母版 Row 1..31 做 clone 到對應列
  - 清空指定儲存格內容（套用 row_offset）
  - 寫入每筆資料（套用 row_offset）
  - Page Break Before：n>=2 對 start_row 套用，且必須重用 Row 1000 分頁資訊、不得新增 style
  - 更新列印範圍：`A1:R{end_row}`
  - 輸出 log（至少包含總筆數、耗時、output.ods 路徑）
- `code\scripts\run_convert_pdf.ps1`：
  - 使用指定 `UserInstallation` profile 呼叫 `soffice --headless --convert-to pdf`
  - 來源：`out\output.ods`
  - 輸出：`out\output.pdf`

## H) 交付順序（里程碑）
- M1：能複製 templates.ods → output.ods，且只留 PRINT_ALL，存檔可用
- M2：能將母版 Row 1..31 clone 到第 2 筆（32..62），且合併/列高/樣式不變
- M3：能清空指定儲存格內容，再寫入第 1 筆資料（先做 1 筆）
- M4：Page Break Before（不得新增 style）
  - 對 n=2 的 start_row=32 套用 Page Break Before
  - Page Break 必須由 Row 1000 分頁樣板列複用而來
  - 驗收：用 LibreOffice 開 output.ods → 分頁預覽，確認第 32 列前有切頁線
- M5：能跑完整 CSV（先 10 筆），最後更新列印範圍
- M6：加入 Decimal 驗算與 log.txt（append）
- M7：提供 run_convert_pdf.ps1 串接 soffice 轉出 PDF
