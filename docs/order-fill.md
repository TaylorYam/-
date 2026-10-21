# 起單回填 / Order Mail Fill

> 本文件中英並列：上半為繁體中文，下半為英文。修改時請同步更新兩區。
>
> This document is bilingual: Traditional Chinese first, English below. Update both sections together.

[繁體中文](#繁體中文) | [English](#english)

## 繁體中文

把上手回覆的起單信件貼到「起單回填」工作表，按「填入」回填 `Term!A2:T2` 與 `Term!O5:O9`；按「清空」清掉貼上區，再貼下一封。原始碼在 `src/vba/`。

### 安裝（每份活頁簿做一次）

1. 開啟 `特別約定事項&高齡適合度聲明書Generator_*.xlsm`，按 `Alt+F11` 開 VBE。
2. 「檔案 → 匯入檔案」，匯入 `src/vba/OrderFill.bas`。測試用的 `src/vba/OrderFillTest.bas` 視需要匯入。
3. 回到 Excel，按 `Alt+F8` 執行 `OrderFill_Setup`，會在 `Term` 後面建立「起單回填」工作表與兩個按鈕。
4. 存檔。

`.bas` 檔在工作目錄是 cp950 編碼（VBE 匯入時使用的編碼），Git 內以 UTF-8 儲存，由 `.gitattributes` 自動轉換。

### 使用方式

1. 在「起單回填」按「清空」。
2. 從上手的信件複製整段條件（含 `Initial price`、`PG rate` 等文字與表格），貼在 A7。
3. 按「填入」。完成後會跳到 `Term`，並列出信中找不到的欄位。
4. 如果判斷不出上手，在 C3 的下拉選單選上手，再按一次「填入」。C3 有值時以它為準。

### 回填規則

- 上手判斷：C3 下拉選單 → 信中 TDCC 前 3 碼（查 `Variable!A:B`）→ 各家表頭特徵。
- 依表頭名稱找值，不依欄位位置。支援三種排法：表頭列加數值列、左右一組、同一格「標題: 數值」。
- 百分比：沒有 `%` 且值 ≤ 3 時視為小數（0.8 → 80）；Coupon 為 < 1 時才換算。`96.20%→96.15%` 取最後的值。
- 年期一律填月數，`Q2 = M`。
- KO Type 會對應到 `Daily Memory`／`Daily`／`Period End`／`Period End Memory`：
  - MS：由 `KO Type` 搭配 `KO Memory` 判斷。
  - Citi：由 `Memory Autocall`／`Daily KO` 判斷。
- Barrier Type 會對應到 `AKI`／`MKI`／`EKI`／`NONE`：`European` → `EKI`，`N/A` → `NONE`。`NONE` 時 KI 留空。
- Citi 特例：Guaranteed Period = Non Callable Periods + 1；Note Price = 100 − Upfront。
- 信中沒有 Product 時預設 `FCN`（例：SG），並提醒。
- 信中沒有 TDCC 時 B2 留空；沒有 Trade date 時 O5 填今天。
- Issue／Final／Maturity date 沒有時，沿用 `Term` 既有邏輯推算，並提醒。
- 期初價格：Opening → 開盤價；Closing → 收盤價。
- 回填時會暫停 `Term` 的事件，避免信裡的日期被重算覆蓋。之後只重新觸發 D2、P2，讓「產品名稱」的幣別與期間跟著更新。
- 不會動到 `Term!U2:V2`、`Q6`、`O10`。

### 測試

在活頁簿**複本**匯入 `OrderFillTest.bas` 後，執行 `Test_OrderFill_AllIssuers`。它會把 `起單信件格式.xlsx` 的 14 張上手表逐一貼入、填入，再和 ELN每日起單建檔中同一 TDCC 的紀錄比對。

已知限制：HSBC 的日期是 `dd/mm/yyyy` 文字，程式依日在前解析。如果 Excel 貼上時已自動轉成日期，而且日、月都 ≤ 12，可能會被轉錯，回填後請核對。

---

## English

Paste an issuer's order reply mail into the `起單回填` sheet and press **填入** (Fill) to populate `Term!A2:T2` and `Term!O5:O9`. Press **清空** (Clear) to clear the paste area before pasting the next mail. Source code is in `src/vba/`.

### Installation (once per workbook)

1. Open `特別約定事項&高齡適合度聲明書Generator_*.xlsm` and press `Alt+F11` to open the VBE.
2. Use **File → Import File** to import `src/vba/OrderFill.bas`. Import the test module `src/vba/OrderFillTest.bas` if needed.
3. Back in Excel, press `Alt+F8` and run `OrderFill_Setup`. It creates the `起單回填` sheet and its two buttons after `Term`.
4. Save the workbook.

The `.bas` files are cp950 in the working tree (the encoding VBE imports) and UTF-8 inside Git; `.gitattributes` converts between them.

### Usage

1. On `起單回填`, press **清空**.
2. Copy the whole term block from the issuer's mail, including text such as `Initial price` and `PG rate` and the table, and paste it at A7.
3. Press **填入**. The workbook then shows `Term` and lists any fields the mail did not contain.
4. If the issuer cannot be detected, pick it from the C3 drop-down and press **填入** again. A value in C3 always takes precedence.

### Fill rules

- Issuer detection: C3 drop-down → first 3 digits of a TDCC in the mail (looked up in `Variable!A:B`) → issuer-specific header signatures.
- Values are found by header name, not column position. Three layouts are supported: a header row with a value row below, label/value side by side, and `Label: value` in one cell.
- Percentages: a value without `%` that is ≤ 3 is treated as a decimal (0.8 → 80); Coupon is converted only when < 1. `96.20%→96.15%` takes the last value.
- Tenor is always written in months, with `Q2 = M`.
- KO Type is mapped to `Daily Memory` / `Daily` / `Period End` / `Period End Memory`:
  - MS: derived from `KO Type` plus `KO Memory`.
  - Citi: derived from `Memory Autocall` / `Daily KO`.
- Barrier Type is mapped to `AKI` / `MKI` / `EKI` / `NONE`: `European` → `EKI`, `N/A` → `NONE`. KI is left blank when the type is `NONE`.
- Citi special cases: Guaranteed Period = Non Callable Periods + 1; Note Price = 100 − Upfront.
- When the mail has no Product, `FCN` is used (for example, SG) and a warning is shown.
- When the mail has no TDCC, B2 is left blank. When it has no Trade date, O5 is set to today.
- Missing Issue / Final / Maturity dates are calculated by the existing `Term` logic, with a warning.
- Initial price: Opening → 開盤價; Closing → 收盤價.
- `Term` events are suspended while writing, so mail dates are not overwritten by recalculation. Only D2 and P2 are re-triggered afterward, so the currency and tenor shown on `產品名稱` stay in sync.
- `Term!U2:V2`, `Q6`, and `O10` are not changed.

### Testing

In a **copy** of the workbook, import `OrderFillTest.bas` and run `Test_OrderFill_AllIssuers`. It pastes each of the 14 issuer sheets in `起單信件格式.xlsx`, runs the fill, and compares the result with the record for the same TDCC in ELN每日起單建檔.

Known limitation: HSBC dates are `dd/mm/yyyy` text and are parsed day-first. If Excel has already converted such a value into a date when it was pasted, and both day and month are ≤ 12, the date may be wrong. Check it after filling.
