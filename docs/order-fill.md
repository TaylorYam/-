# 起單回填 / Order Mail Fill

> 本文件中英並列：上半為繁體中文，下半為英文。修改時請同步更新兩區。
>
> This document is bilingual: Traditional Chinese first, English below. Update both sections together.

[繁體中文](#繁體中文) | [English](#english)

## 繁體中文

把上手回覆的起單信件貼到「起單回填」工作表，按「填入」回填 `Term!A2:T2` 與 `Term!O5:O9`；按「清空」清掉貼上區，再貼下一封。原始碼在 `src/vba/`。

### 安裝

repo 裡的 `特別約定事項&高齡適合度聲明書Generator_*.xlsm` 已內建 `OrderFill`、`OrderFillTest` 模組與「起單回填」工作表，開啟即可使用。

用在其他活頁簿，或模組改版後要更新時：

1. 開啟活頁簿，按 `Alt+F11` 開 VBE。
2. 已有舊版 `OrderFill`／`OrderFillTest` 時，先在專案視窗對模組按右鍵「移除」（不必匯出）。
3. 「檔案 → 匯入檔案」，匯入 `src/vba/OrderFill.bas`。測試用的 `src/vba/OrderFillTest.bas` 視需要匯入。
4. 回到 Excel，按 `Alt+F8` 執行 `OrderFill_Setup`，會在 `Term` 後面建立「起單回填」工作表與兩個按鈕。
5. 存檔。

`src/vba/` 的 `.bas` 是原始碼。修改 `.bas` 後，要把新版匯入 repo 的 `.xlsm` 並一起 commit，讓兩者一致。

`.bas` 檔在工作目錄是 cp950 編碼（VBE 匯入時使用的編碼），Git 內以 UTF-8 儲存，由 `.gitattributes` 自動轉換。

### 使用方式

1. 在「起單回填」按「清空」。
2. 從上手的信件複製整段條件（含 `Initial price`、`PG rate` 等文字與表格），貼在 A7。
3. 信裡沒有 TDCC 時，在 C5 手填 12 碼 TDCC（文字格式，開頭的 0 會保留）。
4. 按「填入」。完成後會跳到 `Term`，並列出信中找不到、認不得或不一致的欄位。
5. 如果判斷不出上手，在 C3 的下拉選單選上手，再按一次「填入」。C3 有值時以它為準。

已經匯入舊版模組的活頁簿，匯入新版後請再執行一次 `OrderFill_Setup`，才會出現 C5。

### 回填規則

對照規則與 structured-product-workflow 讀 `(Order)` 信的規則一致（該 repo commit `f0bcce0` 之前的 `src/sp_workflow/order_formats.py`、`order_template.py`）。原則是不猜值：讀不到、認不得、兩處不一致的欄位留空，並在摘要提醒。

- TDCC：C5 有值時以它為準（和信裡的不同時提醒）；C5 不是 12 碼數字時不填入，`Term` 不變。
- 上手判斷：C3 下拉選單 → TDCC 前 3 碼（查 `Variable!A:B`）→ 各家表頭特徵。
- 依表頭名稱找值，不依欄位位置。支援三種排法：表頭列加數值列、左右一組、同一格「標題: 數值」（也認全形冒號；GS 的 `[ Issue Date: 20Oct26 ]` 這種一格多組會逐組拆開）。
- 同一欄在信裡出現兩個不同的值時留空，提醒會列出各處的值。
- 百分比看儲存格：數字格式有 `%`（信件貼上後 `80.00%` 會變成 0.8 加 `%` 格式），或文字帶 `%`，就是百分比數字（→ 80）。沒有 `%` 時照填；如果值看起來像小數（Strike／KO／KI／Note Price ≤ 3、Coupon < 1），會提醒確認。
- 改價取最後的值並提醒：`96.20%→96.15%`、`99.65->99.62`、`99.65=>99.62`、`99.92 >> 99.89`、`99.92 > 99.89`。
- 日期另認 `20Oct26`、`15-Oct-26`、`20-Oct-2026`、`20 Oct 2026`。
- 年期一律填月數，`Q2 = M`。SG 沒有 Tenor：`Settlement Frequency` 是 `Monthly` 時 Tenor = `No. of Periods`，否則留空並提醒。
- KO Type 只認 `Daily`、`Daily Memory`、`Period End`、`Period End Memory`（不分大小寫）：
  - MS：`OBSERVATION` 是每期觀察，由 `KO Memory`（Yes／No）決定是否記憶式；KO Memory 空白、或和 KO Type 矛盾（例：`Daily Memory` 但 `No`）時留空並提醒。
  - Citi：由 `Daily KO`、`Memory Autocall` 兩欄 TRUE／FALSE 判斷。
  - 其他寫法留空並提醒原文。
- Barrier Type 對應到 `AKI`／`MKI`／`EKI`／`NONE`：`European`、`At Maturity` → `EKI`，`American` → `AKI`，`No` → `NONE`；其他寫法留空並提醒原文。
  - Barrier Type 與 KI 都沒有（或是 `N/A`）、但條件表有讀到時填 `NONE`（例：SG）。
  - `NONE` 時 KI 留空；KI 有非 0 的值時提醒。
- Note Price：沒有 Note Price、只有 `Upfront (%)`（Citi）時，Note Price = 100 − Upfront。
- Citi 特例：Guaranteed Period = Non Callable Periods + 1。
- 標的超過 5 檔時提醒（`Term` 只放得下 5 檔）。
- 信中沒有 Product 時預設 `FCN`（例：SG），並提醒。
- 信中沒有 Trade date 時 O5 填今天。
- Issue／Final／Maturity date 沒有時，沿用 `Term` 既有邏輯推算，並提醒。
- 期初價格：Opening → 開盤價；Closing → 收盤價；VWAP → VWAP；其他寫法留空並提醒。
- 回填時會暫停 `Term` 的事件，避免信裡的日期被重算覆蓋。之後只重新觸發 D2、P2，讓「產品名稱」的幣別與期間跟著更新。
- 不會動到 `Term!U2:V2`、`Q6`、`O10`。

### 測試

在活頁簿**複本**執行以下測試（repo 的 `.xlsm` 已內建 `OrderFillTest`；其他活頁簿要先匯入 `OrderFillTest.bas`）：

- `Test_OrderFill_AllIssuers`：把 `起單信件格式.xlsx` 的 14 張上手表逐一貼入、填入，再和 ELN每日起單建檔中同一 TDCC 的紀錄比對；接著跑規則測試。
- `Test_OrderFill_Rules`：只跑規則測試（不需要樣本檔），在貼上區寫入合成的信件內容，逐條檢查上面的回填規則。
- `Test_OrderFill_Clear`：先選 C3 再按「清空」，確認不會出現執行階段錯誤 1004、貼上區的圖片會刪掉，按鈕與 C3 下拉清單保留。

已知限制：HSBC 的日期是 `dd/mm/yyyy` 文字，程式依日在前解析。如果 Excel 貼上時已自動轉成日期，而且日、月都 ≤ 12，可能會被轉錯，回填後請核對。

---

## English

Paste an issuer's order reply mail into the `起單回填` sheet and press **填入** (Fill) to populate `Term!A2:T2` and `Term!O5:O9`. Press **清空** (Clear) to clear the paste area before pasting the next mail. Source code is in `src/vba/`.

### Installation

The repository's `特別約定事項&高齡適合度聲明書Generator_*.xlsm` already contains the `OrderFill` and `OrderFillTest` modules and the `起單回填` sheet, so it works as soon as it is opened.

To use it in another workbook, or to update after the modules change:

1. Open the workbook and press `Alt+F11` to open the VBE.
2. If an older `OrderFill` / `OrderFillTest` is present, right-click it in the Project window and choose **Remove** (no export needed).
3. Use **File → Import File** to import `src/vba/OrderFill.bas`. Import the test module `src/vba/OrderFillTest.bas` if needed.
4. Back in Excel, press `Alt+F8` and run `OrderFill_Setup`. It creates the `起單回填` sheet and its two buttons after `Term`.
5. Save the workbook.

The `.bas` files in `src/vba/` are the source. After changing a `.bas` file, import the new version into the repository's `.xlsm` and commit both together so they stay in sync.

The `.bas` files are cp950 in the working tree (the encoding VBE imports) and UTF-8 inside Git; `.gitattributes` converts between them.

### Usage

1. On `起單回填`, press **清空**.
2. Copy the whole term block from the issuer's mail, including text such as `Initial price` and `PG rate` and the table, and paste it at A7.
3. If the mail has no TDCC, type the 12-digit TDCC in C5 (a text cell, so leading zeros are kept).
4. Press **填入**. The workbook then shows `Term` and lists any fields that were missing, not recognized, or inconsistent.
5. If the issuer cannot be detected, pick it from the C3 drop-down and press **填入** again. A value in C3 always takes precedence.

In a workbook that already has the older module, run `OrderFill_Setup` again after importing the new one so that C5 appears.

### Fill rules

The mapping follows the rules structured-product-workflow used to read `(Order)` emails (`src/sp_workflow/order_formats.py` and `order_template.py` before commit `f0bcce0` in that repository). Values are never guessed: a field that is missing, not recognized, or inconsistent is left blank and listed in the summary.

- TDCC: a value in C5 takes precedence (with a warning when the mail's TDCC differs). If C5 is not 12 digits, nothing is filled and `Term` is unchanged.
- Issuer detection: C3 drop-down → first 3 digits of the TDCC (looked up in `Variable!A:B`) → issuer-specific header signatures.
- Values are found by header name, not column position. Three layouts are supported: a header row with a value row below, label/value side by side, and `Label: value` in one cell (a full-width colon also works, and a cell with several `[ Label: value ]` groups, such as GS's `[ Issue Date: 20Oct26 ]`, is split group by group).
- When a field appears twice in the mail with different values, it is left blank and the summary lists every value.
- Percentages follow the cell: a number with a `%` number format (pasting `80.00%` from a mail gives 0.8 formatted as `%`) or text containing `%` is a percentage figure (→ 80). A value without `%` is written as is; if it looks like a decimal (Strike / KO / KI / Note Price ≤ 3, Coupon < 1), a warning asks for a check.
- Revised values take the last value, with a warning: `96.20%→96.15%`, `99.65->99.62`, `99.65=>99.62`, `99.92 >> 99.89`, `99.92 > 99.89`.
- Dates also accept `20Oct26`, `15-Oct-26`, `20-Oct-2026`, and `20 Oct 2026`.
- Tenor is always written in months, with `Q2 = M`. SG has no Tenor: when `Settlement Frequency` is `Monthly`, Tenor = `No. of Periods`; otherwise it is left blank with a warning.
- KO Type accepts only `Daily`, `Daily Memory`, `Period End`, and `Period End Memory` (case-insensitive):
  - MS: `OBSERVATION` means period-end observation, and `KO Memory` (Yes / No) decides memory. A blank KO Memory, or one that contradicts KO Type (for example, `Daily Memory` with `No`), leaves it blank with a warning.
  - Citi: derived from the TRUE / FALSE columns `Daily KO` and `Memory Autocall`.
  - Any other value is left blank, and the warning shows the original text.
- Barrier Type is mapped to `AKI` / `MKI` / `EKI` / `NONE`: `European` and `At Maturity` → `EKI`, `American` → `AKI`, `No` → `NONE`. Any other value is left blank, and the warning shows the original text.
  - When both Barrier Type and KI are absent (or `N/A`) but the term table was read, `NONE` is used (for example, SG).
  - KI is left blank for `NONE`; a non-zero KI triggers a warning.
- Note Price: when there is no Note Price but only `Upfront (%)` (Citi), Note Price = 100 − Upfront.
- Citi special case: Guaranteed Period = Non Callable Periods + 1.
- More than 5 underlyings triggers a warning (`Term` holds only 5).
- When the mail has no Product, `FCN` is used (for example, SG) and a warning is shown.
- When the mail has no Trade date, O5 is set to today.
- Missing Issue / Final / Maturity dates are calculated by the existing `Term` logic, with a warning.
- Initial price: Opening → 開盤價; Closing → 收盤價; VWAP → VWAP; any other value is left blank with a warning.
- `Term` events are suspended while writing, so mail dates are not overwritten by recalculation. Only D2 and P2 are re-triggered afterward, so the currency and tenor shown on `產品名稱` stay in sync.
- `Term!U2:V2`, `Q6`, and `O10` are not changed.

### Testing

Run these tests in a **copy** of the workbook (the repository's `.xlsm` already contains `OrderFillTest`; import `OrderFillTest.bas` into other workbooks first):

- `Test_OrderFill_AllIssuers`: pastes each of the 14 issuer sheets in `起單信件格式.xlsx`, runs the fill, and compares the result with the record for the same TDCC in ELN每日起單建檔; then runs the rule tests.
- `Test_OrderFill_Rules`: runs only the rule tests (no sample file needed), writing synthetic mail content into the paste area to check each fill rule above.
- `Test_OrderFill_Clear`: selects C3 and then presses **清空**, checking that no run-time error 1004 appears, pictures in the paste area are deleted, and the buttons and the C3 drop-down remain.

Known limitation: HSBC dates are `dd/mm/yyyy` text and are parsed day-first. If Excel has already converted such a value into a date when it was pasted, and both day and month are ≤ 12, the date may be wrong. Check it after filling.
