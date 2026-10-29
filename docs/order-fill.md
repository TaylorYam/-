# 起單回填 / Order Mail Fill

> 本文件中英並列：上半為繁體中文，下半為英文。修改時請同步更新兩區。
>
> This document is bilingual: Traditional Chinese first, English below. Update both sections together.

[繁體中文](#繁體中文) | [English](#english)

## 繁體中文

把上手回覆的起單信件貼到「起單回填」工作表，按「填入」回填 `Term!A2:T2`、`Term!V2` 與 `Term!O5:O9`；按「清空」清掉貼上區，再貼下一封。支援 FCN、DACN、FCN_Express 的起單信。原始碼在 `src/vba/`。

### 安裝

repo 裡的 `特別約定事項&高齡適合度聲明書Generator_*.xlsm` 已內建 `OrderFill` 模組與「起單回填」工作表，開啟即可使用。測試模組 `OrderFillTest` 只內建在 `Generator_20261005.xlsm`。

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
  - 表格的數值列不當成左右一組。例：SG 條件表的 `Quote ?` 欄值是 `Strike`，不會把右邊的 Coupon 讀成 Strike。
- 同一欄在信裡出現兩個不同的值時留空，提醒會列出各處的值。
- 百分比看儲存格：數字格式有 `%`（信件貼上後 `80.00%` 會變成 0.8 加 `%` 格式），或文字帶 `%`，就是百分比數字（→ 80）。沒有 `%` 時照填；如果值看起來像小數（Strike／KO／KI／Note Price ≤ 3、Coupon < 1），會提醒確認。
- 數值欄位有逗號時不認，留空並提醒原文。例如 `105%, 100%`（列出多個值）、`98,50%`（逗號當小數點）。這些欄位都是百分比、月數或期數，不會用到千分位。
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
- Product 一律看信的內文，不看主旨。上手各自的寫法先轉換成 `Variable!D` 的商品名稱（不分大小寫）：Citi 的 `FCA` → `FCN`；`DAC`（BARC、BNP、CA、Nomura、Natixis 等）、`DRA`（DBS）、`VMRAN`（UBS）→ `DACN`；`Express FCN`（BARC、CA）→ `FCN_Express`。轉換後不在 `Variable!D` 清單上的，A2 留空並提醒原文。
- MS 的 DACN 信 Product 寫 `Range Accrual`：有 Strike（MS 的 `Put Strike`）時判定為 `DACN`，並提醒確認 A2；沒有 Strike 時照填 `Range Accrual`。其他上手的 `Range Accrual` 照填。
- `Accrual Barrier`（MS 的 DACN）在 `Term` 沒有對應欄位：和 Strike 不同時提醒，K2 仍填 Strike。
- FCN_Express：E2（Guaranteed Period）固定填 0，信裡有不是 0 的值時提醒。`First Period Rebate Multiplier`（BARC）／`Express Multiplier`（CA）在 `Term` 沒有對應欄位，值不是 1 時提醒。
- 信中沒有 Product 時不預設，A2 留空並提醒。只有 SG 例外：`AutoCall` 是單一數字、有 `Put Strike`，而且 `Fixed Coupons` 是 `All Periods` 時推定為 `FCN`、是 `First Period` 時推定為 `DACN`，並提醒確認。SG 的樣本只有一封 FCN 和幾封 DACN，無法確認 SG 其他類似產品（例如 `FCN_Stepdown`）的信不會也符合條件，所以推定後請一定核對 A2；如果改成其他產品，V2 的 0 不會自動清掉，請一併清空。
- Floor Coupon（V2）：A2 是 `FCN`、`DACN` 或 `FCN_Express`（包含 SG 推定、Citi 的 `FCA`、MS 判定的 DACN）時填 0；其他產品（包含 `FCN_Stepdown` 等變體）或 A2 留空時清空，避免上一筆的值留到下一筆。
- 信中沒有 Trade date 時 O5 填今天。
- Issue／Final／Maturity date 沒有時，沿用 `Term` 既有邏輯推算，並提醒。
- 期初價格：Opening → 開盤價；Closing → 收盤價；VWAP → VWAP；其他寫法留空並提醒。
- 回填時會暫停 `Term` 的事件，避免信裡的日期被重算覆蓋。之後只重新觸發 D2、P2，讓「產品名稱」的幣別與期間跟著更新。
- 不會動到 `Term!U2`、`Q6`、`O10`。

### 測試

在活頁簿**複本**執行以下測試（`Generator_20261005.xlsm` 已內建 `OrderFillTest`；其他活頁簿要先匯入 `OrderFillTest.bas`）：

- `Test_OrderFill_AllIssuers`：把 `起單信件格式.xlsx` 的 14 張上手表逐一貼入、填入，再和 ELN每日起單建檔中同一 TDCC 的紀錄比對；接著跑規則測試。
- `Test_OrderFill_Rules`：只跑規則測試（不需要樣本檔），在貼上區寫入合成的信件內容，逐條檢查上面的回填規則。
- `Test_OrderFill_Clear`：先選 C3 再按「清空」，確認不會出現執行階段錯誤 1004、貼上區的圖片會刪掉，按鈕與 C3 下拉清單保留。

已知限制：HSBC 的日期是 `dd/mm/yyyy` 文字，程式依日在前解析。如果 Excel 貼上時已自動轉成日期，而且日、月都 ≤ 12，可能會被轉錯，回填後請核對。

### 新增商品的全樣本驗證

起單回填每新增一種商品，或調整既有商品的規則，除了規則測試，還要用該商品 order sample 資料夾的**全部**信件驗證一次。order sample 從 Github-Cloud repo 下載到專案根目錄（已列在 `.gitignore`）。驗證程式在 `tools/sample_check/`，需要 Python 3（套件 `extract_msg`）與 Excel。

1. 抽出起單段落：`python -I -X utf8 tools/sample_check/extract.py "<資料夾>" <輸出資料夾>`
2. 逐封回填：`powershell -File tools/sample_check/run_fill.ps1 -Workbook <活頁簿> -HtmlDir <輸出資料夾>/html -OutFile <輸出資料夾>/result.txt`。程式會在活頁簿複本匯入 `src/vba/OrderFill.bas` 再回填，原活頁簿不會變更。
3. 比對：`python -I -X utf8 tools/sample_check/check.py <輸出資料夾>/manifest.json <輸出資料夾>/result.txt`

輸出資料夾要放在 repo 以外。輸出含客戶資料，不要 commit，也不要貼進 Issue 或 PR（repo 是公開的）。

貼法比照 `起單信件格式.xlsx` 裡各上手的 FCN 樣本，`extract.py` 會自動處理：

- 從玉山的 `We buy` 貼到簽名檔為止。沒有 `We buy` 的信不是起單內容，略過。例：BARC 的 BARX COMET 成交確認、索取 final TS 的信。
- HSBC：HSBC 回覆 final TS 時的日期區（`Strike Date`／`Issue Date`／`Final Valuation`／`Maturity Date`，左右一組）要一起貼。沒有這封回覆時，日期由 `Term` 推算。
- GS：GS 回覆裡的 `Order Summary`（`TDCC Code`／`Issue Date`／`Final Valuation Date`／`Maturity Date`）要一起貼。

`check.py` 自己讀信件表格算出應填值（不呼叫 VBA），再逐欄和回填結果比對。結果分成三類：

- 一致：也包含用到推算的信。信裡沒有日期、工具有提醒時，O5 填今天，O6～O8 由 `Term` 推算，這些欄位不算不一致。不寫 Trade date 的上手（CA、Natixis、Nomura、DBS、MS、UBS 等）屬於這類，維持現狀。
- 不一致：逐封看原信，判斷是工具錯、信件特殊，還是 `check.py` 的規則不足。已知不處理的情況：
  - 一封信兩筆單（兩張表，BNP、BARC）：不同的欄位留空並提醒。請一次只貼一張表。
  - 同一格拆成兩段（BNP，例：`MRVL`／`UW`）：只讀到前半段。
- 略過：不是起單內容。

新增商品時，先在 `check.py` 的 `expected()` 補上該商品各欄的規則（可參考 Super Bull 的寫法），再跑上面三個步驟。PR 要寫統計結果（一致〔其中用到推算〕、不一致、略過各幾封）與不一致的原因，但不要寫出 TDCC、價格、名目本金等交易資料。

已驗證的商品（2026-10-29）：

| 商品 | 信件 | 一致 | 其中用到推算 | 不一致 | 略過 |
|---|---|---|---|---|---|
| Super Bull | 142 | 135 | 42 | 7 | 0 |
| DACN | 85 | 81 | 48 | 0 | 4 |
| FCN_Express | 217 | 148 | 0 | 4 | 65 |

不一致的都是上面列的已知情況：Super Bull 是 BNP 一封信兩筆單 5 封、`MRVL`／`UW` 拆成兩段 2 封；FCN_Express 是 BARC 一封信兩筆單 4 封。

---

## English

Paste an issuer's order reply mail into the `起單回填` sheet and press **填入** (Fill) to populate `Term!A2:T2`, `Term!V2`, and `Term!O5:O9`. Press **清空** (Clear) to clear the paste area before pasting the next mail. Order mails for FCN, DACN, and FCN_Express are supported. Source code is in `src/vba/`.

### Installation

The repository's `特別約定事項&高齡適合度聲明書Generator_*.xlsm` already contains the `OrderFill` module and the `起單回填` sheet, so it works as soon as it is opened. The test module `OrderFillTest` is included only in `Generator_20261005.xlsm`.

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
  - A table's value row is not read as label/value pairs. For example, when SG's `Quote ?` column holds `Strike`, the Coupon to its right is not read as Strike.
- When a field appears twice in the mail with different values, it is left blank and the summary lists every value.
- Percentages follow the cell: a number with a `%` number format (pasting `80.00%` from a mail gives 0.8 formatted as `%`) or text containing `%` is a percentage figure (→ 80). A value without `%` is written as is; if it looks like a decimal (Strike / KO / KI / Note Price ≤ 3, Coupon < 1), a warning asks for a check.
- A numeric field containing a comma is not recognized; it is left blank and the warning shows the original text. Examples: `105%, 100%` (several values listed) and `98,50%` (comma as decimal point). These fields are percentages, months, or period counts, so they never need thousands separators.
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
- Product always comes from the mail body, never the subject. Each issuer's wording is first converted to a product name in `Variable!D` (case-insensitive): Citi's `FCA` → `FCN`; `DAC` (BARC, BNP, CA, Nomura, Natixis, and others), `DRA` (DBS), and `VMRAN` (UBS) → `DACN`; `Express FCN` (BARC, CA) → `FCN_Express`. A product that is still not in the `Variable!D` list after conversion leaves A2 blank, and the warning shows the original text.
- MS writes `Range Accrual` as the Product of a DACN mail: when there is a Strike (MS's `Put Strike`), the product is taken as `DACN` with a warning to confirm A2; without a Strike, `Range Accrual` is written as is. `Range Accrual` from other issuers is written as is.
- `Accrual Barrier` (MS's DACN) has no column in `Term`: a warning appears when it differs from Strike, and K2 still gets Strike.
- FCN_Express: E2 (Guaranteed Period) is always 0, with a warning when the mail has a non-zero value. `First Period Rebate Multiplier` (BARC) / `Express Multiplier` (CA) has no column in `Term`; a value other than 1 triggers a warning.
- When the mail has no Product, nothing is assumed: A2 is left blank with a warning. SG is the only exception: when `AutoCall` is a single number and there is a `Put Strike`, the product is inferred as `FCN` if `Fixed Coupons` is `All Periods` and as `DACN` if it is `First Period`, with a warning to confirm it. Only one SG FCN sample and a few SG DACN samples are available, so it is not known whether mails for similar SG products (such as `FCN_Stepdown`) also meet these conditions; always check A2 after an inference. If you change it to another product, the 0 in V2 is not cleared automatically, so clear it as well.
- Floor Coupon (V2): 0 when A2 is `FCN`, `DACN`, or `FCN_Express` (including the SG inference, Citi's `FCA`, and MS's DACN); cleared for any other product (including variants such as `FCN_Stepdown`) or when A2 is blank, so a previous trade's value is not carried over.
- When the mail has no Trade date, O5 is set to today.
- Missing Issue / Final / Maturity dates are calculated by the existing `Term` logic, with a warning.
- Initial price: Opening → 開盤價; Closing → 收盤價; VWAP → VWAP; any other value is left blank with a warning.
- `Term` events are suspended while writing, so mail dates are not overwritten by recalculation. Only D2 and P2 are re-triggered afterward, so the currency and tenor shown on `產品名稱` stay in sync.
- `Term!U2`, `Q6`, and `O10` are not changed.

### Testing

Run these tests in a **copy** of the workbook (`Generator_20261005.xlsm` already contains `OrderFillTest`; import `OrderFillTest.bas` into other workbooks first):

- `Test_OrderFill_AllIssuers`: pastes each of the 14 issuer sheets in `起單信件格式.xlsx`, runs the fill, and compares the result with the record for the same TDCC in ELN每日起單建檔; then runs the rule tests.
- `Test_OrderFill_Rules`: runs only the rule tests (no sample file needed), writing synthetic mail content into the paste area to check each fill rule above.
- `Test_OrderFill_Clear`: selects C3 and then presses **清空**, checking that no run-time error 1004 appears, pictures in the paste area are deleted, and the buttons and the C3 drop-down remain.

Known limitation: HSBC dates are `dd/mm/yyyy` text and are parsed day-first. If Excel has already converted such a value into a date when it was pasted, and both day and month are ≤ 12, the date may be wrong. Check it after filling.

### Full-sample check for a new product

Each time a product is added to the order fill, or the rules for an existing product change, run the rule tests and also check **every** mail in that product's order sample folder. The order samples are downloaded from the Github-Cloud repository into the project root (already listed in `.gitignore`). The check scripts are in `tools/sample_check/` and need Python 3 (with the `extract_msg` package) and Excel.

1. Extract the order sections: `python -I -X utf8 tools/sample_check/extract.py "<folder>" <output folder>`
2. Fill each mail: `powershell -File tools/sample_check/run_fill.ps1 -Workbook <workbook> -HtmlDir <output folder>/html -OutFile <output folder>/result.txt`. The script imports `src/vba/OrderFill.bas` into a copy of the workbook before filling; the original workbook is not changed.
3. Compare: `python -I -X utf8 tools/sample_check/check.py <output folder>/manifest.json <output folder>/result.txt`

Keep the output folder outside the repository. The output contains customer data: do not commit it or paste it into an Issue or PR (the repository is public).

The paste follows each issuer's FCN sample in `起單信件格式.xlsx`, and `extract.py` handles it automatically:

- Paste from E.SUN's `We buy` to the signature. A mail without `We buy` is not order content and is skipped, for example BARC's BARX COMET execution confirmations or a request for the final TS.
- HSBC: include the date block from HSBC's final TS reply (`Strike Date` / `Issue Date` / `Final Valuation` / `Maturity Date`, label and value side by side). Without that reply, the dates are calculated by `Term`.
- GS: include the `Order Summary` from GS's reply (`TDCC Code` / `Issue Date` / `Final Valuation Date` / `Maturity Date`).

`check.py` reads the mail table itself to work out the expected values (it does not call the VBA), then compares them cell by cell with the fill result. Results fall into three groups:

- Match: this includes mails that rely on calculated dates. When the mail has no date and the tool warns about it, O5 is today and O6 to O8 are calculated by `Term`; these cells do not count as mismatches. Issuers that do not send a Trade date (CA, Natixis, Nomura, DBS, MS, UBS, and others) fall here, and this stays as it is.
- Mismatch: read the original mail and decide whether the tool is wrong, the mail is unusual, or the `check.py` rules are incomplete. Known cases that are not handled:
  - Two trades in one mail (two tables, BNP and BARC): the fields that differ are left blank with a warning. Paste one table at a time.
  - One cell split into two paragraphs (BNP, for example `MRVL` / `UW`): only the first part is read.
- Skipped: not order content.

For a new product, first add its per-column rules to `expected()` in `check.py` (Super Bull is a good model), then run the three steps above. In the PR, report the counts (match [of which calculated], mismatch, skipped) and the reason for each mismatch, but do not include trade data such as TDCC, prices, or notional.

Products checked so far (2026-10-29):

| Product | Mails | Match | Of which calculated | Mismatch | Skipped |
|---|---|---|---|---|---|
| Super Bull | 142 | 135 | 42 | 7 | 0 |
| DACN | 85 | 81 | 48 | 0 | 4 |
| FCN_Express | 217 | 148 | 0 | 4 | 65 |

All mismatches are the known cases listed above: for Super Bull, 5 BNP mails with two trades and 2 with `MRVL` / `UW` split; for FCN_Express, 4 BARC mails with two trades.
