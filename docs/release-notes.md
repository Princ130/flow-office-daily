# Release 7 — Refund recorded on the original receipt row

## Customer refund accounting

- When an overpayment is returned on the same business date, the refund is recorded in the original customer receipt row under `REFUND`.
- `AMOUNT` is the net amount retained from that receipt (`gross cash + CHQ/QR received - same-day refund`). `RECD` remains the gross physical cash received, and `CHQ/QR` remains the gross non-cash amount received.
- The same-day refund is not duplicated in `PAYMENT`; `NET` therefore increases by the retained amount only.
- Example: required ₹1,200, received ₹1,500, returned ₹300 produces `AMOUNT ₹1,200 | RECD ₹1,500 | REFUND ₹300 | CHQ/QR ₹0 | PAYMENT ₹0`, with a NET increase of ₹1,200.
- A refund paid on a later business date remains a separate payment movement on the actual repayment date. This preserves the cash-date audit trail and prevents rewriting an earlier day's closing balance.
- Partial refunds paid on the receipt date are netted in that original row; any later repayment is recorded on its actual date. Notes and coins remain separately tracked in the Refund Ledger.

## Daily register format alignment

- The generated daily tabs follow the supplied office register layout: date at A1, opening/closing cash and coins, platform limits at the right, and the original A:H transaction columns.
- Serial numbers use an integer format and appear only on actual movement rows. The register grows to the number of real entries instead of printing 91 numbered empty rows.
- The running NET starts from the opening total (`SUM(C2:C6)`) and continues as `previous NET + AMOUNT - PAYMENT`.
- The daily cash summary follows the compact register, and Receipts and Payments are displayed side by side in the requested seven-column layout.

## Customer overpayment refunds

- Overpayment is detected when combined RECD and CHQ/QR exceeds the required amount.
- All customer refunds are paid in physical cash, regardless of the original payment mix.
- When the operator marks the full or partial excess returned, the app requires the total refund and a separate coin amount (enter `0` if no coins were returned). The remaining amount is treated as notes.
- Entering a total refund of `0` leaves the excess pending. Notes refunds reduce the cash-note balance; coin refunds reduce the coin balance. The refund ledger stores both components separately while preserving the total refund amount.
- Pending refunds remain visible across business dates and can be paid later on an open business day.
- `Refund Ledger` records the actual payment date so later refunds reduce the cash balance on the day cash is returned, not the original transaction date.

## Head Ledger

- Send to Head and Receive from Head are recorded as append-only Head Ledger events.
- Each head has separate cash and virtual settlement balances.
- Cash transfers affect physical office cash; virtual transfers record the selected platform.
- Partial settlements carry forward. Excess returns become advances and automatically offset future obligations of the corresponding settlement type.
- The ledger history and current head balances are visible in the Head Ledger page.

## Daily platform balances

- The opening-day form now asks for the actual available balance of Platform 1, Platform 2, and Remitr every business day.
- These figures are not carried forward automatically. Each new day starts from the amounts entered by the operator.
- During the day, customer CHQ/QR receipts and virtual receipts from heads increase the selected platform balance; virtual sends to heads reduce it.
- Virtual sends are blocked when they exceed the available balance. Daily platform limits remain a separate control.

## Workbook compatibility

- Opening an earlier app-format workbook adds missing columns and the new Head Ledger and Refund Ledger worksheets while preserving existing transaction rows.
- Keep an untouched backup before upgrading a workbook.

## Verification performed

- JavaScript module syntax checks.
- HTML parse, duplicate-ID check, and JavaScript-to-HTML ID reference check.
- Mock-workbook accounting checks for cash-only refunds, QR-only overpayment refunds, later-day refunds, Head Ledger advances, physical cash totals, daily register generation, and app-workbook schema migration.
- Additional platform-balance checks: fresh daily opening amounts, virtual sends deducting the selected platform, virtual receipts and customer CHQ/QR receipts increasing it, overspend prevention, and completion of missing opening-balance setup on an already-open day.
- JavaScript module syntax, HTML parsing, unique IDs, and JavaScript-to-HTML ID references passed.
- No full interactive Chrome/Edge browser test was available in this build environment.

- Register-generation mock-workbook check: original A:H layout, opening total formula, actual-receipt amount mapping, payment-side sends, integer serial formatting, and prefilled running NET formulas.

## Release 8 — Day-closing outstanding summary

- Added a carry-forward summary to Day Closing with separate **To Give** and **To Receive** sections and distinct red/green styling.
- **To Give** lists pending customer refunds and virtual amounts the office still owes to heads.
- **To Receive** lists customer balances due and physical cash still due from heads.
- Each record shows the person, the type/date or head obligation, and the remaining unsettled amount. Totals show the outstanding amount in each direction.
- The summary is calculated from the existing Transactions, Refund Ledger-maintained transaction balances, and Head Ledger balances. It does not create duplicate accounting entries or change the cash closing calculation.


## Release 9 — Missed-day recovery and dated outstanding history

- On workbook open, the app checks the Daily Summary for any unclosed business day and automatically resumes the earliest open day instead of silently opening today's date.
- A missed-day warning identifies the date that is still open and reminds the operator to reconcile the prior day's actual notes and coins before starting a new business day.
- The app blocks creation of another business day while a different business day remains open. After closing an overdue day, it moves to the next open day or presents today's opening form.
- Closed days are read-only in the closing form.
- Outstanding customer refunds and customer balances due retain their original transaction dates. The Day Closing summary now displays the origin date and age in days. These are calculated from the original transaction records; no daily duplicate/copy-forward rows are created.
- Head obligations and advances are reconstructed from the append-only Head Ledger using FIFO allocation. Each remaining obligation/advance retains the date of the original transfer/receipt lot, even after partial settlement or partial offset.
- Cash advances and virtual advances remain separate and are displayed separately in Day Closing with their origin dates and ages. They are not counted as ordinary To Give/To Receive obligations.
- Existing workbook sheets and transaction history remain the source of truth; no new workbook migration sheet is required.

## Release 9 verification

- All JavaScript files pass `node --check`.
- Head-ledger FIFO test scenarios cover partial settlements, excess returns becoming dated advances, and advances offsetting later transfers without losing the original date of the remaining lot.
- HTML parsing, duplicate-ID checks, and JavaScript-to-HTML element reference checks pass.
- ZIP archive integrity passes.
- No full interactive Chrome/Edge test or real office workbook day-close simulation was available in this build environment.

## Release 10 — Customer Receivables

- Added a dedicated Customer Receivables section listing every transaction with an outstanding customer balance, including the original transaction date and transaction ID.
- Added a Record Customer Payment workflow. Each receipt must be assigned to one exact original transaction, preventing a payment from reducing another customer's balance.
- Supports cash, CHQ/QR, or split receipts. QR receipts are assigned to a selected platform and increase that platform's available balance; cash receipts increase physical cash.
- Added an append-only `Customer Receipts` worksheet containing receipt date, original transaction ID, customer, amount, cash/QR split, platform, and creation timestamp.
- Partial payments reduce only the selected transaction's `Balance Due`. Original transaction dates remain unchanged; receipt dates are stored separately.
- Added later customer due receipts to the daily cash register and Receipts & Payments summary without counting them as a new sale or adding them to the original customer collection again.
- Added a customer due received metric to Receipts & Payments.
- Existing workbooks are upgraded by adding the new worksheet and headers without removing existing transaction data.

## Release 10 verification

- JavaScript syntax checked across all application modules.
- HTML IDs checked for duplicates.
- Customer receipt ledger and per-transaction due reduction tested with two separate customer transactions.
- Cash receipt impact on physical cash and QR receipt impact on selected platform balance tested independently.
- Customer due receipts verified as separate receipts (not new customer sales) in the daily reporting logic.
- Full interactive browser and real Excel workbook export testing remains outstanding.


## Release 11 — Head exchange net-position correction

- Corrected the four Head Ledger summary balances so a head cannot simultaneously show a cash advance and virtual return due, or a virtual advance and cash return due.
- FIFO cross-offsets are calculated between **cash advance ↔ virtual return due** and **virtual advance ↔ cash return due**.
- Cash advance and virtual advance are intentionally kept separate; they are not netted against each other.
- The calculation is strictly Head Ledger logic. Customer receivables and refunds are unaffected.
- Original transaction/advance dates and source entry IDs remain on the residual lots. Cross-offset allocation details are deterministically reconstructed from the append-only Head Ledger, retaining both source dates and IDs. No source transactions are deleted or date-rewritten.


## Release 12 — Day Closing dashboard and Head transfer default

- Replaced the **Office Obligation — To give** card in the Day Closing carry-forward summary with **Head Advances — Available advances**.
- The advance card now occupies the first summary position beside **To receive**, and retains separate cash and virtual advance lists, offset descriptions, origin dates, and age.
- The total is the combined remaining cash and virtual advances; the underlying Head Ledger calculations and cross-offset rules are unchanged.
- In the **Send to Head** form, selecting **Virtual** now defaults the platform selector to **Remitr (P3)**. The operator may still choose another platform manually. Receive-from-Head platform selection is not automatically changed by this rule.

## Release 12 verification

- JavaScript modules pass `node --check`.
- HTML parses and contains no duplicate IDs.
- Dashboard DOM references were checked after moving the advance card.
- Verified the virtual Send-to-Head mode handler assigns P3 while preserving manual platform selection after the mode change.
- ZIP integrity passes. No interactive browser test or live workbook export test was performed.
