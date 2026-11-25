# Workbook structure

The application maintains these worksheets:

- Daily Register (generated daily register tabs)
- Daily Summary
- Transactions
- Platform Limits
- Denominations
- Settings
- Audit Log
- Daily Office Entries
- Head Ledger
- Refund Ledger

`Platform Limits` stores each day's daily limit and the manually entered opening balance for P1, P2 and P3. During the day, the displayed available balance is calculated from that day's opening figure, customer CHQ/QR receipts, and virtual Head Ledger movements. The next day requires fresh platform balance entry.

`Transactions` stores customer receipts, including total overpayment, cash actually refunded, and any refund still pending. `Refunded Amount` means cash actually returned to the customer; the daily register's `REFUND` records the actual customer refund paid and is not a formula for the non-cash portion as the cash refund field.

`Refund Ledger` is an append-only history of each cash refund payment, with the date the refund was actually returned and separate `Notes Amount` and `Coins Amount` fields. The `Amount` field remains the total. Notes reduce the cash-note balance; coins reduce the coin balance. This lets a pending refund be paid on a later business day without changing the original day's cash closing.

`Head Ledger` is an append-only event history with these columns:

`Entry ID | Date | Head | Action | Mode | Platform | Amount | Details | Created At`

The Head Ledger keeps Cash and Virtual transfers distinct. Sending Cash creates a virtual-return obligation; sending Virtual creates a cash-return obligation. Cash and virtual advances are kept separately and automatically offset future obligations of the matching settlement type.

`Daily Office Entries` stores manually entered Receipts & Payments items. Generated daily register tabs follow the supplied office format: opening/closing cash and coins at the top, platform daily limits at the right, the `SR NO` through `NET` columns on row 10, and the running balance in `NET`. `AMOUNT` is the gross receipt less any same-day customer refund, `RECD` is gross physical cash received, and `REFUND` is the actual customer refund paid on that same day, and `PAYMENT` is money sent out. Same-day customer refunds appear in the original receipt row under `REFUND` and are not duplicated in `PAYMENT`; refunds paid on a later date appear as a payment on the actual repayment date. Send transactions populate `PAYMENT`. The running balance starts from `SUM(C2:C6)` and continues as `previous NET + AMOUNT - PAYMENT`. Serial numbers are plain integers and appear only on actual movement rows; the register does not generate 91 numbered blank rows. The daily cash summary follows the actual entries. The compact Receipts / Payments area places receipt labels and amounts on the left and payment amounts on the right, side by side.

When an app-format workbook from an earlier release is opened, missing application columns plus the Head Ledger and Refund Ledger worksheets are added while retaining existing transaction rows. Keep an untouched backup before opening an older workbook.

Do not rename worksheet tabs or change header names unless the application code is updated accordingly.
