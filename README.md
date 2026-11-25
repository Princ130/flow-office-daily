# CashFlow Office — Release 12

A browser-based, local-first cash register prototype for GitHub Pages. The Excel workbook is the persistent data store; no database or backend server is used.

## Run locally

Because the app uses JavaScript ES modules, do not open `index.html` directly as a `file://` URL. Use a local static server, for example:

```bash
python -m http.server 8000
```

Then visit `http://localhost:8000`.

## Deploy to GitHub Pages

1. Create a GitHub repository.
2. Upload the project files and folders (do not upload real office workbooks).
3. In repository Settings → Pages, deploy from the `main` branch and `/ (root)`.
4. Open the published Pages URL.

## Workbook workflow

- Choose **Open existing Excel workbook** to load a `.xlsx` file.
- Choose **Create a new workbook** to start with an empty in-browser workbook.
- Start each day by confirming physical opening cash/coins, entering the three platform limits, and recording the actual opening balance in each platform. Platform balances are entered fresh every business day and are not carried forward automatically.
- Record customer receipts. When combined RECD plus CHQ/QR exceeds the required amount, the app asks how much was returned. Customer refunds are always paid in cash; paid and pending amounts are tracked separately. The append-only `Refund Ledger` records the day the cash was actually returned, including later-day refunds.
- Use **Head Ledger** to record Send to Head and Receive from Head actions. Cash and virtual balances/obligations are separate per head, partial returns carry forward, and excess returns become advances.
- Use **Receipts & Payments** to manage manual office entries inside the app. Entries are stored in the `Daily Office Entries` worksheet and included in the generated daily register. Head transfers and receipts are included automatically.
- Generated daily tabs follow the original office register layout. `AMOUNT` is the amount retained after any same-day customer refund, `RECD` is gross cash received, `REFUND` records the actual same-day customer refund, and `PAYMENT` is money sent out (including refunds paid on a later date), and `NET` is the running balance.
- Save after each transaction or manual office entry. On supported Chrome/Edge browsers, the selected workbook can be saved directly if opened through the File System Access API. Otherwise, the app downloads an updated `.xlsx`; replace the previous workbook with that download before continuing.

## Important prototype notes

- Keep the workbook in a private office folder. Never commit it to GitHub.
- This version is intended for one active operator/workbook at a time. Do not let multiple computers edit separate copies simultaneously.
- Platform daily limits remain separate from head virtual balances. Actual platform balances are captured at day opening and adjusted during that day by recorded customer CHQ/QR receipts and virtual transfers to/from heads. The next business day always starts with fresh manually entered platform balances.
- The app has no authentication or encryption. Do not expose it to public/shared devices with sensitive data.
- The app does not yet migrate the older daily-tab-only office workbooks into its internal `Transactions` and `Daily Summary` data model. Load only an app-format workbook until the migration workflow is implemented; preserve a separate untouched copy of any legacy workbook.
- The daily serial is assigned in insertion order. The transaction ID is unique across workbook records.
- A closed day is intended to be read-only; the current prototype displays the closed state but does not yet implement a full admin reopening workflow.
- Back up the workbook daily to a separate drive.

## Folder structure

- `css/` stylesheets
- `js/` ES modules and application logic
- `assets/` optional images, icons and fonts
- `templates/` workbook template
- `docs/` operating and workbook documentation


## Release 6 note
Same-day customer refunds are consolidated into the originating customer receipt row in the daily register. Later-day refunds remain separate payment rows on their actual payment date.

## Day-closing outstanding summary

The Day Closing screen includes a live carry-forward view of unsettled balances. Red **To Give** includes pending customer refunds and virtual amounts still owed to heads. Green **To Receive** includes customer balances due and physical cash pending from heads. The lists and totals use the outstanding balances in the existing transaction and head ledgers; they do not create new entries or alter physical cash reconciliation.


## Release 9 additions

- **Missed-day recovery:** the earliest open business day is resumed automatically. The app blocks starting another day until the open day is reconciled and closed.
- **Dated outstanding summary:** customer dues/refunds, head obligations, and head advances display their original dates and age. Head settlements and offsets are allocated oldest-first (FIFO). The Head Ledger also nets opposing exchange positions: cash advances against virtual return due, and virtual advances against cash return due. Cash and virtual advances remain distinct from each other. This head-only netting never applies to customer balances.
- **No duplicate carry-forward rows:** the workbook's original transaction and Head Ledger events remain the source of truth; outstanding balances are recalculated from that history.

When closing a missed day, enter the actual notes/coins count as of that missed day's end. If transactions happened outside the app on a later date, reconcile those separately before entering the earlier closing count.

### Customer receivables (Release 10)

Use the **Customer receivables — pending payments** section on the Transactions page to see each original transaction that still has an amount due. Choose **Receive payment** for the exact transaction (or use **Record customer payment** and select it from the list), then enter the cash and/or CHQ/QR amount received. A QR portion requires a platform selection. The payment is applied only to that transaction; the original transaction date is preserved and the actual receipt date is stored in the `Customer Receipts` worksheet. The remaining due is reduced by the payment, and the receipt is included in that day's physical cash / platform balance and Receipts & Payments report without creating a second sale.
