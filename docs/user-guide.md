# Quick user guide

1. Open the CashFlow Office application.
2. Select the office `.xlsx` workbook, or create a new workbook.
3. On a new business day, check the suggested opening cash and coin values against the physical count.
4. Enter P1, P2 and P3 daily limits, then enter the actual opening balance currently available in each platform. Do this afresh every business day; platform balances are not carried forward.
5. Record customer receipts. If RECD plus CHQ/QR exceeds the required amount, the app asks how much excess was returned. Customer refunds are always paid in physical cash. Choose **Returned** for the full amount or **Returned partially** and enter the amount actually paid. Any remaining refund appears in the cross-day **Pending customer refunds** list and can be paid on a later open business day using **Return cash**. The payment is recorded against the day cash is actually returned.
6. Use **Head Ledger** for both **Send to head** and **Receive from head**. Select Cash or Virtual; virtual transfers require a platform. A virtual send reduces that platform's available balance, while a virtual receipt and customer CHQ/QR collection increase it. Cash and virtual obligations remain separate, partial returns carry forward, and excess returns become advances that offset later transfers of the corresponding settlement type.
7. Use **Receipts & Payments** to manage manual office entries. Head transfers and receipts are included automatically in its totals.
8. Save after each entry. At day end, enter the physical closing cash and coin count and close the day.
9. Back up the updated workbook to a separate drive.

**Important:** If the browser downloads a new workbook instead of saving the selected file directly, replace the previous workbook with the downloaded copy. Always continue using the latest copy.


## Customer refunds: notes and coins

When an overpayment is returned, choose **Returned** or **Returned partially**. The next step asks for the total amount returned and the portion paid in coins. Enter `0` when no coins are included. The remaining refund amount is counted as notes. For later partial refunds, the same separate coin entry is required. The app deducts the notes portion from expected cash and the coin portion from expected coins on the date the refund is actually paid.

## Daily register numbering and Receipts / Payments

The daily register numbers only real receipt/payment movement rows; it no longer prints 91 numbered blank lines. Its NET balance runs across the actual entries. The bottom Receipts / Payments summary places receipt names and amounts on the left and payment amounts on the right, in the compact layout requested.


## Receiving a pending customer payment

Open **Transactions** and use **Customer receivables — pending payments**. The oldest outstanding transaction is selected first by default (FIFO), but you can choose another exact transaction from the dropdown. Confirm the customer name, original date, and transaction ID before saving. Enter the cash amount and/or CHQ/QR amount. Select the receiving platform if any CHQ/QR amount is entered. The app rejects a payment larger than the selected transaction's outstanding balance. The receipt is saved to `Customer Receipts` with today's active business date, while the original transaction date remains unchanged. Only the selected transaction's balance is reduced.
