import { rows } from "./excel-engine.js";
const n = value => Number(value) || 0;

export function currentDay(date) {
  return rows("Daily Summary").find(row => String(row.Date) === date) || null;
}

export function dayTransactions(date) {
  return rows("Transactions").filter(row => String(row.Date) === date);
}

export function calculateDay(date) {
  const day = currentDay(date);
  if (!day) return {
    openingCash: 0, openingCoins: 0, cashIn: 0, cashOut: 0,
    expectedCash: 0, expectedCoins: 0, cashDifference: 0,
    coinsDifference: 0, opening: 0, balance: 0, pending: 0, transactions: []
  };

  const transactions = dayTransactions(date);
  const headEntries = rows("Head Ledger").filter(entry => String(entry.Date) === date);
  const refundEntries = rows("Refund Ledger");
  const dayRefunds = refundEntries.filter(entry => String(entry.Date) === date);
  const dayCustomerReceipts = rows("Customer Receipts").filter(entry => String(entry.Date) === date);
  const refundByTransaction = new Map();
  refundEntries.forEach(entry => refundByTransaction.set(entry["Transaction ID"], (refundByTransaction.get(entry["Transaction ID"]) || 0) + n(entry.Amount)));
  const legacyRefunds = transactions.reduce((sum, tx) => sum + Math.max(0, n(tx["Refunded Amount"]) - (refundByTransaction.get(tx["Transaction ID"]) || 0)), 0);
  const cashIn = transactions.reduce((sum, tx) => {
    // Refunds are always paid in physical cash, including QR overpayments.
    if (tx["Cash Received"] !== "" && tx["Cash Received"] !== undefined) {
      return sum + n(tx["Cash Received"]);
    }
    return sum + (tx["Payment Mode"] === "Cash" ? n(tx["Received Amount"]) : 0);
  }, 0) + headEntries.filter(entry => entry.Action === "Receive" && entry.Mode === "Cash")
    .reduce((sum, entry) => sum + n(entry.Amount), 0)
    + dayCustomerReceipts.reduce((sum, entry) => sum + n(entry["Cash Amount"]), 0);

  const cashOut = dayRefunds.reduce((sum, entry) => sum + (
      entry["Notes Amount"] !== "" && entry["Notes Amount"] !== undefined
        ? n(entry["Notes Amount"]) : n(entry.Amount)
    ), 0) + legacyRefunds
    + transactions.filter(tx => tx.Type === "send")
      .reduce((sum, tx) => sum + Math.abs(Math.min(0, n(tx["Net Cash Flow"]))), 0)
    + headEntries.filter(entry => entry.Action === "Send" && entry.Mode === "Cash")
      .reduce((sum, entry) => sum + n(entry.Amount), 0);

  const openingCash = n(day["Opening Cash"]);
  const openingCoins = n(day["Opening Coins"]);
  const expectedCash = openingCash + cashIn - cashOut;
  const coinOut = dayRefunds.reduce((sum, entry) => sum + n(entry["Coins Amount"]), 0);
  const expectedCoins = openingCoins - coinOut;
  const closingCash = day["Closing Cash"] === "" ? null : n(day["Closing Cash"]);
  const closingCoins = day["Closing Coins"] === "" ? null : n(day["Closing Coins"]);

  return {
    openingCash, openingCoins, cashIn, cashOut, coinOut, expectedCash, expectedCoins,
    cashDifference: closingCash === null ? 0 : closingCash - expectedCash,
    coinsDifference: closingCoins === null ? 0 : closingCoins - expectedCoins,
    opening: openingCash + openingCoins,
    balance: expectedCash + expectedCoins,
    pending: transactions.reduce((sum, tx) => sum + n(tx["Balance Due"]), 0),
    transactions,
    headEntries
  };
}

export const money = value => new Intl.NumberFormat("en-IN", {
  style: "currency", currency: "INR", maximumFractionDigits: 2
}).format(Number(value) || 0);
export const num = value => Number(value) || 0;
