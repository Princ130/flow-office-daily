import { rows } from "./excel-engine.js";
export function getRecentTransactions(date, limit = 8) {
  return rows("Transactions").filter(row => String(row.Date) === date).slice(-limit).reverse();
}
