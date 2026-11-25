import { rows, appendRow, replaceRows, logAction } from "./excel-engine.js";
import { num } from "./cash-manager.js";

const SHEET = "Daily Office Entries";

function makeId() {
  if (globalThis.crypto?.randomUUID) return `OFF-${globalThis.crypto.randomUUID()}`;
  return `OFF-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

function cleanForm(form) {
  const section = String(form.section || "").trim();
  const category = String(form.category || "").trim();
  const amount = num(form.amount);
  const details = String(form.details || "").trim();

  if (!["Receipt", "Payment"].includes(section)) {
    throw new Error("Choose Receipts or Payments.");
  }
  if (!category) throw new Error("Enter a category.");
  if (!(amount > 0) || !Number.isFinite(amount)) {
    throw new Error("Enter an amount greater than zero.");
  }
  if (category.length > 100 || details.length > 240) {
    throw new Error("Category or details are too long.");
  }

  return { section, category, amount, details };
}

export function getOfficeEntries(date) {
  return rows(SHEET)
    .filter(entry => String(entry.Date) === date)
    .sort((a, b) => String(a["Created At"]).localeCompare(String(b["Created At"])));
}

export function addOfficeEntry(date, form) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date))) {
    throw new Error("Choose a valid business date.");
  }
  const values = cleanForm(form);
  const id = makeId();
  appendRow(SHEET, {
    "Entry ID": id,
    Date: date,
    Section: values.section,
    Category: values.category,
    Amount: values.amount,
    Details: values.details,
    "Created At": new Date().toISOString()
  });
  logAction("OFFICE_ENTRY_ADDED", id, `${values.section}; ${values.category}; ₹${values.amount}`);
  return id;
}

export function updateOfficeEntry(id, form) {
  const values = cleanForm(form);
  let found = false;
  const updated = rows(SHEET).map(entry => {
    if (entry["Entry ID"] !== id) return entry;
    found = true;
    return {
      ...entry,
      Section: values.section,
      Category: values.category,
      Amount: values.amount,
      Details: values.details,
      "Updated At": new Date().toISOString()
    };
  });
  if (!found) throw new Error("The office entry could not be found.");
  replaceRows(SHEET, updated);
  logAction("OFFICE_ENTRY_UPDATED", id, `${values.section}; ${values.category}; ₹${values.amount}`);
}

export function deleteOfficeEntry(id) {
  const all = rows(SHEET);
  const target = all.find(entry => entry["Entry ID"] === id);
  if (!target) throw new Error("The office entry could not be found.");
  replaceRows(SHEET, all.filter(entry => entry["Entry ID"] !== id));
  logAction("OFFICE_ENTRY_DELETED", id, `${target.Section}; ${target.Category}; ₹${num(target.Amount)}`);
}

export function getOfficeSummary(date) {
  const transactions = rows("Transactions").filter(row => String(row.Date) === date);
  const entries = getOfficeEntries(date);
  const headEntries = rows("Head Ledger").filter(row => String(row.Date) === date);
  const refundEntries = rows("Refund Ledger");
  const dayRefunds = refundEntries.filter(row => String(row.Date) === date);
  const customerDueReceipts = rows("Customer Receipts").filter(row => String(row.Date) === date);
  const refundByTransaction = new Map();
  refundEntries.forEach(row => refundByTransaction.set(row["Transaction ID"], (refundByTransaction.get(row["Transaction ID"]) || 0) + num(row.Amount)));
  const legacyRefunds = transactions.reduce((sum, row) => sum + Math.max(0, num(row["Refunded Amount"]) - (refundByTransaction.get(row["Transaction ID"]) || 0)), 0);
  const customerCollection = transactions
    .filter(row => row.Type === "receive")
    .reduce((sum, row) => sum + num(row["Required Amount"]), 0);
  const headPayments = transactions
    .filter(row => row.Type === "send")
    .reduce((sum, row) => sum + Math.abs(num(row["Net Cash Flow"])), 0)
    + headEntries.filter(row => row.Action === "Send").reduce((sum, row) => sum + num(row.Amount), 0);
  const headReceipts = headEntries.filter(row => row.Action === "Receive")
    .reduce((sum, row) => sum + num(row.Amount), 0);
  const customerRefunds = dayRefunds.reduce((sum, row) => sum + num(row.Amount), 0) + legacyRefunds;
  const customerDueReceived = customerDueReceipts.reduce((sum, row) => sum + num(row.Amount), 0);
  const manualReceipts = entries
    .filter(entry => entry.Section === "Receipt")
    .reduce((sum, entry) => sum + num(entry.Amount), 0);
  const manualPayments = entries
    .filter(entry => entry.Section === "Payment")
    .reduce((sum, entry) => sum + num(entry.Amount), 0);

  return {
    customerCollection,
    headPayments,
    manualReceipts,
    manualPayments,
    headReceipts,
    customerRefunds,
    customerDueReceived,
    totalReceipts: customerCollection + customerDueReceived + headReceipts + manualReceipts,
    totalPayments: headPayments + customerRefunds + manualPayments,
    entries
  };
}
