import { rows, appendRow, replaceRows, logAction } from "./excel-engine.js";
import { validateLimit } from "./platform-manager.js";
import { num } from "./cash-manager.js";

export function addTransaction(date, form) {
  const isReceive = form.type === "receive";
  const isSend = form.type === "send";
  if (!isReceive && !isSend) throw new Error("Invalid transaction type.");

  const required = num(form.required);
  if (required <= 0) throw new Error("Enter a valid amount greater than zero.");

  let platform = isSend ? "NA" : form.platform;
  let cashReceived = isSend ? 0 : num(form.cashReceived ?? form.received);
  let qrReceived = isSend ? 0 : num(form.qrReceived);
  let received = cashReceived + qrReceived;
  let refunded = isSend ? 0 : num(form.refundPaid);
  const refundCoins = isSend ? 0 : num(form.refundCoins);
  let overpayment = 0;
  let refundPending = 0;
  let net = 0;
  let due = 0;
  let limitUsage = 0;
  let status = "Settled Transaction";

  if (isReceive) {
    if (!form.policy?.trim()) throw new Error("Enter the policy number or customer name.");
    if (!platform || platform === "NA") throw new Error("Select a platform.");
    if (cashReceived < 0 || qrReceived < 0) throw new Error("Cash and CHQ/QR amounts cannot be negative.");
    if (received <= 0) throw new Error("Enter an amount received in cash or through CHQ/QR.");
    overpayment = Math.max(0, received - required);
    if (refunded < 0 || refunded > overpayment) throw new Error("Refunded cash cannot exceed the customer's overpayment.");
    if (refundCoins < 0 || refundCoins > refunded) throw new Error("Coins returned cannot exceed the total refund.");
    validateRefundCoinsAvailable(date, refundCoins);
    refundPending = overpayment - refunded;
    net = cashReceived - refunded;
    due = Math.max(0, required - received);
    limitUsage = required;
    status = refundPending > 0 ? (refunded > 0 ? "Partially Refunded" : "Refund Pending")
      : overpayment > 0 ? "Refunded Excess" : due > 0 ? "Pending Due" : "Settled Transaction";

    const validation = validateLimit(date, platform, limitUsage);
    if (!validation.ok) throw new Error(validation.message);
  }

  if (isSend) {
    if (!form.headName?.trim()) throw new Error("Enter the recipient/head name.");
    net = -required;
    due = 0;
    limitUsage = 0;
    received = 0;
    cashReceived = 0;
    qrReceived = 0;
    status = "Settled Transaction";
  }

  const todays = rows("Transactions").filter(row => String(row.Date) === date);
  const all = rows("Transactions");
  const id = `TX-${date.replaceAll("-", "")}-${String(all.length + 1).padStart(6, "0")}`;

  const record = {
    "Transaction ID": id,
    "Date": date,
    "Daily Serial": todays.length + 1,
    "Policy / Name": isSend ? form.headName.trim() : form.policy.trim(),
    "Details": form.details?.trim() || "",
    "Type": form.type,
    "Platform": platform,
    "Required Amount": required,
    "Received Amount": received,
    "Refunded Amount": refunded,
    "Net Cash Flow": net,
    "Balance Due": due,
    "Status": status,
    "Payment Mode": isSend ? "Cash" : (cashReceived > 0 && qrReceived > 0 ? "Split" : qrReceived > 0 ? "Other" : "Cash"),
    "Limit Usage": limitUsage,
    "Created At": new Date().toISOString(),
    "Cash Received": cashReceived,
    "CHQ/QR Received": qrReceived,
    "Overpayment Amount": overpayment,
    "Refund Pending": refundPending
  };

  appendRow("Transactions", record);
  if (refunded > 0) {
    appendRow("Refund Ledger", {
      "Refund ID": `RF-${id}-001`, Date: date, "Transaction ID": id,
      Customer: record["Policy / Name"], Amount: refunded,
      "Notes Amount": refunded - refundCoins, "Coins Amount": refundCoins,
      "Created At": new Date().toISOString()
    });
  }
  logAction("TRANSACTION_ADDED", id, `${form.type}; ${record["Policy / Name"]}; cash ${cashReceived}; CHQ/QR ${qrReceived}; cash refund ${refunded} (notes ${refunded - refundCoins}, coins ${refundCoins}); refund pending ${refundPending}`);
  return record;
}


function validateRefundCoinsAvailable(date, coins) {
  if (!(coins > 0)) return;
  const day = rows("Daily Summary").find(row => String(row.Date) === String(date));
  const openingCoins = num(day?.["Opening Coins"]);
  const alreadyReturned = rows("Refund Ledger")
    .filter(row => String(row.Date) === String(date))
    .reduce((sum, row) => sum + num(row["Coins Amount"]), 0);
  const available = Math.max(0, openingCoins - alreadyReturned);
  if (coins > available) throw new Error(`Only ₹${available.toFixed(2)} in opening coins is available to return today.`);
}

export function recordRefund(date, transactionId, amount, coins = 0) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date))) throw new Error("Choose a valid refund date.");
  const value = num(amount);
  const coinValue = num(coins);
  if (!(value > 0) || !Number.isFinite(value)) throw new Error("Enter a refund amount greater than zero.");
  if (!Number.isFinite(coinValue) || coinValue < 0 || coinValue > value) throw new Error("Enter a valid coin amount not exceeding the refund.");
  validateRefundCoinsAvailable(date, coinValue);
  const allTransactions = rows("Transactions");
  const target = allTransactions.find(row => row["Transaction ID"] === transactionId);
  if (!target) throw new Error("The transaction with pending refund could not be found.");
  const pending = num(target["Refund Pending"]);
  if (value > pending) throw new Error(`Refund cannot exceed the pending amount of ₹${pending.toFixed(2)}.`);

  const refundRows = rows("Refund Ledger");
  const refundId = `RF-${transactionId}-${String(refundRows.length + 1).padStart(4, "0")}`;
  appendRow("Refund Ledger", {
    "Refund ID": refundId, Date: date, "Transaction ID": transactionId,
    Customer: target["Policy / Name"], Amount: value,
    "Notes Amount": value - coinValue, "Coins Amount": coinValue,
    "Created At": new Date().toISOString()
  });

  const refunded = num(target["Refunded Amount"]) + value;
  const remaining = Math.max(0, pending - value);
  const sameDayRefunds = rows("Refund Ledger")
    .filter(row => row["Transaction ID"] === transactionId && String(row.Date) === String(target.Date))
    .reduce((sum, row) => sum + num(row.Amount), 0);
  const updatedRecord = {
    ...target,
    "Refunded Amount": refunded,
    "Refund Pending": remaining,
    "Net Cash Flow": num(target["Cash Received"]) - sameDayRefunds,
    Status: remaining > 0 ? "Partially Refunded" : "Refunded Excess"
  };
  replaceRows("Transactions", allTransactions.map(row => row["Transaction ID"] === transactionId ? updatedRecord : row));
  logAction("CUSTOMER_REFUND_RECORDED", transactionId, `Cash refund ₹${value} paid on ${date} (notes ₹${value - coinValue}, coins ₹${coinValue}); remaining ₹${remaining}`);
  return updatedRecord;
}


export function recordCustomerReceipt(date, transactionId, form) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date))) throw new Error("Choose a valid receipt date.");
  const allTransactions = rows("Transactions");
  const target = allTransactions.find(row => row["Transaction ID"] === transactionId);
  if (!target || target.Type !== "receive") throw new Error("Select a valid customer transaction.");
  const pending = num(target["Balance Due"]);
  if (!(pending > 0)) throw new Error("This customer transaction has no outstanding balance.");

  const cash = num(form.cashAmount);
  const qr = num(form.qrAmount);
  const total = Math.round((cash + qr) * 100) / 100;
  const platform = qr > 0 ? String(form.platform || "") : "NA";
  if (!Number.isFinite(cash) || !Number.isFinite(qr) || cash < 0 || qr < 0 || !(total > 0)) {
    throw new Error("Enter a valid cash and/or CHQ/QR receipt amount.");
  }
  if (total > pending + 0.000001) throw new Error(`Receipt cannot exceed this transaction's outstanding balance of ₹${pending.toFixed(2)}.`);
  if (qr > 0 && !["P1", "P2", "P3"].includes(platform)) throw new Error("Select the platform where the CHQ/QR payment was received.");

  const ledger = rows("Customer Receipts");
  const id = `CR-${date.replaceAll("-", "")}-${String(ledger.length + 1).padStart(5, "0")}`;
  const createdAt = new Date().toISOString();
  appendRow("Customer Receipts", {
    "Receipt ID": id, Date: date, "Transaction ID": transactionId,
    Customer: target["Policy / Name"] || "Unnamed customer", Amount: total,
    "Cash Amount": cash, "CHQ/QR Amount": qr, Platform: platform, "Created At": createdAt
  });

  const remaining = Math.max(0, Math.round((pending - total) * 100) / 100);
  const updated = {
    ...target,
    "Balance Due": remaining,
    Status: remaining > 0 ? "Pending Due" : (num(target["Refund Pending"]) > 0 ? "Refund Pending" : "Settled Transaction")
  };
  replaceRows("Transactions", allTransactions.map(row => row["Transaction ID"] === transactionId ? updated : row));
  logAction("CUSTOMER_DUE_RECEIVED", transactionId, `Receipt ${id}; received ₹${total} (cash ₹${cash}, CHQ/QR ₹${qr}, platform ${platform}); remaining due ₹${remaining}; origin ${target.Date}`);
  return { receiptId: id, transaction: updated, amount: total, remaining };
}
