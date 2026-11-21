import { rows, appendRow, logAction } from "./excel-engine.js";
import { validateLimit } from "./platform-manager.js";
import { num } from "./cash-manager.js";

export function addTransaction(date, form) {
  const type = form.type;
  const isReceive = type === "receive";
  const isSend = type === "send";

  if (!isReceive && !isSend) {
    throw new Error("Invalid transaction type.");
  }

  const required = num(form.required);

  if (required <= 0) {
    throw new Error("Enter a valid amount greater than zero.");
  }

  let platform = isSend ? "NA" : form.platform;
  let received = isSend ? 0 : num(form.received);
  let refunded = isSend ? 0 : num(form.refunded);

  let net = 0;
  let due = 0;
  let limitUsage = 0;
  let status = "Settled Transaction";

  if (isReceive) {
    if (!form.policy?.trim()) {
      throw new Error("Enter the policy number or customer name.");
    }

    if (!platform || platform === "NA") {
      throw new Error("Select a platform.");
    }

    if (received < 0 || refunded < 0) {
      throw new Error("Received and refunded amounts cannot be negative.");
    }

    if (refunded > received) {
      throw new Error("Refunded amount cannot exceed received amount.");
    }

    if (!["Cash", "Other"].includes(form.mode)) {
  throw new Error("Select a valid payment mode.");
}

    net = received - refunded;
    due = Math.max(0, required - received);
    limitUsage = required;

    status = due > 0
      ? "Pending Due"
      : "Settled Transaction";

    const validation = validateLimit(date, platform, limitUsage);

    if (!validation.ok) {
      throw new Error(validation.message);
    }
  }

  if (isSend) {
    if (!form.headName?.trim()) {
      throw new Error("Enter the recipient/head name.");
    }

    // Money sent to the head is always a cash outflow.
    net = -required;

    // Head payments never consume a platform limit.
    limitUsage = 0;
    due = 0;
    received = 0;
    refunded = 0;
    status = "Settled Transaction";
  }

  const todays = rows("Transactions")
    .filter(r => String(r.Date) === date);

  const all = rows("Transactions");

  const id = `TX-${date.replaceAll("-", "")}-${String(all.length + 1).padStart(6, "0")}`;

  const record = {
    "Transaction ID": id,
    "Date": date,
    "Daily Serial": todays.length + 1,

    "Policy / Name": isSend
      ? form.headName.trim()
      : form.policy.trim(),

    "Details": form.details?.trim() || "",
    "Type": type,
    "Platform": platform,

    "Required Amount": required,
    "Received Amount": received,
    "Refunded Amount": refunded,

    "Net Cash Flow": net,
    "Balance Due": due,

    "Status": status,
    "Payment Mode": isReceive ? form.mode : "Cash",
    "Limit Usage": limitUsage,

    "Created At": new Date().toISOString()
  };

  appendRow("Transactions", record);

  logAction(
    "TRANSACTION_ADDED",
    id,
    `${type}; ${record["Policy / Name"]}; net ${net}`
  );

  return record;
}
