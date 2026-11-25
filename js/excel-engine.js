
const SHEETS = {
  "Daily Summary": [
    "Date", "Opening Cash", "Opening Coins", "Closing Cash",
    "Closing Coins", "Expected Closing", "Cash Difference",
    "Expected Cash", "Cash Difference Separate", "Expected Coins", "Coins Difference",
    "Total Cash In", "Total Cash Out", "Status", "Remarks"
  ],
  "Transactions": [
    "Transaction ID", "Date", "Daily Serial", "Policy / Name",
    "Details", "Type", "Platform", "Required Amount",
    "Received Amount", "Refunded Amount", "Net Cash Flow",
    "Balance Due", "Status", "Payment Mode", "Limit Usage", "Created At",
    "Cash Received", "CHQ/QR Received", "Overpayment Amount", "Refund Pending"
  ],
  "Platform Limits": [
    "Date", "Platform", "Daily Limit", "Used Limit", "Remaining Limit",
    "Opening Balance"
  ],
  "Denominations": [
    "Date", "Opening Notes", "Opening Coins", "Closing Notes", "Closing Coins"
  ],
  "Settings": ["Key", "Value"],
  "Audit Log": ["Timestamp", "Action", "Reference", "Details"],
  "Daily Office Entries": [
    "Entry ID", "Date", "Section", "Category", "Amount", "Details", "Created At", "Updated At"
  ],
  "Head Ledger": [
    "Entry ID", "Date", "Head", "Action", "Mode", "Platform", "Amount", "Details", "Created At"
  ],
  "Refund Ledger": [
    "Refund ID", "Date", "Transaction ID", "Customer", "Amount", "Notes Amount", "Coins Amount", "Created At"
  ],
  "Customer Receipts": [
    "Receipt ID", "Date", "Transaction ID", "Customer", "Amount", "Cash Amount", "CHQ/QR Amount", "Platform", "Created At"
  ]
};

let workbook = null;
const REGISTER_SHEET = "Daily Register";
let fileHandle = null;
let workbookName = "Office_Cash_Register.xlsx";

function sheetFromRows(data) {
  return XLSX.utils.aoa_to_sheet(data);
}

function ensureSheets(wb) {
  // Ensure every application worksheet exists and safely add newly introduced
  // columns to older workbooks without discarding existing rows or values.
  for (const [name, headers] of Object.entries(SHEETS)) {
    if (!wb.Sheets[name]) {
      XLSX.utils.book_append_sheet(wb, sheetFromRows([headers]), name);
      continue;
    }
    const existingAoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: "", raw: true });
    const existingHeaders = (existingAoa[0] || []).map(String).filter(Boolean);
    const mergedHeaders = [...existingHeaders, ...headers.filter(h => !existingHeaders.includes(h))];
    if (mergedHeaders.length !== existingHeaders.length) {
      const existingRows = existingAoa.slice(1);
      wb.Sheets[name] = sheetFromRows([
        mergedHeaders,
        ...existingRows.map(row => mergedHeaders.map((_, index) => row[index] ?? ""))
      ]);
    }
  }

  // Create the presentation register if missing.
  if (!wb.Sheets[REGISTER_SHEET]) {
    XLSX.utils.book_append_sheet(
      wb,
      sheetFromRows([
        ["CASHFLOW OFFICE"],
        ["Daily Register"],
        ["Daily Transaction Register"]
      ]),
      REGISTER_SHEET
    );
  }

  // Always make Daily Register the first tab.
  wb.SheetNames = [
    REGISTER_SHEET,
    ...wb.SheetNames.filter(name => name !== REGISTER_SHEET)
  ];
}

export function getWorkbook() {
  return workbook;
}

export function getWorkbookName() {
  return workbookName;
}

export function hasWorkbook() {
  return !!workbook;
}

// Create a workbook and retain its selected file handle.
export async function createWorkbook(handle = null) {
  workbook = XLSX.utils.book_new();

  for (const [name, headers] of Object.entries(SHEETS)) {
    XLSX.utils.book_append_sheet(
      workbook,
      sheetFromRows([headers]),
      name
    );
  }
  ensureSheets(workbook);
  console.log("Workbook worksheets:", workbook.SheetNames);
  fileHandle = handle;
  workbookName = handle?.name || "Office_Cash_Register.xlsx";
}

// Load a normal file upload.
export async function loadWorkbook(file) {
  const data = await file.arrayBuffer();

  workbook = XLSX.read(data, {
    type: "array",
    cellDates: false
  });

  ensureSheets(workbook);

  workbookName = file.name || workbookName;
  fileHandle = null;
}

// Load a file through the File System Access API.
// Preserve its handle so future saves update the same file.
export async function loadFromHandle(handle) {
  const file = await handle.getFile();
  const data = await file.arrayBuffer();

  workbook = XLSX.read(data, {
    type: "array",
    cellDates: false
  });

  ensureSheets(workbook);

  fileHandle = handle;
  workbookName = handle.name || workbookName;
}

function normalizeExcelDate(value) {
  if (value === "" || value === null || value === undefined) {
    return "";
  }

  // Already a JavaScript Date
  if (value instanceof Date) {
    return [
      value.getFullYear(),
      String(value.getMonth() + 1).padStart(2, "0"),
      String(value.getDate()).padStart(2, "0")
    ].join("-");
  }

  // Excel numeric date serial
  if (typeof value === "number" ||
      (typeof value === "string" && /^\d{5,6}$/.test(value))) {

    const parsed = XLSX.SSF.parse_date_code(Number(value));

    if (parsed) {
      return [
        parsed.y,
        String(parsed.m).padStart(2, "0"),
        String(parsed.d).padStart(2, "0")
      ].join("-");
    }
  }

  // Already stored as YYYY-MM-DD
  return String(value).slice(0, 10);
}

export function rows(sheetName) {
  const ws = workbook?.Sheets[sheetName];

  if (!ws) return [];

  const data = XLSX.utils.sheet_to_json(ws, {
    defval: "",
    raw: true
  });

  return data.map(row => ({
    ...row,
    ...(Object.hasOwn(row, "Date")
      ? { Date: normalizeExcelDate(row.Date) }
      : {})
  }));
}


function buildDailyRegister() {
  if (!workbook) return;

  const dailyRecords = rows("Daily Summary");
  const transactions = rows("Transactions");
  const platformRows = rows("Platform Limits");
  const officeEntries = rows("Daily Office Entries");
  const headEntries = rows("Head Ledger");
  const refundEntries = rows("Refund Ledger");
  const customerReceipts = rows("Customer Receipts");
  const number = value => Number(value) || 0;
  const dateKey = value => String(value || "").slice(0, 10);
  const moneyFormat = '₹#,##0.00;[Red](₹#,##0.00)';
  const baseSheets = new Set(Object.keys(SHEETS));
  const dates = [...new Set([
    ...dailyRecords.map(r => dateKey(r.Date)),
    ...transactions.map(t => dateKey(t.Date)),
    ...officeEntries.map(entry => dateKey(entry.Date)),
    ...headEntries.map(entry => dateKey(entry.Date)),
    ...refundEntries.map(entry => dateKey(entry.Date)),
    ...customerReceipts.map(entry => dateKey(entry.Date))
  ].filter(Boolean))].sort();
  const generatedDayNames = new Set(dates.map(date => {
    const [year, month, day] = date.split("-").map(Number);
    return `${String(day).padStart(2, "0")}${String(month).padStart(2, "0")}${year}`;
  }));

  // Remove only prior generated day tabs represented in the app's data, plus the old combined register.
  for (const name of [...workbook.SheetNames]) {
    if (name === REGISTER_SHEET || generatedDayNames.has(name)) {
      delete workbook.Sheets[name];
    }
  }
  workbook.SheetNames = workbook.SheetNames.filter(name =>
    name !== REGISTER_SHEET && !generatedDayNames.has(name)
  );

  const border = {
    top: { style: "thin", color: { rgb: "B7C9B8" } },
    bottom: { style: "thin", color: { rgb: "B7C9B8" } },
    left: { style: "thin", color: { rgb: "B7C9B8" } },
    right: { style: "thin", color: { rgb: "B7C9B8" } }
  };
  const titleStyle = { font: { name: "Aptos Display", sz: 16, bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "245B35" } }, alignment: { horizontal: "center", vertical: "center" } };
  const headerStyle = { font: { name: "Aptos", sz: 10, bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "245B35" } }, alignment: { horizontal: "center", vertical: "center", wrapText: true }, border };
  const labelStyle = { font: { name: "Aptos", sz: 10, bold: true, color: { rgb: "183F29" } }, fill: { fgColor: { rgb: "E2EFDA" } }, alignment: { vertical: "center" }, border };
  const valueStyle = { font: { name: "Aptos", sz: 10, color: { rgb: "24352A" } }, alignment: { horizontal: "right", vertical: "center" }, numFmt: moneyFormat, border };
  const totalStyle = { ...valueStyle, font: { name: "Aptos", sz: 10, bold: true, color: { rgb: "183F29" } }, fill: { fgColor: { rgb: "E2EFDA" } } };

  const dayName = date => {
    const [y, m, d] = date.split("-").map(Number);
    return `${String(d).padStart(2, "0")}${String(m).padStart(2, "0")}${y}`.slice(0, 31);
  };
  const col = n => XLSX.utils.encode_col(n);

  for (const date of dates) {
    const day = dailyRecords.find(r => dateKey(r.Date) === date) || {};
    const tx = transactions.filter(t => dateKey(t.Date) === date)
      .sort((a, b) => number(a["Daily Serial"]) - number(b["Daily Serial"]));
    const dayOfficeEntries = officeEntries.filter(entry => dateKey(entry.Date) === date);
    const dayHeadEntries = headEntries.filter(entry => dateKey(entry.Date) === date);
    const dayRefundEntries = refundEntries.filter(entry => dateKey(entry.Date) === date);
    const dayCustomerReceipts = customerReceipts.filter(entry => dateKey(entry.Date) === date);
    const limits = platformRows.filter(r => dateKey(r.Date) === date);
    const platformKeys = ["P1", "P2", "P3"];
    const platformNames = ["PLATFORM 1", "PLATFORM 2", "REMITR"];
    const openingCash = number(day["Opening Cash"]);
    const openingCoins = number(day["Opening Coins"]);
    const closingCash = day["Closing Cash"] === "" ? "" : number(day["Closing Cash"]);
    const closingCoins = day["Closing Coins"] === "" ? "" : number(day["Closing Coins"]);
    const cashIn = tx.reduce((sum, t) => {
      const explicit = t["Cash Received"];
      if (explicit !== "" && explicit !== undefined) return sum + number(explicit);
      return sum + (t["Payment Mode"] === "Cash" ? number(t["Received Amount"]) : 0);
    }, 0) + dayHeadEntries.filter(e => e.Action === "Receive" && e.Mode === "Cash").reduce((sum, e) => sum + number(e.Amount), 0)
      + dayCustomerReceipts.reduce((sum, entry) => sum + number(entry["Cash Amount"]), 0);
    const ledgerRefundByTx = new Map();
    refundEntries.forEach(e => ledgerRefundByTx.set(e["Transaction ID"], (ledgerRefundByTx.get(e["Transaction ID"]) || 0) + number(e.Amount)));
    const legacyRefunds = tx.reduce((sum, t) => sum + Math.max(0, number(t["Refunded Amount"]) - (ledgerRefundByTx.get(t["Transaction ID"]) || 0)), 0);
    const cashOut = dayRefundEntries.reduce((sum, e) => sum + number(e.Amount), 0) + legacyRefunds
      + tx.filter(t => t.Type === "send").reduce((sum, t) => sum + Math.abs(Math.min(0, number(t["Net Cash Flow"]))), 0)
      + dayHeadEntries.filter(e => e.Action === "Send" && e.Mode === "Cash").reduce((sum, e) => sum + number(e.Amount), 0);
    const expectedCash = openingCash + cashIn - cashOut;
    const expectedCoins = openingCoins;

    const customerCollection = tx.filter(t => t.Type === "receive")
      .reduce((sum, t) => sum + number(t["Required Amount"]), 0);
    const customerDueReceipts = dayCustomerReceipts.reduce((sum, entry) => sum + number(entry.Amount), 0);
    const headPayments = tx.filter(t => t.Type === "send")
      .reduce((sum, t) => sum + Math.abs(number(t["Net Cash Flow"])), 0)
      + dayHeadEntries.filter(e => e.Action === "Send").reduce((sum, e) => sum + number(e.Amount), 0);
    const headReceipts = dayHeadEntries.filter(e => e.Action === "Receive").reduce((sum, e) => sum + number(e.Amount), 0);
    const customerRefunds = dayRefundEntries.reduce((sum, e) => sum + number(e.Amount), 0) + legacyRefunds;
    const receiptItems = [];
    if (customerCollection > 0) {
      receiptItems.push({ label: "COLLECTION", amount: customerCollection });
    }
    dayCustomerReceipts.forEach(entry => receiptItems.push({ label: `CUSTOMER DUE — ${entry.Customer || entry["Transaction ID"]}`, amount: number(entry.Amount) }));
    dayHeadEntries.filter(entry => entry.Action === "Receive")
      .forEach(entry => receiptItems.push({ label: entry.Head || "RECEIVED FROM HEAD", amount: number(entry.Amount) }));
    dayOfficeEntries.filter(entry => entry.Section === "Receipt")
      .forEach(entry => receiptItems.push({ label: entry.Category, amount: number(entry.Amount) }));

    const paymentItems = [];
    const transactionDateById = new Map(transactions.map(t => [t["Transaction ID"], dateKey(t.Date)]));
    // Same-day customer refunds are shown in the original receipt row's REFUND
    // column and are already netted from AMOUNT. Only a refund paid on a later
    // date appears as a standalone cash payment on that later day's register.
    dayRefundEntries
      .filter(entry => transactionDateById.get(entry["Transaction ID"]) !== date)
      .forEach(entry => paymentItems.push({ amount: number(entry.Amount) }));
    tx.filter(t => t.Type === "send")
      .forEach(t => paymentItems.push({ amount: Math.abs(number(t["Net Cash Flow"])) }));
    dayHeadEntries.filter(entry => entry.Action === "Send")
      .forEach(entry => paymentItems.push({ amount: number(entry.Amount) }));
    dayOfficeEntries.filter(entry => entry.Section === "Payment")
      .forEach(entry => paymentItems.push({ amount: number(entry.Amount) }));

    // Same-day refunds stay on the original receipt row. A refund paid on a
    // later date gets a payment row on the actual repayment date.
    const sameDayTransactionIds = new Set(tx.map(t => t["Transaction ID"]));
    const separateRefundEntries = dayRefundEntries.filter(entry =>
      !sameDayTransactionIds.has(entry["Transaction ID"])
    );
    const cashMovements = [
      ...tx.map(t => ({ kind: "transaction", record: t, at: t["Created At"] || "", serial: number(t["Daily Serial"]) })),
      ...dayHeadEntries.filter(e => e.Mode === "Cash").map(e => ({ kind: "head", record: e, at: e["Created At"] || "", serial: 0 })),
      ...separateRefundEntries.map(e => ({ kind: "refund", record: e, at: e["Created At"] || "", serial: 0 })),
      ...dayCustomerReceipts.map(e => ({ kind: "customer-receipt", record: e, at: e["Created At"] || "", serial: 0 }))
    ].sort((a, b) => String(a.at).localeCompare(String(b.at)) || a.serial - b.serial);

    const firstMovementRow = 11;
    const lastMovementRow = firstMovementRow + cashMovements.length - 1;
    const dailySummaryTitleRow = Math.max(13, lastMovementRow + 2);
    const expectedCashRow = dailySummaryTitleRow + 1;
    const expectedCoinsRow = dailySummaryTitleRow + 2;
    const differencesRow = dailySummaryTitleRow + 3;
    const remarksRow = dailySummaryTitleRow + 4;
    const noteRow = dailySummaryTitleRow + 6;
    const rpHeaderRow = dailySummaryTitleRow + 8;
    const rpDataStart = rpHeaderRow + 1;
    const rpRows = Math.max(receiptItems.length, paymentItems.length, 1);
    const rpTotalRow = rpDataStart + rpRows;
    const lastRow = rpTotalRow;
    const aoa = Array.from({ length: lastRow }, () => Array(11).fill(""));
    const set = (r, c, v) => { aoa[r - 1][c - 1] = (typeof v === "string" && v.startsWith("=")) ? { f: v.slice(1) } : (v ?? ""); };

    // Original cash-register header and opening/closing fields.
    set(1, 1, new Date(`${date}T00:00:00`));
    set(2, 1, "OPENING CASH"); set(2, 3, openingCash);
    set(2, 8, "CLOSING CASH"); set(2, 9, closingCash);
    set(3, 1, "OPENING COINS"); set(3, 3, openingCoins);
    set(3, 8, "CLOSING COINS"); set(3, 9, closingCoins);

    for (let i = 0; i < 3; i++) {
      const platform = platformKeys[i];
      const item = limits.find(x => x.Platform === platform) || {};
      const dailyLimit = number(item["Daily Limit"]);
      const row = 4 + i;
      set(row, 8, "LIMIT");
      set(row, 9, dailyLimit);
      set(row, 10, platformNames[i]);
    }
    set(7, 3, "=SUM(C2:C6)");
    set(7, 8, "TOTAL");
    set(7, 9, "=SUM(I2:I6)");

    const headers = ["SR NO", "PARTICULARS", "AMOUNT", "RECD", "REFUND", "CHQ/QR", "PAYMENT", "NET"];
    headers.forEach((h, i) => set(10, i + 1, h));

    for (let i = 0; i < cashMovements.length; i++) {
      const excelRow = firstMovementRow + i;
      const movement = cashMovements[i];
      set(excelRow, 1, i + 1);
      if (movement.kind === "transaction") {
        const t = movement.record;
        const cashReceived = t["Cash Received"] !== "" && t["Cash Received"] !== undefined
          ? number(t["Cash Received"])
          : (t["Payment Mode"] === "Cash" ? number(t["Received Amount"]) : 0);
        const nonCash = t["CHQ/QR Received"] !== "" && t["CHQ/QR Received"] !== undefined
          ? number(t["CHQ/QR Received"])
          : (t["Payment Mode"] === "Other" ? number(t["Received Amount"]) : 0);
        set(excelRow, 2, [t["Policy / Name"], t.Details].filter(Boolean).join(" — "));
        const received = t.Type === "receive"
          ? (t["Received Amount"] !== "" && t["Received Amount"] !== undefined
            ? number(t["Received Amount"])
            : cashReceived + nonCash)
          : 0;
        const transactionRefundEntries = refundEntries.filter(e => e["Transaction ID"] === t["Transaction ID"]);
        const ledgerRefunded = transactionRefundEntries
          .reduce((sum, e) => sum + number(e.Amount), 0);
        const sameDayLedgerRefunded = transactionRefundEntries
          .filter(e => dateKey(e.Date) === date)
          .reduce((sum, e) => sum + number(e.Amount), 0);
        const legacyRefunded = Math.max(0, number(t["Refunded Amount"]) - ledgerRefunded);
        const sameDayRefunded = sameDayLedgerRefunded + legacyRefunded;
        // AMOUNT is the net amount retained from this receipt. RECD and
        // CHQ/QR preserve the gross tender split; REFUND records the actual
        // customer refund. Do not duplicate that refund in PAYMENT.
        set(excelRow, 3, t.Type === "receive" ? Math.max(0, received - sameDayRefunded) : 0);
        set(excelRow, 4, t.Type === "receive" ? cashReceived : 0);
        set(excelRow, 5, t.Type === "receive" ? sameDayRefunded : 0);
        set(excelRow, 6, t.Type === "receive" ? nonCash : 0);
        const payment = t.Type === "send" ? Math.abs(number(t["Net Cash Flow"])) : 0;
        set(excelRow, 7, payment);
      } else if (movement.kind === "head") {
        const e = movement.record;
        const receivedFromHead = e.Action === "Receive" ? number(e.Amount) : 0;
        const paymentToHead = e.Action === "Send" ? number(e.Amount) : 0;
        set(excelRow, 2, `${e.Action === "Send" ? "To" : "From"} Head — ${e.Head}`);
        set(excelRow, 3, receivedFromHead);
        set(excelRow, 4, receivedFromHead);
        set(excelRow, 5, `=C${excelRow}-D${excelRow}`);
        set(excelRow, 6, 0);
        set(excelRow, 7, paymentToHead);
      } else if (movement.kind === "refund") {
        const e = movement.record;
        set(excelRow, 2, `Customer Refund — ${e.Customer || e["Transaction ID"]}`);
        set(excelRow, 3, 0); set(excelRow, 4, 0);
        set(excelRow, 5, `=C${excelRow}-D${excelRow}`);
        set(excelRow, 6, 0); set(excelRow, 7, number(e.Amount));
      } else if (movement.kind === "customer-receipt") {
        const e = movement.record;
        set(excelRow, 2, `Customer Due Received — ${e.Customer || e["Transaction ID"]}`);
        set(excelRow, 3, number(e.Amount));
        set(excelRow, 4, number(e["Cash Amount"]));
        set(excelRow, 5, 0);
        set(excelRow, 6, number(e["CHQ/QR Amount"]));
        set(excelRow, 7, 0);
      }
      set(excelRow, 8, i === 0
        ? `=C7+C${excelRow}-G${excelRow}`
        : `=H${excelRow - 1}+C${excelRow}-G${excelRow}`);
    }

    const coinRefunds = dayRefundEntries.reduce((sum, entry) => sum + number(entry["Coins Amount"]), 0);
    const legacyCoinRefunds = 0;
    const notesRefunds = dayRefundEntries.reduce((sum, entry) => sum + (
      entry["Notes Amount"] !== "" && entry["Notes Amount"] !== undefined
        ? number(entry["Notes Amount"]) : number(entry.Amount)
    ), 0) + legacyRefunds;
    const cashSendOut = tx.filter(t => t.Type === "send")
      .reduce((sum, t) => sum + Math.abs(number(t["Net Cash Flow"])), 0)
      + dayHeadEntries.filter(e => e.Action === "Send" && e.Mode === "Cash")
        .reduce((sum, e) => sum + number(e.Amount), 0);
    const cashReceiveIn = tx.reduce((sum, t) => sum + (
      t["Cash Received"] !== "" && t["Cash Received"] !== undefined
        ? number(t["Cash Received"]) : (t["Payment Mode"] === "Cash" ? number(t["Received Amount"]) : 0)
    ), 0) + dayHeadEntries.filter(e => e.Action === "Receive" && e.Mode === "Cash")
      .reduce((sum, e) => sum + number(e.Amount), 0)
      + dayCustomerReceipts.reduce((sum, entry) => sum + number(entry["Cash Amount"]), 0);
    const expectedNotesCash = openingCash + cashReceiveIn - notesRefunds - cashSendOut;
    const expectedCoinBalance = openingCoins - coinRefunds - legacyCoinRefunds;

    set(dailySummaryTitleRow, 1, "DAILY CASH SUMMARY");
    set(expectedCashRow, 1, "EXPECTED CASH"); set(expectedCashRow, 3, expectedNotesCash);
    set(expectedCashRow, 5, "ACTUAL CASH"); set(expectedCashRow, 7, closingCash);
    set(expectedCoinsRow, 1, "EXPECTED COINS"); set(expectedCoinsRow, 3, expectedCoinBalance);
    set(expectedCoinsRow, 5, "ACTUAL COINS"); set(expectedCoinsRow, 7, closingCoins);
    set(differencesRow, 1, "CASH DIFFERENCE"); set(differencesRow, 3, closingCash === "" ? "" : `=I2-C${expectedCashRow}`);
    set(differencesRow, 5, "COINS DIFFERENCE"); set(differencesRow, 7, closingCoins === "" ? "" : `=I3-C${expectedCoinsRow}`);
    set(remarksRow, 1, "REMARKS"); set(remarksRow, 3, day.Remarks || "");
    set(noteRow, 1, "AMOUNT = GROSS RECEIPT − CUSTOMER REFUND | REFUND = ACTUAL CUSTOMER REFUND | NET = PREVIOUS NET + AMOUNT − PAYMENT");

    // Compact side-by-side Receipts / Payments view matching the supplied layout.
    set(rpHeaderRow, 2, "RECEIPTS");
    set(rpHeaderRow, 7, "PAYMENTS");
    for (let i = 0; i < rpRows; i++) {
      const row = rpDataStart + i;
      if (i === 0) {
        set(row, 1, new Date(`${date}T00:00:00`));
        if (receiptItems.length === 0) set(row, 2, "");
      }
      const receipt = receiptItems[i];
      if (receipt) {
        set(row, 2, receipt.label);
        set(row, 5, receipt.amount);
      }
      const payment = paymentItems[i];
      if (payment) set(row, 7, payment.amount);
    }
    set(rpTotalRow, 2, "TOTAL");
    set(rpTotalRow, 5, `=SUM(E${rpDataStart}:E${rpTotalRow - 1})`);
    set(rpTotalRow, 7, `=SUM(G${rpDataStart}:G${rpTotalRow - 1})`);

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = [
      { s: { r: dailySummaryTitleRow - 1, c: 0 }, e: { r: dailySummaryTitleRow - 1, c: 7 } },
      { s: { r: noteRow - 1, c: 0 }, e: { r: noteRow - 1, c: 7 } }
    ];
    ws["!cols"] = [
      { wch: 14 }, { wch: 22 }, { wch: 12 }, { wch: 10 }, { wch: 14 },
      { wch: 4 }, { wch: 14 }, { wch: 15 }, { wch: 12 }, { wch: 12 }, { wch: 12 }
    ];
    ws["!rows"] = Array.from({ length: lastRow }, (_, i) => ({
      hpt: i === 0 ? 30 : (i === 9 ? 28 : 21)
    }));

    if (ws.A1) {
      ws.A1.s = { font: { name: "Aptos", sz: 10, bold: true }, alignment: { horizontal: "left" } };
      ws.A1.z = "dd/mm/yyyy";
    }
    for (const r of [2, 3, 4, 5, 6]) {
      for (const c of [0, 7, 9]) {
        const address = XLSX.utils.encode_cell({ r: r - 1, c });
        if (ws[address]) ws[address].s = labelStyle;
      }
      for (const c of [2, 8]) {
        const address = XLSX.utils.encode_cell({ r: r - 1, c });
        if (ws[address]) ws[address].s = valueStyle;
      }
    }
    for (const address of ["C7", "I7"]) {
      if (ws[address]) ws[address].s = totalStyle;
    }
    for (let c = 0; c < 8; c++) {
      const address = XLSX.utils.encode_cell({ r: 9, c });
      if (ws[address]) ws[address].s = headerStyle;
    }
    for (let r = 10; r < 10 + cashMovements.length; r++) {
      for (let c = 0; c < 8; c++) {
        const address = XLSX.utils.encode_cell({ r, c });
        if (!ws[address]) ws[address] = { t: "s", v: "" };
        if (c === 0) {
          ws[address].s = { font: { name: "Aptos", sz: 10 }, alignment: { horizontal: "center", vertical: "center" }, border };
          ws[address].z = "0";
        } else if (c === 1) {
          ws[address].s = { alignment: { vertical: "center" }, border };
        } else {
          ws[address].s = valueStyle;
        }
      }
    }
    for (let r = 10; r < 10 + cashMovements.length; r++) {
      for (const c of [2, 3, 4, 5, 6, 7]) {
        const address = XLSX.utils.encode_cell({ r, c });
        if (ws[address]) ws[address].z = moneyFormat;
      }
    }
    for (let c = 0; c < 8; c++) {
      const address = XLSX.utils.encode_cell({ r: dailySummaryTitleRow - 1, c });
      if (!ws[address]) ws[address] = { t: "s", v: "" };
      ws[address].s = titleStyle;
      const noteAddress = XLSX.utils.encode_cell({ r: noteRow - 1, c });
      if (!ws[noteAddress]) ws[noteAddress] = { t: "s", v: "" };
      ws[noteAddress].s = { font: { name: "Aptos", sz: 9, italic: true, color: { rgb: "66736A" } } };
    }
    for (const r of [expectedCashRow, expectedCoinsRow, differencesRow]) {
      for (let c = 0; c < 8; c++) {
        const address = XLSX.utils.encode_cell({ r: r - 1, c });
        if (ws[address] && [0, 2, 4, 6].includes(c)) ws[address].s = totalStyle;
      }
    }
    for (const r of [expectedCashRow, expectedCoinsRow, differencesRow]) {
      for (const c of [2, 6]) {
        const address = XLSX.utils.encode_cell({ r: r - 1, c });
        if (ws[address]) ws[address].z = moneyFormat;
      }
    }
    for (let c = 0; c < 7; c++) {
      const address = XLSX.utils.encode_cell({ r: rpHeaderRow - 1, c });
      if (!ws[address]) ws[address] = { t: "s", v: "" };
      if (c === 1 || c === 6) ws[address].s = titleStyle;
    }
    for (let r = rpDataStart - 1; r < rpTotalRow; r++) {
      for (const c of [0, 1, 4, 6]) {
        const address = XLSX.utils.encode_cell({ r, c });
        if (!ws[address]) ws[address] = { t: "s", v: "" };
        ws[address].s = (r === rpTotalRow - 1) ? totalStyle
          : (c === 4 || c === 6 ? valueStyle : { alignment: { vertical: "center" }, border });
      }
      for (const c of [4, 6]) {
        const address = XLSX.utils.encode_cell({ r, c });
        if (ws[address]) ws[address].z = moneyFormat;
      }
    }
    if (ws[XLSX.utils.encode_cell({ r: rpDataStart - 1, c: 0 })]) {
      ws[XLSX.utils.encode_cell({ r: rpDataStart - 1, c: 0 })].z = "m/d/yy";
    }
    ws["!autofilter"] = { ref: `A10:H${Math.max(10, lastMovementRow)}` };
    ws["!freeze"] = { xSplit: 2, ySplit: 10 };
    ws["!pageSetup"] = { orientation: "landscape", fitToWidth: 1, fitToHeight: 0 };

    const name = dayName(date);
    if (workbook.Sheets[name]) {
      // Preserve a user sheet with a colliding date-like name rather than overwrite it.
      let suffix = 2;
      let unique = `${name}_${suffix}`;
      while (workbook.Sheets[unique]) unique = `${name}_${++suffix}`;
      workbook.Sheets[unique] = ws;
      workbook.SheetNames.push(unique);
    } else {
      workbook.Sheets[name] = ws;
      workbook.SheetNames.push(name);
    }
  }

  // Keep internal data sheets first, followed by daily tabs in date order.
  const internal = workbook.SheetNames.filter(name => baseSheets.has(name));
  const generated = workbook.SheetNames.filter(name => !baseSheets.has(name));
  workbook.SheetNames = [...generated, ...internal];
  workbook.Workbook = workbook.Workbook || {};
  workbook.Workbook.Views = [{ activeTab: 0 }];
  workbook.Workbook.CalcPr = { calcMode: "auto", fullCalcOnLoad: true, forceFullCalc: true };
}


export function replaceRows(
  sheetName,
  objects,
  headers = SHEETS[sheetName]
) {
  if (!workbook) {
    throw new Error("No workbook loaded");
  }

  if (!headers) {
    throw new Error(`Unknown worksheet: ${sheetName}`);
  }

  const data = [
    headers,
    ...objects.map(obj => headers.map(header => obj[header] ?? ""))
  ];

  workbook.Sheets[sheetName] = sheetFromRows(data);

  if (!workbook.SheetNames.includes(sheetName)) {
    workbook.SheetNames.push(sheetName);
  }
}

export function appendRow(
  sheetName,
  object,
  headers = SHEETS[sheetName]
) {
  const current = rows(sheetName);
  current.push(object);

  replaceRows(sheetName, current, headers);
}

export function updateRows(
  sheetName,
  updater,
  headers = SHEETS[sheetName]
) {
  replaceRows(sheetName, updater(rows(sheetName)), headers);
}

// Save to the same selected file whenever possible.
export async function saveWorkbook(silent = false) {
  if (!workbook) {
    throw new Error("No workbook loaded");
  }
  buildDailyRegister();

  const output = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "array"
  });

  const blob = new Blob([output], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });

  if (fileHandle && "createWritable" in fileHandle) {
    try {
      if (
        fileHandle.queryPermission &&
        await fileHandle.queryPermission({ mode: "readwrite" }) !== "granted"
      ) {
        const permission = await fileHandle.requestPermission({
          mode: "readwrite"
        });

        if (permission !== "granted") {
          throw new Error("Permission to update the workbook was not granted.");
        }
      }

      const writable = await fileHandle.createWritable();

      await writable.write(blob);
      await writable.close();

      return "direct";

    } catch (err) {
      console.warn("Direct save failed:", err);
      throw new Error(
        `Could not update ${workbookName}. Please check file permissions and try again. ${err.message}`
      );
    }
  }

  // Fallback for browsers without a writable file handle.
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");

  a.href = url;
  a.download = workbookName;
  document.body.appendChild(a);
  a.click();
  a.remove();

  setTimeout(() => URL.revokeObjectURL(url), 1500);

  return "download";
}

export async function chooseWorkbookHandle() {
  if (!window.showOpenFilePicker) return null;

  const [handle] = await window.showOpenFilePicker({
    multiple: false,
    types: [{
      description: "Excel workbook",
      accept: {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"]
      }
    }]
  });

  return handle || null;
}

export async function chooseSaveHandle() {
  if (!window.showSaveFilePicker) return null;

  return await window.showSaveFilePicker({
    suggestedName: workbookName,
    types: [{
      description: "Excel workbook",
      accept: {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"]
      }
    }]
  });
}

export function logAction(action, reference, details) {
  appendRow("Audit Log", {
    "Timestamp": new Date().toISOString(),
    "Action": action,
    "Reference": reference,
    "Details": details
  });
}

export function getSheetHeaders() {
  return SHEETS;
}
