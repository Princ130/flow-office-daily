
const SHEETS = {
  "Daily Summary": [
    "Date", "Opening Cash", "Opening Coins", "Closing Cash",
    "Closing Coins", "Expected Closing", "Cash Difference",
    "Total Cash In", "Total Cash Out", "Status", "Remarks"
  ],
  "Transactions": [
    "Transaction ID", "Date", "Daily Serial", "Policy / Name",
    "Details", "Type", "Platform", "Required Amount",
    "Received Amount", "Refunded Amount", "Net Cash Flow",
    "Balance Due", "Status", "Payment Mode", "Limit Usage", "Created At"
  ],
  "Platform Limits": [
    "Date", "Platform", "Daily Limit", "Used Limit", "Remaining Limit"
  ],
  "Denominations": [
    "Date", "Opening Notes", "Opening Coins", "Closing Notes", "Closing Coins"
  ],
  "Settings": ["Key", "Value"],
  "Audit Log": ["Timestamp", "Action", "Reference", "Details"]
};

let workbook = null;
const REGISTER_SHEET = "Daily Register";
let fileHandle = null;
let workbookName = "Office_Cash_Register.xlsx";

function sheetFromRows(data) {
  return XLSX.utils.aoa_to_sheet(data);
}

function ensureSheets(wb) {
  // Ensure all six original worksheets exist.
  for (const [name, headers] of Object.entries(SHEETS)) {
    if (!wb.Sheets[name]) {
      XLSX.utils.book_append_sheet(
        wb,
        sheetFromRows([headers]),
        name
      );
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

  const number = value => Number(value) || 0;
  const dateKey = value => String(value || "").slice(0, 10);
  const moneyFormat = '₹#,##0.00;[Red](₹#,##0.00)';

  const dates = [...new Set([
    ...dailyRecords.map(r => dateKey(r.Date)),
    ...transactions.map(t => dateKey(t.Date))
  ].filter(Boolean))].sort();

  const data = [];
  const merges = [];
  const styles = [];
  const rowHeights = {};

  const GREEN = "245B35";
  const DARK_GREEN = "183F29";
  const LIGHT_GREEN = "E2EFDA";
  const PALE_GREEN = "F3F8EF";
  const GOLD = "D9E8B5";
  const WHITE = "FFFFFF";
  const TEXT = "24352A";
  const MUTED = "647568";
  const BORDER = "C8D8C5";

  function addRow(values = []) {
    const row = Array(10).fill("");
    values.forEach((value, index) => {
      row[index] = value ?? "";
    });

    data.push(row);
    return data.length - 1;
  }

  function mergeRow(row, startCol, endCol) {
    merges.push({
      s: { r: row, c: startCol },
      e: { r: row, c: endCol }
    });
  }

  function mergePairs(row) {
    [[0, 1], [2, 3], [4, 5], [6, 7], [8, 9]]
      .forEach(([start, end]) => mergeRow(row, start, end));
  }

  function styleCell(row, col, style) {
    styles.push({ row, col, style });
  }

  function styleRange(row, startCol, endCol, style) {
    for (let col = startCol; col <= endCol; col++) {
      styleCell(row, col, style);
    }
  }

  function setHeight(row, height) {
    rowHeights[row] = height;
  }

  function formatDate(date) {
    if (!date) return "";

    const [year, month, day] = date.split("-").map(Number);

    return new Date(year, month - 1, day).toLocaleDateString("en-IN", {
      day: "2-digit",
      month: "long",
      year: "numeric"
    });
  }

  function moneyStyle(fill = WHITE, color = TEXT, size = 10) {
    return {
      font: {
        name: "Aptos",
        sz: size,
        bold: true,
        color: { rgb: color }
      },
      fill: { fgColor: { rgb: fill } },
      alignment: {
        horizontal: "center",
        vertical: "center"
      },
      numFmt: moneyFormat,
      border: {
        bottom: { style: "thin", color: { rgb: BORDER } }
      }
    };
  }

  dates.forEach((date, dayIndex) => {
    const day = dailyRecords.find(r => dateKey(r.Date) === date) || {};

    const dayTransactions = transactions
      .filter(t => dateKey(t.Date) === date)
      .sort((a, b) =>
        number(a["Daily Serial"]) - number(b["Daily Serial"])
      );

    const openingNotes = number(day["Opening Cash"]);
    const openingCoins = number(day["Opening Coins"]);

    const isClosed = day.Status === "Closed";

    const closingNotes = isClosed
      ? number(day["Closing Cash"])
      : "—";

    const closingCoins = isClosed
      ? number(day["Closing Coins"])
      : "—";

    const cashReceived = dayTransactions.reduce((sum, t) => {
      return sum + (
        t["Payment Mode"] === "Cash"
          ? Math.max(0, number(t["Net Cash Flow"]))
          : 0
      );
    }, 0);

    const cashPaidOut = dayTransactions.reduce((sum, t) => {
      return sum + (
        t["Payment Mode"] === "Cash"
          ? Math.max(0, -number(t["Net Cash Flow"]))
          : 0
      );
    }, 0);

    const currentCash = openingNotes + openingCoins
      + cashReceived - cashPaidOut;

    let runningBalance = openingNotes + openingCoins;

    // --------------------------------------------------
    // 1. MAIN TITLE
    // --------------------------------------------------

    let r = addRow(["CASHFLOW OFFICE"]);
    mergeRow(r, 0, 9);

    styleRange(r, 0, 9, {
      font: {
        name: "Aptos Display",
        sz: 20,
        bold: true,
        color: { rgb: WHITE }
      },
      fill: { fgColor: { rgb: DARK_GREEN } },
      alignment: {
        horizontal: "left",
        vertical: "center",
        indent: 1
      }
    });

    setHeight(r, 38);

    // --------------------------------------------------
    // 2. REGISTER SUBTITLE
    // --------------------------------------------------

    r = addRow(["DAILY TRANSACTION REGISTER"]);
    mergeRow(r, 0, 9);

    styleRange(r, 0, 9, {
      font: {
        name: "Aptos",
        sz: 12,
        bold: true,
        color: { rgb: WHITE }
      },
      fill: { fgColor: { rgb: GREEN } },
      alignment: {
        horizontal: "left",
        vertical: "center",
        indent: 2
      }
    });

    setHeight(r, 27);

    // --------------------------------------------------
    // 3. BUSINESS DATE / CURRENCY / STATUS
    // --------------------------------------------------

    r = addRow([
      "BUSINESS DATE", "",
      formatDate(date), "",
      "CURRENCY", "",
      "INDIAN RUPEE (₹)", "",
      "DAY STATUS", "",
    ]);

    mergeRow(r, 0, 1);
    mergeRow(r, 2, 3);
    mergeRow(r, 4, 5);
    mergeRow(r, 6, 7);
    mergeRow(r, 8, 9);

    styleRange(r, 0, 9, {
      font: {
        name: "Aptos",
        sz: 10,
        bold: true,
        color: { rgb: GREEN }
      },
      fill: { fgColor: { rgb: "EAF2E7" } },
      alignment: {
        horizontal: "center",
        vertical: "center"
      },
      border: {
        bottom: { style: "thin", color: { rgb: BORDER } }
      }
    });

    styleRange(r, 2, 3, {
      font: {
        name: "Aptos",
        sz: 11,
        bold: true,
        color: { rgb: DARK_GREEN }
      },
      fill: { fgColor: { rgb: "EAF2E7" } },
      alignment: { horizontal: "center", vertical: "center" }
    });

    styleRange(r, 8, 9, {
      font: {
        bold: true,
        color: {
          rgb: isClosed ? GREEN : "B7791F"
        }
      },
      fill: {
        fgColor: {
          rgb: isClosed ? LIGHT_GREEN : "FFF2CC"
        }
      },
      alignment: { horizontal: "center", vertical: "center" }
    });

    setHeight(r, 28);

    // --------------------------------------------------
    // 4. OPENING / CLOSING CASH AND COINS
    // --------------------------------------------------

    r = addRow([
      "OPENING NOTES", "",
      "OPENING COINS", "",
      "CLOSING NOTES", "",
      "CLOSING COINS", "",
      "DAY RECORD", ""
    ]);

    mergePairs(r);

    styleRange(r, 0, 9, {
      font: {
        name: "Aptos",
        sz: 9,
        bold: true,
        color: { rgb: MUTED }
      },
      fill: { fgColor: { rgb: "F1F5EF" } },
      alignment: { horizontal: "center", vertical: "center" }
    });

    setHeight(r, 23);

    r = addRow([
      openingNotes, "",
      openingCoins, "",
      closingNotes, "",
      closingCoins, "",
      day.Status || "No day record", ""
    ]);

    mergePairs(r);

    styleRange(r, 0, 7, moneyStyle("F8FBF6", DARK_GREEN, 12));

    styleRange(r, 8, 9, {
      font: {
        bold: true,
        sz: 11,
        color: { rgb: isClosed ? GREEN : "B7791F" }
      },
      fill: {
        fgColor: {
          rgb: isClosed ? LIGHT_GREEN : "FFF2CC"
        }
      },
      alignment: { horizontal: "center", vertical: "center" }
    });

    setHeight(r, 32);

    // --------------------------------------------------
    // 5. FINANCIAL SUMMARY
    // --------------------------------------------------

    r = addRow([
      "TOTAL REQUIRED", "",
      "TOTAL RECEIVED", "",
      "TOTAL REFUNDED", "",
      "NET CASH CHANGE", "",
      "CURRENT CASH IN HAND",
      "OUTSTANDING DUE"
    ]);

    mergeRow(r, 0, 1);
    mergeRow(r, 2, 3);
    mergeRow(r, 4, 5);
    mergeRow(r, 6, 7);

    styleRange(r, 0, 9, {
      font: {
        name: "Aptos",
        sz: 9,
        bold: true,
        color: { rgb: WHITE }
      },
      fill: { fgColor: { rgb: GREEN } },
      alignment: { horizontal: "center", vertical: "center" }
    });

    setHeight(r, 26);

    let totalRequired = 0;
    let totalReceived = 0;
    let totalRefunded = 0;
    let totalNet = 0;
    let totalDue = 0;

    dayTransactions.forEach(t => {
      totalRequired += number(t["Required Amount"]);
      totalReceived += number(t["Received Amount"]);
      totalRefunded += number(t["Refunded Amount"]);
      totalNet += number(t["Net Cash Flow"]);
      totalDue += number(t["Balance Due"]);
    });

    r = addRow([
      totalRequired, "",
      totalReceived, "",
      totalRefunded, "",
      totalNet, "",
      currentCash,
      totalDue
    ]);

    mergeRow(r, 0, 1);
    mergeRow(r, 2, 3);
    mergeRow(r, 4, 5);
    mergeRow(r, 6, 7);

    styleRange(r, 0, 7, moneyStyle("F3F8EF", DARK_GREEN, 13));

    styleCell(r, 8, moneyStyle(LIGHT_GREEN, GREEN, 13));
    styleCell(r, 9, moneyStyle("FFF0D9", "A45A00", 13));

    setHeight(r, 36);

    // --------------------------------------------------
    // 6. TRANSACTION TABLE HEADER
    // --------------------------------------------------

    addRow([]);

    r = addRow([
      "SERIAL #",
      "POLICY NUMBER / NAME",
      "TRANSACTION DETAILS",
      "AMOUNT REQUIRED",
      "AMOUNT RECEIVED",
      "AMOUNT REFUNDED",
      "NET CASH FLOW",
      "RUNNING BALANCE",
      "BALANCE DUE",
      "STATUS / REMARKS"
    ]);

    const tableHeaderRow = r;

    styleRange(r, 0, 9, {
      font: {
        name: "Aptos",
        sz: 10,
        bold: true,
        color: { rgb: WHITE }
      },
      fill: { fgColor: { rgb: GREEN } },
      alignment: {
        horizontal: "center",
        vertical: "center",
        wrapText: true
      },
      border: {
        top: { style: "medium", color: { rgb: DARK_GREEN } },
        bottom: { style: "medium", color: { rgb: DARK_GREEN } }
      }
    });

    setHeight(r, 38);

    // --------------------------------------------------
    // 7. TRANSACTION ENTRIES
    // --------------------------------------------------

    dayTransactions.forEach((t, index) => {
      const required = number(t["Required Amount"]);
      const received = number(t["Received Amount"]);
      const refunded = number(t["Refunded Amount"]);
      const net = number(t["Net Cash Flow"]);
      const due = number(t["Balance Due"]);

      const cashEffect = t["Payment Mode"] === "Cash" ? net : 0;
      runningBalance += cashEffect;

      const transactionLabel = t.Type === "send"
        ? "Send to Head"
        : "";

      const status = [
        t.Status || "",
        transactionLabel
      ].filter(Boolean).join(" · ");

      r = addRow([
        t["Daily Serial"] || index + 1,
        t["Policy / Name"] || "",
        t.Details || "",
        required,
        received,
        refunded,
        net,
        runningBalance,
        due,
        status
      ]);

      const rowFill = index % 2 === 0 ? WHITE : PALE_GREEN;

      styleRange(r, 0, 9, {
        font: {
          name: "Aptos",
          sz: 10,
          color: { rgb: TEXT }
        },
        fill: { fgColor: { rgb: rowFill } },
        alignment: {
          vertical: "center",
          horizontal: "left",
          wrapText: true
        },
        border: {
          bottom: { style: "thin", color: { rgb: BORDER } }
        }
      });

      styleCell(r, 0, {
        font: { bold: true, color: { rgb: GREEN } },
        fill: { fgColor: { rgb: rowFill } },
        alignment: { horizontal: "center", vertical: "center" },
        border: {
          bottom: { style: "thin", color: { rgb: BORDER } }
        }
      });

      styleCell(r, 1, {
        font: { bold: true, color: { rgb: TEXT } },
        fill: { fgColor: { rgb: rowFill } },
        alignment: { horizontal: "left", vertical: "center", wrapText: true },
        border: {
          bottom: { style: "thin", color: { rgb: BORDER } }
        }
      });

      styleCell(r, 2, {
        font: { color: { rgb: TEXT } },
        fill: { fgColor: { rgb: rowFill } },
        alignment: { horizontal: "left", vertical: "center", wrapText: true },
        border: {
          bottom: { style: "thin", color: { rgb: BORDER } }
        }
      });

      for (let c = 3; c <= 8; c++) {
        styleCell(r, c, {
          font: {
            name: "Aptos",
            sz: 10,
            color: { rgb: TEXT }
          },
          fill: { fgColor: { rgb: rowFill } },
          alignment: { horizontal: "right", vertical: "center" },
          numFmt: moneyFormat,
          border: {
            bottom: { style: "thin", color: { rgb: BORDER } }
          }
        });
      }

      // Highlight running balance.
      styleCell(r, 7, {
        font: { bold: true, color: { rgb: GREEN } },
        fill: { fgColor: { rgb: "E7F1E2" } },
        alignment: { horizontal: "right", vertical: "center" },
        numFmt: moneyFormat,
        border: {
          bottom: { style: "thin", color: { rgb: BORDER } }
        }
      });

      // Colour-coded transaction status.
      let statusFill = LIGHT_GREEN;
      let statusColor = GREEN;

      const lowerStatus = String(t.Status || "").toLowerCase();

      if (lowerStatus.includes("pending")) {
        statusFill = "FCE8C3";
        statusColor = "A45A00";
      } else if (lowerStatus.includes("refund")) {
        statusFill = "FCE4E4";
        statusColor = "B42318";
      }

      styleCell(r, 9, {
        font: {
          bold: true,
          sz: 9,
          color: { rgb: statusColor }
        },
        fill: { fgColor: { rgb: statusFill } },
        alignment: {
          horizontal: "center",
          vertical: "center",
          wrapText: true
        },
        border: {
          bottom: { style: "thin", color: { rgb: BORDER } }
        }
      });

      setHeight(r, 27);
    });

    // --------------------------------------------------
    // 8. EMPTY DAY MESSAGE
    // --------------------------------------------------

    if (dayTransactions.length === 0) {
      r = addRow(["NO TRANSACTIONS RECORDED FOR THIS BUSINESS DAY"]);
      mergeRow(r, 0, 9);

      styleRange(r, 0, 9, {
        font: {
          italic: true,
          color: { rgb: MUTED },
          sz: 10
        },
        fill: { fgColor: { rgb: "F1F5EF" } },
        alignment: {
          horizontal: "center",
          vertical: "center"
        }
      });

      setHeight(r, 30);
    }

    // --------------------------------------------------
    // 9. DAILY TOTAL
    // --------------------------------------------------

    r = addRow([
      "DAILY TOTAL", "", "",
      totalRequired,
      totalReceived,
      totalRefunded,
      totalNet,
      runningBalance,
      totalDue,
      ""
    ]);

    mergeRow(r, 0, 2);

    styleRange(r, 0, 9, {
      font: {
        name: "Aptos",
        sz: 10,
        bold: true,
        color: { rgb: WHITE }
      },
      fill: { fgColor: { rgb: DARK_GREEN } },
      alignment: {
        horizontal: "right",
        vertical: "center"
      },
      border: {
        top: { style: "medium", color: { rgb: DARK_GREEN } }
      }
    });

    styleCell(r, 0, {
      font: {
        bold: true,
        sz: 11,
        color: { rgb: WHITE }
      },
      fill: { fgColor: { rgb: DARK_GREEN } },
      alignment: { horizontal: "left", vertical: "center" },
      border: {
        top: { style: "medium", color: { rgb: DARK_GREEN } }
      }
    });

    for (let c = 3; c <= 8; c++) {
      styleCell(r, c, {
        font: { bold: true, color: { rgb: WHITE } },
        fill: { fgColor: { rgb: DARK_GREEN } },
        alignment: { horizontal: "right", vertical: "center" },
        numFmt: moneyFormat,
        border: {
          top: { style: "medium", color: { rgb: DARK_GREEN } }
        }
      });
    }

    setHeight(r, 32);

    // --------------------------------------------------
    // 10. CLOSING REMARKS
    // --------------------------------------------------

    if (day.Remarks) {
      r = addRow([`CLOSING REMARKS: ${day.Remarks}`]);
      mergeRow(r, 0, 9);

      styleRange(r, 0, 9, {
        font: {
          italic: true,
          color: { rgb: MUTED },
          sz: 10
        },
        fill: { fgColor: { rgb: "F1F5EF" } },
        alignment: {
          horizontal: "left",
          vertical: "center",
          wrapText: true,
          indent: 1
        }
      });

      setHeight(r, 30);
    }

    // --------------------------------------------------
    // 11. SEPARATION BETWEEN DAYS
    // --------------------------------------------------

    if (dayIndex < dates.length - 1) {
      addRow([]);
      setHeight(data.length - 1, 12);

      addRow([]);
      setHeight(data.length - 1, 12);
    }
  });

  if (dates.length === 0) {
    const r = addRow(["NO BUSINESS DAYS RECORDED YET"]);
    mergeRow(r, 0, 9);

    styleRange(r, 0, 9, {
      font: { bold: true, color: { rgb: GREEN }, sz: 12 },
      fill: { fgColor: { rgb: PALE_GREEN } },
      alignment: { horizontal: "center", vertical: "center" }
    });
  }

  // --------------------------------------------------
  // 12. CREATE AND FORMAT WORKSHEET
  // --------------------------------------------------

  const ws = XLSX.utils.aoa_to_sheet(data);

  ws["!merges"] = merges;

  ws["!cols"] = [
    { wch: 10 },
    { wch: 23 },
    { wch: 34 },
    { wch: 16 },
    { wch: 16 },
    { wch: 16 },
    { wch: 17 },
    { wch: 19 },
    { wch: 15 },
    { wch: 24 }
  ];

  ws["!rows"] = data.map((_, index) => ({
    hpt: rowHeights[index] || 23
  }));

  // Apply cell styles, including blank cells in merged areas.
  styles.forEach(({ row, col, style }) => {
    const address = XLSX.utils.encode_cell({
      r: row,
      c: col
    });

    if (!ws[address]) {
      ws[address] = { t: "s", v: "" };
    }

    ws[address].s = style;
  });

  // Keep Daily Register as the first worksheet.
  workbook.Sheets[REGISTER_SHEET] = ws;

  workbook.SheetNames = [
    REGISTER_SHEET,
    ...workbook.SheetNames.filter(name => name !== REGISTER_SHEET)
  ];

  workbook.Workbook = workbook.Workbook || {};
  workbook.Workbook.Views = [{ activeTab: 0 }];
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
