import { rows, appendRow, logAction } from "./excel-engine.js";
import { num } from "./cash-manager.js";
import { validatePlatformBalance } from "./platform-manager.js";

const SHEET = "Head Ledger";
const makeId = () => `HL-${Date.now()}-${Math.random().toString(36).slice(2, 8).toUpperCase()}`;

export function getHeadLedger() {
  return rows(SHEET).sort((a, b) =>
    String(a.Date).localeCompare(String(b.Date)) ||
    String(a["Created At"]).localeCompare(String(b["Created At"]))
  );
}

function consumeLots(lots, amount) {
  let remaining = amount;
  while (remaining > 0.000001 && lots.length) {
    const lot = lots[0];
    const used = Math.min(lot.remaining, remaining);
    lot.remaining -= used;
    remaining -= used;
    if (lot.remaining <= 0.000001) lots.shift();
  }
  return remaining;
}

export function calculateHeadPosition(entries) {
  const positions = new Map();
  const get = name => {
    const key = name.toLocaleLowerCase();
    if (!positions.has(key)) positions.set(key, {
      head: name, cashReturnDueLots: [], virtualReturnDueLots: [],
      cashAdvanceLots: [], virtualAdvanceLots: [], cashSent: 0,
      cashReceived: 0, virtualSent: 0, virtualReceived: 0
    });
    return positions.get(key);
  };

  const ordered = [...entries].sort((a, b) =>
    String(a.Date).localeCompare(String(b.Date)) ||
    String(a["Created At"]).localeCompare(String(b["Created At"]))
  );

  for (const entry of ordered) {
    const name = String(entry.Head || "").trim();
    if (!name) continue;
    const b = get(name);
    const amount = num(entry.Amount);
    const date = String(entry.Date || "").slice(0, 10);
    const lot = (remaining, mode, sourceEntry) => ({
      head: name, date, mode, remaining,
      sourceId: sourceEntry["Entry ID"] || "",
      details: sourceEntry.Details || ""
    });

    if (entry.Action === "Send" && entry.Mode === "Cash") {
      b.cashSent += amount;
      const left = consumeLots(b.virtualAdvanceLots, amount);
      if (left > 0.000001) b.virtualReturnDueLots.push(lot(left, "Virtual", entry));
    } else if (entry.Action === "Send" && entry.Mode === "Virtual") {
      b.virtualSent += amount;
      const left = consumeLots(b.cashAdvanceLots, amount);
      if (left > 0.000001) b.cashReturnDueLots.push(lot(left, "Cash", entry));
    } else if (entry.Action === "Receive" && entry.Mode === "Cash") {
      b.cashReceived += amount;
      const excess = consumeLots(b.cashReturnDueLots, amount);
      if (excess > 0.000001) b.cashAdvanceLots.push(lot(excess, "Cash", entry));
    } else if (entry.Action === "Receive" && entry.Mode === "Virtual") {
      b.virtualReceived += amount;
      const excess = consumeLots(b.virtualReturnDueLots, amount);
      if (excess > 0.000001) b.virtualAdvanceLots.push(lot(excess, "Virtual", entry));
    }
  }
  // Net only the two opposing balances created by the head-exchange workflow.
  // Cash advance offsets virtual return due; virtual advance offsets cash return due.
  // Cash and virtual advances do NOT offset each other. These rules apply only to Heads.
  const netOpposingLots = (advanceLots, dueLots, advanceKind, dueKind) => {
    const advances = advanceLots.map(item => ({ ...item }));
    const dues = dueLots.map(item => ({ ...item }));
    let ai = 0;
    let di = 0;
    const allocations = [];
    while (ai < advances.length && di < dues.length) {
      const advance = advances[ai];
      const due = dues[di];
      const used = Math.min(advance.remaining, due.remaining);
      if (used > 0.000001) {
        advance.remaining -= used;
        due.remaining -= used;
        allocations.push({
          type: `${advanceKind} offset ${dueKind}`,
          amount: used,
          advanceDate: advance.date,
          advanceSourceId: advance.sourceId,
          dueDate: due.date,
          dueSourceId: due.sourceId,
          head: advance.head
        });
      }
      if (advance.remaining <= 0.000001) ai++;
      if (due.remaining <= 0.000001) di++;
    }
    return {
      advances: advances.filter(item => item.remaining > 0.000001),
      dues: dues.filter(item => item.remaining > 0.000001),
      allocations
    };
  };

  for (const b of positions.values()) {
    const cashAdvanceAgainstVirtualDue = netOpposingLots(
      b.cashAdvanceLots, b.virtualReturnDueLots, "Cash advance", "virtual return due"
    );
    b.cashAdvanceLots = cashAdvanceAgainstVirtualDue.advances;
    b.virtualReturnDueLots = cashAdvanceAgainstVirtualDue.dues;

    const virtualAdvanceAgainstCashDue = netOpposingLots(
      b.virtualAdvanceLots, b.cashReturnDueLots, "Virtual advance", "cash return due"
    );
    b.virtualAdvanceLots = virtualAdvanceAgainstCashDue.advances;
    b.cashReturnDueLots = virtualAdvanceAgainstCashDue.dues;

    // Recomputed deterministically from the append-only Head Ledger; dates and
    // source entry IDs remain available for an auditable allocation trail.
    b.crossOffsetAllocations = [
      ...cashAdvanceAgainstVirtualDue.allocations,
      ...virtualAdvanceAgainstCashDue.allocations
    ];
  }

  return [...positions.values()].sort((a, b) => a.head.localeCompare(b.head));
}

export function getHeadPositions() {
  return calculateHeadPosition(getHeadLedger());
}

export function getHeadBalances() {
  return getHeadPositions().map(position => ({
    head: position.head,
    cashReturnDue: position.cashReturnDueLots.reduce((sum, lot) => sum + lot.remaining, 0),
    virtualReturnDue: position.virtualReturnDueLots.reduce((sum, lot) => sum + lot.remaining, 0),
    cashAdvance: position.cashAdvanceLots.reduce((sum, lot) => sum + lot.remaining, 0),
    virtualAdvance: position.virtualAdvanceLots.reduce((sum, lot) => sum + lot.remaining, 0),
    cashSent: position.cashSent, cashReceived: position.cashReceived,
    virtualSent: position.virtualSent, virtualReceived: position.virtualReceived
  }));
}

export function addHeadEntry(date, form) {
  const head = String(form.head || "").trim();
  const action = String(form.action || "");
  const mode = String(form.mode || "");
  const amount = num(form.amount);
  const platform = mode === "Virtual" ? String(form.platform || "") : "NA";
  if (!head) throw new Error("Enter the head's name.");
  if (!["Send", "Receive"].includes(action)) throw new Error("Choose Send to Head or Receive from Head.");
  if (!["Cash", "Virtual"].includes(mode)) throw new Error("Choose Cash or Virtual.");
  if (!(amount > 0) || !Number.isFinite(amount)) throw new Error("Enter a valid amount greater than zero.");
  if (mode === "Virtual" && !["P1", "P2", "P3"].includes(platform)) throw new Error("Choose the platform used for this virtual transfer.");

  if (action === "Send" && mode === "Virtual") {
    const validation = validatePlatformBalance(date, platform, amount);
    if (!validation.ok) throw new Error(validation.message);
  }

  const id = makeId();
  appendRow(SHEET, {
    "Entry ID": id, Date: date, Head: head, Action: action, Mode: mode,
    Platform: platform, Amount: amount, Details: String(form.details || "").trim(),
    "Created At": new Date().toISOString()
  });
  logAction("HEAD_LEDGER_ENTRY", id, `${action} ${mode}; ${head}; ₹${amount}; ${platform}`);
  return id;
}
