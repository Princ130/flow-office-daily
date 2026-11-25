import { calculateDay, money, num } from "./cash-manager.js";
import { platformUsage } from "./platform-manager.js";
import { getRecentTransactions } from "./reports.js";
import { rows } from "./excel-engine.js";
import { getOfficeSummary } from "./office-summary.js";
import { getHeadBalances, getHeadLedger, getHeadPositions } from "./head-ledger.js";

let currentPage = 1, pageSize = 50;
const $ = id => document.getElementById(id);

export function toast(message) {
  const el = $("toast"); el.textContent = message; el.classList.add("show");
  setTimeout(() => el.classList.remove("show"), 2800);
}
export function showApp() { $("setupScreen").classList.add("hidden"); $("appShell").classList.remove("hidden"); }
export function showSetup() { $("appShell").classList.add("hidden"); $("setupScreen").classList.remove("hidden"); }
export function setWorkbookStatus(name) { $("workbookStatus").textContent = name; }
export function setDaySetupVisible(visible) { $("daySetupBanner").classList.toggle("hidden", !visible); }
export function setMissedDayBanner(date, visible) {
  const banner = $("missedDayBanner");
  if (!banner) return;
  banner.classList.toggle("hidden", !visible);
  if (visible) $("missedDayDate").textContent = new Date(`${date}T12:00:00`).toLocaleDateString("en-IN", { day: "numeric", month: "long", year: "numeric" });
}

export function fillOpeningSuggestions(date, suggestion, limits) {
  const form = $("openingForm");
  form.elements.businessDate.value = date;
  form.elements.openingCash.value = suggestion.cash;
  form.elements.openingCoins.value = suggestion.coins;
  form.elements.p1Limit.value = limits.P1 || "";
  form.elements.p2Limit.value = limits.P2 || "";
  form.elements.p3Limit.value = limits.P3 || "";
  form.elements.p1Balance.value = limits.balances?.P1 ?? "";
  form.elements.p2Balance.value = limits.balances?.P2 ?? "";
  form.elements.p3Balance.value = limits.balances?.P3 ?? "";
  $("carryForwardNote").textContent = suggestion.source
    ? `Suggested from ${suggestion.source}'s closing physical cash count. Enter today's actual platform balances separately; they are not carried forward.`
    : "First working day: enter the opening cash and coin count and each platform's actual balance.";
}

export function setActiveView(name) {
  document.querySelectorAll(".view").forEach(view => view.classList.add("hidden"));
  $(`${name}View`).classList.remove("hidden");
  document.querySelectorAll(".nav-link").forEach(button => button.classList.toggle("active", button.dataset.view === name));
  $("pageTitle").textContent = ({ dashboard: "Dashboard", transactions: "Transactions", limits: "Platform limits", closing: "Day closing", "office-summary": "Receipts & Payments", "head-ledger": "Head Ledger" })[name] || "Dashboard";
}

function badge(status) {
  const cls = ["Settled Transaction", "Refunded Excess"].includes(status) ? "status-settled" : ["Pending Due", "Refund Pending", "Partially Refunded"].includes(status) ? "status-pending" : "status-expense";
  return `<span class="status-tag ${cls}">${escapeHtml(status)}</span>`;
}
function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, char => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[char]));
}
function platformCard(platform) {
  const percent = platform.limit ? Math.min(100, platform.used / platform.limit * 100) : 0;
  const title = platform.platform === "P1" ? "Platform 1" : platform.platform === "P2" ? "Platform 2" : "Remitr";
  return `<article class="platform-card"><h4>${title}</h4><div class="platform-amounts"><span>Used ${money(platform.used)}</span><span>Limit ${money(platform.limit)}</span></div><div class="progress"><span class="${percent > 85 ? "warning" : ""}" style="width:${percent}%"></span></div><div class="platform-bottom"><span>${percent.toFixed(1)}% limit used</span><b>${money(platform.remaining)} limit left</b></div><div class="platform-bottom"><span>Available balance</span><b>${money(platform.balance)}</b></div></article>`;
}

export function renderAll(date) {
  $("todayLabel").textContent = new Date(`${date}T12:00:00`).toLocaleDateString("en-IN", { weekday: "short", day: "numeric", month: "short", year: "numeric" });
  const calc = calculateDay(date);
  $("cashBalance").textContent = money(calc.balance);
  $("cashIn").textContent = money(calc.cashIn);
  $("cashOut").textContent = money(calc.cashOut);
  $("pendingDues").textContent = money(calc.pending);
  const usage = platformUsage(date);
  $("platformCards").innerHTML = usage.map(platformCard).join("");
  $("limitsDetail").innerHTML = usage.map(platformCard).join("");

  const recent = getRecentTransactions(date);
  $("recentTransactions").innerHTML = recent.length
    ? recent.map(tx => `<tr><td>${tx["Daily Serial"]}</td><td>${escapeHtml(tx["Policy / Name"] || "—")}</td><td>${escapeHtml(tx.Platform)}</td><td>${money(tx["Net Cash Flow"])}</td><td>${badge(tx.Status)}</td></tr>`).join("")
    : `<tr><td colspan="5" class="empty-state">No transactions recorded today.</td></tr>`;

  renderTransactions(date);
  renderPendingRefunds();
  renderCustomerReceivables();
  renderOfficeSummary(date);
  renderHeadLedger();
  renderOutstandingSummary(date);
  $("expectedClosingCash").textContent = money(calc.expectedCash);
  $("expectedClosingCash").dataset.value = String(calc.expectedCash);
  $("expectedClosingCoins").textContent = money(calc.expectedCoins);
  $("expectedClosingCoins").dataset.value = String(calc.expectedCoins);
  updateClosingDifference();
}

export function changeTransactionPage(delta, date) {
  currentPage = Math.max(1, currentPage + delta); renderTransactions(date);
}

export function renderTransactions(date) {
  const search = $("transactionSearch").value.toLowerCase();
  const platformFilter = $("platformFilter").value;
  const statusFilter = $("statusFilter").value;
  let data = rows("Transactions").filter(row => String(row.Date) === date);
  data = data.filter(tx => (!search || [tx["Policy / Name"], tx.Details, tx["Transaction ID"]].join(" ").toLowerCase().includes(search))
    && (!platformFilter || tx.Platform === platformFilter) && (!statusFilter || tx.Status === statusFilter));
  const total = data.length, pages = Math.max(1, Math.ceil(total / pageSize));
  currentPage = Math.min(currentPage, pages);
  const slice = data.slice((currentPage - 1) * pageSize, currentPage * pageSize);
  $("transactionsBody").innerHTML = slice.length ? slice.map(tx => `<tr>
    <td>${tx["Daily Serial"]}</td><td>${escapeHtml(tx["Policy / Name"] || "—")}</td>
    <td>${escapeHtml(tx.Details)}</td><td>${escapeHtml(tx.Type)}</td><td>${escapeHtml(tx.Platform)}</td>
    <td>${money(tx["Required Amount"])}</td><td>${money(tx["Cash Received"] !== "" && tx["Cash Received"] !== undefined ? tx["Cash Received"] : (tx["Payment Mode"] === "Cash" ? tx["Received Amount"] : 0))}</td>
    <td>${money(tx["CHQ/QR Received"] !== "" && tx["CHQ/QR Received"] !== undefined ? tx["CHQ/QR Received"] : (tx["Payment Mode"] === "Other" ? tx["Received Amount"] : 0))}</td><td>${money(tx["Refunded Amount"])}</td><td>${money(tx["Net Cash Flow"])}</td>
    <td>${money(tx["Balance Due"])}</td><td>${badge(tx.Status)}${num(tx["Refund Pending"]) > 0 ? `<small class="muted">Refund pending ${money(tx["Refund Pending"])}</small>` : ""}</td>
    <td>${num(tx["Refund Pending"]) > 0 ? `<button class="button button-small button-outline" data-refund-id="${escapeHtml(tx["Transaction ID"])}">Return cash</button>` : "—"}</td></tr>`).join("")
    : `<tr><td colspan="13" class="empty-state">No matching transactions.</td></tr>`;
  $("recordCount").textContent = `${total} records`;
  $("pageInfo").textContent = `Page ${currentPage} of ${pages}`;
  $("prevPage").disabled = currentPage <= 1;
  $("nextPage").disabled = currentPage >= pages;
}


function ageDays(originDate, asOfDate) {
  if (!originDate || !asOfDate) return 0;
  const start = new Date(`${String(originDate).slice(0, 10)}T00:00:00`);
  const end = new Date(`${String(asOfDate).slice(0, 10)}T00:00:00`);
  return Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())
    ? 0 : Math.max(0, Math.floor((end - start) / 86400000));
}

function outstandingRows(items, emptyText, asOfDate) {
  if (!items.length) return `<p class="outstanding-empty">${escapeHtml(emptyText)}</p>`;
  return items.map(item => {
    const age = ageDays(item.date, asOfDate);
    const meta = `${item.meta} · Origin ${item.date || "date unavailable"} · ${age} day${age === 1 ? "" : "s"} old`;
    return `<div class="outstanding-row">
      <div class="outstanding-row-main"><strong>${escapeHtml(item.name)}</strong><small>${escapeHtml(meta)}</small></div>
      <strong class="outstanding-amount">${money(item.amount)}</strong>
    </div>`;
  }).join("");
}

export function renderOutstandingSummary(asOfDate = new Date().toISOString().slice(0, 10)) {
  const transactions = rows("Transactions");
  const customerGive = transactions
    .filter(tx => num(tx["Refund Pending"]) > 0)
    .map(tx => ({ name: tx["Policy / Name"] || "Unnamed customer", date: String(tx.Date).slice(0, 10), meta: `Customer refund · ${tx["Transaction ID"]}`, amount: num(tx["Refund Pending"]) }))
    .sort((a, b) => a.date.localeCompare(b.date));
  const customerReceive = transactions
    .filter(tx => num(tx["Balance Due"]) > 0)
    .map(tx => ({ name: tx["Policy / Name"] || "Unnamed customer", date: String(tx.Date).slice(0, 10), meta: `Customer balance due · ${tx["Transaction ID"]}`, amount: num(tx["Balance Due"]) }))
    .sort((a, b) => a.date.localeCompare(b.date));

  const positions = getHeadPositions();
  const headGive = positions.flatMap(position => position.virtualReturnDueLots
    .filter(lot => lot.remaining > 0.000001)
    .map(lot => ({ name: position.head, date: lot.date, meta: `Virtual amount to return${lot.details ? ` · ${lot.details}` : ""}`, amount: lot.remaining })));
  const headReceive = positions.flatMap(position => position.cashReturnDueLots
    .filter(lot => lot.remaining > 0.000001)
    .map(lot => ({ name: position.head, date: lot.date, meta: `Physical cash to receive${lot.details ? ` · ${lot.details}` : ""}`, amount: lot.remaining })));
  const headCashAdvances = positions.flatMap(position => position.cashAdvanceLots
    .filter(lot => lot.remaining > 0.000001)
    .map(lot => ({ name: position.head, date: lot.date, meta: "Cash advance available · offsets future virtual sends", amount: lot.remaining })));
  const headVirtualAdvances = positions.flatMap(position => position.virtualAdvanceLots
    .filter(lot => lot.remaining > 0.000001)
    .map(lot => ({ name: position.head, date: lot.date, meta: "Virtual advance available · offsets future cash sends", amount: lot.remaining })));

  const receiveTotal = [...customerReceive, ...headReceive].reduce((sum, item) => sum + item.amount, 0);
  $("outstandingReceiveTotal").textContent = money(receiveTotal);
  $("customerReceiveList").innerHTML = outstandingRows(customerReceive, "No customer balances due.", asOfDate);
  $("headReceiveList").innerHTML = outstandingRows(headReceive, "No cash pending from heads.", asOfDate);
  $("headCashAdvanceList").innerHTML = outstandingRows(headCashAdvances, "No cash advances available.", asOfDate);
  $("headVirtualAdvanceList").innerHTML = outstandingRows(headVirtualAdvances, "No virtual advances available.", asOfDate);
  $("headAdvanceTotal").textContent = money([...headCashAdvances, ...headVirtualAdvances].reduce((sum, item) => sum + item.amount, 0));
}

export function updateClosingDifference() {
  const form = $("closingForm"); if (!form) return;
  const actualCash = num(form.elements.closingCash.value);
  const actualCoins = num(form.elements.closingCoins.value);
  const expectedCash = num($("expectedClosingCash").dataset.value);
  const expectedCoins = num($("expectedClosingCoins").dataset.value);
  $("closingCashDifference").textContent = money(actualCash - expectedCash);
  $("closingCoinsDifference").textContent = money(actualCoins - expectedCoins);
}

export function showClosingForm(day) {
  const form = $("closingForm");
  form.elements.closingCash.value = day?.["Closing Cash"] ?? "";
  form.elements.closingCoins.value = day?.["Closing Coins"] ?? "";
  const locked = !day || day.Status === "Closed";
  form.querySelectorAll("input, textarea, button").forEach(control => { control.disabled = locked; });
}


export function renderOfficeSummary(date) {
  const summary = getOfficeSummary(date);
  const setText = (id, value) => { const el = $(id); if (el) el.textContent = money(value); };
  setText("automaticCustomerCollection", summary.customerCollection);
  setText("automaticCustomerDueReceipts", summary.customerDueReceived);
  setText("automaticHeadPayments", summary.headPayments);
  setText("automaticHeadReceipts", summary.headReceipts);
  setText("automaticCustomerRefunds", summary.customerRefunds);
  setText("manualReceiptsTotal", summary.manualReceipts);
  setText("manualPaymentsTotal", summary.manualPayments);
  setText("officeReceiptsTotal", summary.totalReceipts);
  setText("officePaymentsTotal", summary.totalPayments);

  const body = $("officeEntriesBody");
  if (!body) return;
  body.innerHTML = summary.entries.length
    ? summary.entries.map(entry => `<tr>
        <td>${escapeHtml(entry.Section)}</td>
        <td>${escapeHtml(entry.Category)}</td>
        <td>${money(entry.Amount)}</td>
        <td>${escapeHtml(entry.Details || "—")}</td>
        <td class="office-entry-actions">
          <button class="button button-small button-outline" data-office-edit="${escapeHtml(entry["Entry ID"])}">Edit</button>
          <button class="button button-small button-outline" data-office-delete="${escapeHtml(entry["Entry ID"])}">Delete</button>
        </td>
      </tr>`).join("")
    : `<tr><td colspan="5" class="empty-state">No manual Receipts & Payments entries for this day.</td></tr>`;
  $("officeEntryCount").textContent = `${summary.entries.length} manual entries`;
}


export function renderHeadLedger() {
  const balancesBody = $("headBalancesBody");
  const ledgerBody = $("headLedgerBody");
  if (!balancesBody || !ledgerBody) return;
  const balances = getHeadBalances();
  balancesBody.innerHTML = balances.length ? balances.map(b => `<tr>
    <td>${escapeHtml(b.head)}</td><td>${money(b.cashReturnDue)}</td><td>${money(b.virtualReturnDue)}</td>
    <td>${money(b.cashAdvance)}</td><td>${money(b.virtualAdvance)}</td></tr>`).join("")
    : `<tr><td colspan="5" class="empty-state">No head balances yet.</td></tr>`;
  const entries = getHeadLedger().slice().reverse();
  ledgerBody.innerHTML = entries.length ? entries.map(e => `<tr>
    <td>${escapeHtml(e.Date)}</td><td>${escapeHtml(e.Head)}</td><td>${escapeHtml(e.Action)}</td>
    <td>${escapeHtml(e.Mode)}</td><td>${escapeHtml(e.Platform === "NA" ? "—" : e.Platform)}</td>
    <td>${money(e.Amount)}</td><td>${escapeHtml(e.Details || "—")}</td></tr>`).join("")
    : `<tr><td colspan="7" class="empty-state">No Head Ledger entries recorded.</td></tr>`;
  $("headLedgerCount").textContent = `${entries.length} entries`;
}


export function renderCustomerReceivables() {
  const body = $("customerReceivablesBody");
  if (!body) return;
  const pending = rows("Transactions").filter(row => row.Type === "receive" && num(row["Balance Due"]) > 0)
    .sort((a, b) => String(a.Date).localeCompare(String(b.Date)) || String(a["Created At"]).localeCompare(String(b["Created At"])));
  const receipts = rows("Customer Receipts");
  body.innerHTML = pending.length ? pending.map(row => {
    const received = receipts.filter(item => item["Transaction ID"] === row["Transaction ID"])
      .reduce((sum, item) => sum + num(item.Amount), 0);
    return `<tr>
      <td>${escapeHtml(row.Date)}</td><td>${escapeHtml(row["Policy / Name"] || "Unnamed customer")}</td>
      <td>${escapeHtml(row["Transaction ID"])}</td><td>${money(num(row["Balance Due"]) + received)}</td>
      <td>${money(received)}</td><td>${money(row["Balance Due"])}</td>
      <td><button class="button button-small button-primary" data-customer-receipt-id="${escapeHtml(row["Transaction ID"])}">Receive payment</button></td></tr>`;
  }).join("") : `<tr><td colspan="7" class="empty-state">No customer payments are pending.</td></tr>`;
  const total = pending.reduce((sum, row) => sum + num(row["Balance Due"]), 0);
  const count = $("customerReceivablesCount");
  if (count) count.textContent = `${pending.length} pending · ${money(total)}`;
}

export function renderPendingRefunds() {
  const body = $("pendingRefundsBody");
  if (!body) return;
  const pending = rows("Transactions").filter(row => num(row["Refund Pending"]) > 0)
    .sort((a, b) => String(a.Date).localeCompare(String(b.Date)) || String(a["Created At"]).localeCompare(String(b["Created At"])));
  body.innerHTML = pending.length ? pending.map(row => {
    const paid = rows("Refund Ledger").filter(item => item["Transaction ID"] === row["Transaction ID"]);
    const notes = paid.reduce((sum, item) => sum + (
      item["Notes Amount"] !== "" && item["Notes Amount"] !== undefined
        ? num(item["Notes Amount"]) : num(item.Amount)
    ), 0);
    const coins = paid.reduce((sum, item) => sum + num(item["Coins Amount"]), 0);
    return `<tr>
      <td>${escapeHtml(row.Date)}</td><td>${escapeHtml(row["Policy / Name"])}</td>
      <td>${money(row["Overpayment Amount"])}</td><td>${money(row["Refunded Amount"])}</td>
      <td>${money(notes)}</td><td>${money(coins)}</td><td>${money(row["Refund Pending"])}</td>
      <td><button class="button button-small button-outline" data-refund-id="${escapeHtml(row["Transaction ID"])}">Return cash</button></td></tr>`;
  }).join("")
    : `<tr><td colspan="8" class="empty-state">No customer refunds are pending.</td></tr>`;
  $("pendingRefundsCount").textContent = `${pending.length} pending`;
}
