
import {
  createWorkbook,
  loadWorkbook,
  loadFromHandle,
  chooseWorkbookHandle,
  saveWorkbook,
  rows,
  getWorkbookName
} from "./excel-engine.js";

import {
  getSuggestedOpening,
  getOpenBusinessDays,
  startDay,
  closeDay
} from "./daily-manager.js";

import { addTransaction, recordRefund, recordCustomerReceipt } from "./transaction-manager.js";
import { addHeadEntry } from "./head-ledger.js";
import { setPlatformOpeningBalances } from "./platform-manager.js";
import {
  addOfficeEntry,
  updateOfficeEntry,
  deleteOfficeEntry,
  getOfficeEntries
} from "./office-summary.js";
import { currentDay } from "./cash-manager.js";

import {
  showApp,
  showSetup,
  setWorkbookStatus,
  setDaySetupVisible,
  setMissedDayBanner,
  fillOpeningSuggestions,
  setActiveView,
  renderAll,
  renderTransactions,
  renderPendingRefunds,
  renderCustomerReceivables,
  renderOfficeSummary,
  renderHeadLedger,
  changeTransactionPage,
  toast,
  updateClosingDifference,
  showClosingForm
} from "./ui.js";

import { saveSession, loadSession } from "./storage.js";

const $ = id => document.getElementById(id);

let activeDate = localDate();

function localDate() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function currentLimitsForDate(date) {
  const rec = rows("Platform Limits").filter(r => String(r.Date) === date);

  return {
    P1: rec.find(r => r.Platform === "P1")?.["Daily Limit"] || "",
    P2: rec.find(r => r.Platform === "P2")?.["Daily Limit"] || "",
    P3: rec.find(r => r.Platform === "P3")?.["Daily Limit"] || "",
    balances: {
      P1: rec.find(r => r.Platform === "P1")?.["Opening Balance"] ?? "",
      P2: rec.find(r => r.Platform === "P2")?.["Opening Balance"] ?? "",
      P3: rec.find(r => r.Platform === "P3")?.["Opening Balance"] ?? ""
    }
  };
}

async function persist() {
  const mode = await saveWorkbook(true);

  saveSession({
    date: activeDate,
    workbookName: getWorkbookName(),
    savedAt: new Date().toISOString()
  });

  if (mode === "download") {
    toast("Updated workbook downloaded. Replace your previous copy with this file.");
  } else {
    toast("Workbook saved successfully.");
  }
}

function enterApp() {
  showApp();
  setWorkbookStatus(getWorkbookName());

  const today = localDate();
  const openDays = getOpenBusinessDays();
  const pendingDay = openDays[0] || null;
  activeDate = pendingDay ? String(pendingDay.Date).slice(0, 10) : today;
  const day = currentDay(activeDate);
  const isMissedDay = activeDate < today && !!pendingDay;

  setMissedDayBanner(activeDate, isMissedDay);
  if (!day) {
    const suggestion = getSuggestedOpening(activeDate);
    fillOpeningSuggestions(activeDate, suggestion, currentLimitsForDate(activeDate));
    setDaySetupVisible(true);
  } else {
    const platformRows = rows("Platform Limits").filter(r => String(r.Date) === activeDate);
    const needsPlatformBalances = day.Status !== "Closed" && ["P1", "P2", "P3"].some(p => {
      const item = platformRows.find(r => r.Platform === p);
      return !item || item["Opening Balance"] === "" || item["Opening Balance"] === undefined;
    });
    if (needsPlatformBalances) {
      fillOpeningSuggestions(activeDate, {
        cash: day["Opening Cash"], coins: day["Opening Coins"], source: "this business day's saved opening count"
      }, currentLimitsForDate(activeDate));
      setDaySetupVisible(true);
    } else {
      setDaySetupVisible(false);
    }

    if (day.Status === "Closed") {
      toast("This business day is already closed. It is read-only.");
    } else if (isMissedDay) {
      toast(`Resumed the unclosed business day ${activeDate}. Close it before starting today.`);
    }
  }

  showClosingForm(day);
  renderAll(activeDate);
}

async function openFile(file) {
  await loadWorkbook(file);
  enterApp();
  toast("Workbook loaded.");
}

$("openWorkbookBtn").addEventListener("click", async () => {
  try {
    if (window.showOpenFilePicker) {
      const handle = await chooseWorkbookHandle();

      if (handle) {
        await loadFromHandle(handle);
        enterApp();
        toast("Workbook loaded.");
      }
    } else {
      $("workbookInput").click();
    }
  } catch (e) {
    if (e.name !== "AbortError") {
      toast(e.message || "Could not open workbook.");
    }
  }
});

$("workbookInput").addEventListener("change", async e => {
  const file = e.target.files[0];

  if (file) {
    try {
      await openFile(file);
    } catch (err) {
      toast(err.message);
    }
  }
});

// Create ONE workbook at the beginning.
$("createWorkbookBtn").addEventListener("click", async () => {
  try {
    let handle = null;

    if (window.showSaveFilePicker) {
      handle = await window.showSaveFilePicker({
        suggestedName: "Office_Cash_Register.xlsx",
        types: [{
          description: "Excel Workbook",
          accept: {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"]
          }
        }]
      });
    }

    await createWorkbook(handle);

    // Save the initial empty workbook to the selected location.
    await saveWorkbook();

    enterApp();
    toast("Workbook created successfully!");

  } catch (error) {
    if (error.name !== "AbortError") {
      toast(error.message || "Could not create workbook.");
    }
  }
});

$("saveWorkbookBtn").addEventListener("click", persist);

$("exportWorkbookBtn").addEventListener("click", async () => {
  try {
    await saveWorkbook(false);
    toast("Excel export created.");
  } catch (e) {
    toast(e.message);
  }
});

$("switchWorkbookBtn").addEventListener("click", () => {
  showSetup();
});

document.querySelectorAll(".nav-link").forEach(b => {
  b.addEventListener("click", () => {

    // Handle transaction actions
    if (b.dataset.action) {
      openTransaction(b.dataset.action);
      return;
    }

    // Handle normal page navigation
    setActiveView(b.dataset.view);

  });
});
document.querySelectorAll("[data-goto]").forEach(b => {
  b.addEventListener("click", () => setActiveView(b.dataset.goto));
});

document.querySelectorAll("[data-quick]").forEach(b => {
  b.addEventListener("click", () => openTransaction(b.dataset.quick));
});
document.querySelectorAll("[data-head-action]").forEach(b => {
  b.addEventListener("click", () => openHeadEntry(b.dataset.headAction));
});

$("newTransactionBtn").addEventListener("click", () => openTransaction("receive"));

function openHeadEntry(action = "Send") {
  const day = currentDay(activeDate);
  if (!day || day.Status === "Closed") {
    toast("Start an open business day before adding Head Ledger entries.");
    return;
  }
  const form = $("headEntryForm");
  form.reset();
  $("headAction").value = action;
  updateHeadEntryMode();
  $("headEntryDialog").showModal();
}

function updateHeadEntryMode() {
  const isVirtual = $("headMode").value === "Virtual";
  const isSend = $("headAction").value === "Send";
  const platformSelect = $("headEntryForm").elements.platform;
  $("headPlatformField").classList.toggle("hidden", !isVirtual);
  // For virtual sends to a Head, default the platform to Remitr (P3).
  // The user can still change it manually after the mode is selected.
  if (isVirtual && isSend) platformSelect.value = "P3";
  $("headEntryTitle").textContent = isSend ? "Send money to head" : "Receive money from head";
}
$("headMode").addEventListener("change", updateHeadEntryMode);
$("headAction").addEventListener("change", updateHeadEntryMode);
$("sendToHeadBtn").addEventListener("click", () => openHeadEntry("Send"));
$("receiveFromHeadBtn").addEventListener("click", () => openHeadEntry("Receive"));
$("closeHeadEntry").addEventListener("click", () => $("headEntryDialog").close());
$("cancelHeadEntry").addEventListener("click", () => $("headEntryDialog").close());
$("headEntryForm").addEventListener("submit", async e => {
  e.preventDefault();
  try {
    const values = Object.fromEntries(new FormData(e.currentTarget).entries());
    addHeadEntry(activeDate, values);
    await persist();
    $("headEntryDialog").close();
    renderAll(activeDate);
    toast("Head Ledger entry recorded.");
  } catch (err) {
    toast(err.message || "Could not save Head Ledger entry.");
  }
});

function openTransaction(type) {
  const day = currentDay(activeDate);

  if (!day || day.Status === "Closed") {
    toast("Start an open business day before adding transactions.");
    return;
  }

  $("transactionForm").reset();
  $("transactionType").value = type;

  updateTransactionFormMode();

  $("transactionDialog").showModal();
}

// Dynamically configure the form for customer receipts or head payments.
function updateTransactionFormMode() {
  const form = $("transactionForm");
  const isSend = $("transactionType").value === "send";
  const platform = form.elements.platform;
  const policy = form.elements.policy;
  const received = form.elements.received;
  const qrReceived = form.elements.qrReceived;
  const headName = form.elements.headName;
  const platformField = $("platformField");
  const policyField = $("policyField");
  const receivedField = $("receivedField");
  const qrReceivedField = $("qrReceivedField");
  const headNameField = $("headNameField");

  if (platformField) platformField.classList.toggle("hidden", isSend);
  if (policyField) policyField.classList.toggle("hidden", isSend);
  if (receivedField) receivedField.classList.toggle("hidden", isSend);
  if (qrReceivedField) qrReceivedField.classList.toggle("hidden", isSend);
  if (headNameField) headNameField.classList.toggle("hidden", !isSend);
  if (platform) platform.required = !isSend;
  if (policy) policy.required = !isSend;
  if (received) received.required = !isSend;
  if (qrReceived) qrReceived.required = false;
  if (headName) headName.required = isSend;
  if (isSend && platform) platform.value = "NA";
}
$("transactionType").addEventListener("change", updateTransactionFormMode);

$("closeTransactionDialog").addEventListener("click", () => {
  $("transactionDialog").close();
});

$("cancelTransaction").addEventListener("click", () => {
  $("transactionDialog").close();
});

let pendingRefundTransaction = null;

function showRefundDetails(defaultAmount = "") {
  $("refundChoiceActions").classList.add("hidden");
  $("partialRefundField").classList.remove("hidden");
  $("partialRefundAmount").value = defaultAmount === "" ? "" : String(defaultAmount);
  $("refundCoinsAmount").value = "";
  $("refundCoinsAmount").focus();
}

async function saveCustomerTransaction(formValues, refundPaid = 0, refundCoins = 0) {
  formValues.refundPaid = refundPaid;
  formValues.refundCoins = refundCoins;
  const record = addTransaction(activeDate, formValues);
  await persist();
  $("refundDialog").close();
  $("transactionDialog").close();
  pendingRefundTransaction = null;
  renderAll(activeDate);
  toast(`Transaction #${record["Daily Serial"]} recorded.`);
}

$("transactionForm").addEventListener("submit", async e => {
  e.preventDefault();
  try {
    const f = Object.fromEntries(new FormData(e.currentTarget).entries());
    f.cashReceived = f.received || 0;
    f.qrReceived = f.qrReceived || 0;
    const totalReceived = Number(f.cashReceived || 0) + Number(f.qrReceived || 0);
    const excess = Math.max(0, totalReceived - Number(f.required || 0));
    if (f.type === "receive" && excess > 0) {
      pendingRefundTransaction = f;
      $("refundPromptText").textContent = `The customer paid ₹${excess.toFixed(2)} more than required. If returning any amount now, enter the total cash refund and separately specify how much of it is coins.`;
      $("partialRefundField").classList.add("hidden");
      $("refundChoiceActions").classList.remove("hidden");
      $("partialRefundAmount").value = "";
      $("partialRefundAmount").max = String(excess);
      $("refundCoinsAmount").value = "";
      $("refundDialog").showModal();
      return;
    }
    await saveCustomerTransaction(f, 0, 0);
  } catch (err) {
    toast(err.message || "Could not save transaction.");
  }
});

function cancelRefundPrompt() {
  pendingRefundTransaction = null;
  $("refundDialog").close();
}
$("cancelRefundPrompt").addEventListener("click", cancelRefundPrompt);
$("cancelRefundPromptMain").addEventListener("click", cancelRefundPrompt);

$("refundFullyBtn").addEventListener("click", () => {
  if (!pendingRefundTransaction) return;
  const max = pendingRefundTransaction.existingId
    ? pendingRefundTransaction.pendingAmount
    : Number(pendingRefundTransaction.cashReceived || 0) + Number(pendingRefundTransaction.qrReceived || 0) - Number(pendingRefundTransaction.required || 0);
  showRefundDetails(max);
});
$("refundPartiallyBtn").addEventListener("click", () => showRefundDetails(""));

$("confirmPartialRefund").addEventListener("click", async () => {
  if (!pendingRefundTransaction) return;
  const rawRefundAmount = $("partialRefundAmount").value;
  const rawCoinsAmount = $("refundCoinsAmount").value;
  if (rawRefundAmount === "") {
    toast("Enter the total cash returned. Enter 0 if none has been returned yet.");
    return;
  }
  if (rawCoinsAmount === "") {
    toast("Specify the coin amount returned. Enter 0 if no coins were returned.");
    return;
  }
  const amount = Number(rawRefundAmount);
  const coins = Number(rawCoinsAmount);
  const excess = pendingRefundTransaction.existingId ? pendingRefundTransaction.pendingAmount
    : Number(pendingRefundTransaction.cashReceived || 0) + Number(pendingRefundTransaction.qrReceived || 0) - Number(pendingRefundTransaction.required || 0);
  if (!Number.isFinite(amount) || amount < 0 || amount > excess) {
    toast(`Enter a cash refund from ₹0 up to ₹${excess.toFixed(2)}.`);
    return;
  }
  if (!Number.isFinite(coins) || coins < 0 || coins > amount) {
    toast("Coins returned must be between ₹0 and the total refund amount.");
    return;
  }
  try {
    if (pendingRefundTransaction.existingId) {
      if (amount === 0) {
        if (coins !== 0) {
          toast("Coins cannot be returned when the total refund is ₹0.");
          return;
        }
        pendingRefundTransaction = null;
        $("refundDialog").close();
        toast("No additional refund recorded; the pending amount remains due.");
        return;
      }
      recordRefund(activeDate, pendingRefundTransaction.existingId, amount, coins);
      pendingRefundTransaction = null;
      $("refundDialog").close();
      await persist();
      renderAll(activeDate);
      renderPendingRefunds();
      toast("Cash refund recorded, with notes and coins tracked separately.");
    } else {
      await saveCustomerTransaction(pendingRefundTransaction, amount, coins);
    }
  } catch (err) { toast(err.message || "Could not save refund."); }
});

function openPendingRefundPrompt(transactionId) {
  const day = currentDay(activeDate);
  if (!day || day.Status === "Closed") {
    toast("Start an open business day before paying a pending customer refund.");
    return;
  }
  const transaction = rows("Transactions").find(row => row["Transaction ID"] === transactionId);
  if (!transaction) return;
  const pending = Number(transaction["Refund Pending"] || 0);
  if (!(pending > 0)) return;
  pendingRefundTransaction = { existingId: transaction["Transaction ID"], pendingAmount: pending };
  $("refundPromptText").textContent = `₹${pending.toFixed(2)} is still pending for this customer. Enter how much is returned now, and separately enter the coin portion (0 if none).`;
  $("partialRefundField").classList.add("hidden");
  $("refundChoiceActions").classList.remove("hidden");
  $("partialRefundAmount").value = "";
  $("partialRefundAmount").max = String(pending);
  $("refundCoinsAmount").value = "";
  $("refundDialog").showModal();
}
["transactionsBody", "pendingRefundsBody"].forEach(id => $(id).addEventListener("click", event => {
  const button = event.target.closest("[data-refund-id]");
  if (button) openPendingRefundPrompt(button.dataset.refundId);
}));

function openCustomerReceiptDialog(transactionId = "") {
  const day = currentDay(activeDate);
  if (!day || day.Status === "Closed") {
    toast("Start an open business day before receiving a pending customer payment.");
    return;
  }
  const pending = rows("Transactions").filter(row => row.Type === "receive" && Number(row["Balance Due"] || 0) > 0)
    .sort((a, b) => String(a.Date).localeCompare(String(b.Date)));
  const select = $("customerReceiptTransaction");
  select.innerHTML = pending.map(row => `<option value="${String(row["Transaction ID"]).replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]))}">${String(row["Policy / Name"] || "Unnamed customer").replace(/[&<>"']/g, " ")} — ${row.Date} — due ₹${Number(row["Balance Due"]).toFixed(2)} — ${row["Transaction ID"]}</option>`).join("");
  if (!pending.length) { toast("There are no outstanding customer payments."); return; }
  if (transactionId && pending.some(row => row["Transaction ID"] === transactionId)) select.value = transactionId;
  $("customerReceiptCash").value = "";
  $("customerReceiptQr").value = "";
  $("customerReceiptPlatform").value = "P1";
  updateCustomerReceiptLimit();
  $("customerReceiptDialog").showModal();
}

function updateCustomerReceiptLimit() {
  const selected = rows("Transactions").find(row => row["Transaction ID"] === $("customerReceiptTransaction").value);
  const due = Number(selected?.["Balance Due"] || 0);
  const cash = Number($("customerReceiptCash").value || 0);
  const qr = Number($("customerReceiptQr").value || 0);
  const remaining = due - cash - qr;
  $("customerReceiptDue").textContent = `Outstanding: ₹${due.toFixed(2)} · After this receipt: ₹${Math.max(0, remaining).toFixed(2)}`;
  $("customerReceiptPlatformField").classList.toggle("hidden", qr <= 0);
}

$("customerReceivablesBody").addEventListener("click", event => {
  const button = event.target.closest("[data-customer-receipt-id]");
  if (button) openCustomerReceiptDialog(button.dataset.customerReceiptId);
});
$("customerReceiptBtn").addEventListener("click", () => openCustomerReceiptDialog());
["customerReceiptTransaction", "customerReceiptCash", "customerReceiptQr", "customerReceiptPlatform"].forEach(id => $(id).addEventListener("input", updateCustomerReceiptLimit));
$("customerReceiptForm").addEventListener("submit", async event => {
  event.preventDefault();
  try {
    const result = recordCustomerReceipt(activeDate, $("customerReceiptTransaction").value, {
      cashAmount: $("customerReceiptCash").value, qrAmount: $("customerReceiptQr").value, platform: $("customerReceiptPlatform").value
    });
    await persist();
    $("customerReceiptDialog").close();
    renderAll(activeDate);
    toast(`₹${result.amount.toFixed(2)} received. Remaining customer due: ₹${result.remaining.toFixed(2)}.`);
  } catch (error) { toast(error.message || "Could not record customer payment."); }
});
$("cancelCustomerReceipt").addEventListener("click", () => $("customerReceiptDialog").close());
$("closeCustomerReceipt").addEventListener("click", () => $("customerReceiptDialog").close());
$("customerReceiptQr").addEventListener("input", () => {
  const hasQr = Number($("customerReceiptQr").value || 0) > 0;
  $("customerReceiptPlatformField").classList.toggle("hidden", !hasQr);
});



function resetOfficeEntryForm() {
  const form = $("officeEntryForm");
  form.reset();
  form.elements.entryId.value = "";
  $("officeEntryFormTitle").textContent = "Add a Receipts & Payments entry";
  $("officeEntrySubmit").textContent = "Add entry";
  $("cancelOfficeEdit").classList.add("hidden");
}

$("officeEntryForm").addEventListener("submit", async e => {
  e.preventDefault();
  try {
    const form = e.currentTarget;
    const values = Object.fromEntries(new FormData(form).entries());
    const id = values.entryId;
    if (id) {
      updateOfficeEntry(id, values);
      toast("Office entry updated.");
    } else {
      addOfficeEntry(activeDate, values);
      toast("Office entry added.");
    }
    await persist();
    resetOfficeEntryForm();
    renderOfficeSummary(activeDate);
  } catch (err) {
    toast(err.message || "Could not save office entry.");
  }
});

$("cancelOfficeEdit").addEventListener("click", resetOfficeEntryForm);

$("officeEntriesBody").addEventListener("click", async e => {
  const editButton = e.target.closest("[data-office-edit]");
  const deleteButton = e.target.closest("[data-office-delete]");

  if (editButton) {
    const id = editButton.dataset.officeEdit;
    const entry = getOfficeEntries(activeDate).find(item => item["Entry ID"] === id);
    if (!entry) {
      toast("The office entry could not be found.");
      return;
    }
    const form = $("officeEntryForm");
    form.elements.entryId.value = id;
    form.elements.section.value = entry.Section;
    form.elements.category.value = entry.Category;
    form.elements.amount.value = entry.Amount;
    form.elements.details.value = entry.Details || "";
    $("officeEntryFormTitle").textContent = "Edit Receipts & Payments entry";
    $("officeEntrySubmit").textContent = "Save changes";
    $("cancelOfficeEdit").classList.remove("hidden");
    form.scrollIntoView({ behavior: "smooth", block: "center" });
    return;
  }

  if (deleteButton) {
    const id = deleteButton.dataset.officeDelete;
    if (!window.confirm("Delete this Receipts & Payments entry?")) return;
    try {
      deleteOfficeEntry(id);
      await persist();
      resetOfficeEntryForm();
      renderOfficeSummary(activeDate);
      toast("Office entry deleted.");
    } catch (err) {
      toast(err.message || "Could not delete office entry.");
    }
  }
});

$("openingForm").addEventListener("submit", async e => {
  e.preventDefault();

  try {
    const f = Object.fromEntries(new FormData(e.currentTarget).entries());
    const blockingOpenDay = getOpenBusinessDays()[0] || null;
    if (blockingOpenDay && String(blockingOpenDay.Date).slice(0, 10) !== f.businessDate) {
      throw new Error(`Business day ${String(blockingOpenDay.Date).slice(0, 10)} is still open. Complete its opening setup and close it before selecting another date.`);
    }

    activeDate = f.businessDate;

    const existingDay = rows("Daily Summary").find(row => String(row.Date) === activeDate);
    const platformBalances = { P1: f.p1Balance, P2: f.p2Balance, P3: f.p3Balance };
    if (existingDay && existingDay.Status !== "Closed") {
      setPlatformOpeningBalances(activeDate, platformBalances);
    } else {
      startDay(
        activeDate,
        f.openingCash,
        f.openingCoins,
        { P1: f.p1Limit, P2: f.p2Limit, P3: f.p3Limit },
        platformBalances
      );
    }

    await persist();

    setDaySetupVisible(false);
    renderAll(activeDate);
    toast(existingDay ? "Today's platform opening balances saved." : "Business day started.");

  } catch (err) {
    toast(err.message);
  }
});

$("closingForm").addEventListener("input", updateClosingDifference);

$("closingForm").addEventListener("submit", async e => {
  e.preventDefault();

  try {
    const f = Object.fromEntries(new FormData(e.currentTarget).entries());

    const closedDate = activeDate;
    const result = closeDay(
      closedDate,
      f.closingCash,
      f.closingCoins,
      f.remarks
    );

    await persist();

    const nextOpenDay = getOpenBusinessDays()[0] || null;
    const today = localDate();
    if (nextOpenDay) {
      activeDate = String(nextOpenDay.Date).slice(0, 10);
      setMissedDayBanner(activeDate, activeDate < today);
      setDaySetupVisible(false);
      showClosingForm(currentDay(activeDate));
    } else if (closedDate < today) {
      activeDate = today;
      setMissedDayBanner(activeDate, false);
      const todayRecord = currentDay(today);
      if (!todayRecord) {
        const suggestion = getSuggestedOpening(activeDate);
        fillOpeningSuggestions(activeDate, suggestion, currentLimitsForDate(activeDate));
        setDaySetupVisible(true);
        showClosingForm(null);
      } else {
        setDaySetupVisible(false);
        showClosingForm(todayRecord);
      }
    } else {
      activeDate = closedDate;
      setMissedDayBanner(activeDate, false);
      setDaySetupVisible(false);
      showClosingForm(currentDay(activeDate));
    }

    renderAll(activeDate);
    toast(`Day ${closedDate} closed. Cash difference: ₹${result.cashDifference.toFixed(2)}; coin difference: ₹${result.coinsDifference.toFixed(2)}.${activeDate !== closedDate ? ` Next step: ${activeDate === today ? "start today's business day" : `resume ${activeDate}`}.` : ""}`);

  } catch (err) {
    toast(err.message);
  }
});

["transactionSearch", "platformFilter", "statusFilter"].forEach(id => {
  $(id).addEventListener("input", () => renderTransactions(activeDate));
});

$("prevPage").addEventListener("click", () => changeTransactionPage(-1, activeDate));
$("nextPage").addEventListener("click", () => changeTransactionPage(1, activeDate));

const recovery = loadSession();

if (recovery) {
  console.info("Previous session detected:", recovery.savedAt);
}
