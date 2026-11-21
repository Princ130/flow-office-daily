
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
  startDay,
  closeDay
} from "./daily-manager.js";

import { addTransaction } from "./transaction-manager.js";
import { currentDay } from "./cash-manager.js";

import {
  showApp,
  showSetup,
  setWorkbookStatus,
  setDaySetupVisible,
  fillOpeningSuggestions,
  setActiveView,
  renderAll,
  renderTransactions,
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
    P3: rec.find(r => r.Platform === "P3")?.["Daily Limit"] || ""
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

  activeDate = localDate();

  const day = currentDay(activeDate);

  if (!day) {
    const suggestion = getSuggestedOpening(activeDate);
    fillOpeningSuggestions(activeDate, suggestion, currentLimitsForDate(activeDate));
    setDaySetupVisible(true);
  } else {
    setDaySetupVisible(false);

    if (day.Status === "Closed") {
      toast("Today's business day is already closed. It is read-only in this prototype.");
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

$("newTransactionBtn").addEventListener("click", () => openTransaction("receive"));

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
  const refunded = form.elements.refunded;
  const mode = form.elements.mode;
  const headName = form.elements.headName;

  const platformField = $("platformField");
  const policyField = $("policyField");
  const receivedField = $("receivedField");
  const refundedField = $("refundedField");
  const modeField = $("paymentModeField");
  const headNameField = $("headNameField");

  if (platformField) platformField.classList.toggle("hidden", isSend);
  if (policyField) policyField.classList.toggle("hidden", isSend);
  if (receivedField) receivedField.classList.toggle("hidden", isSend);
  if (refundedField) refundedField.classList.toggle("hidden", isSend);
  if (modeField) modeField.classList.toggle("hidden", isSend);
  if (headNameField) headNameField.classList.toggle("hidden", !isSend);

  if (platform) platform.required = !isSend;
  if (policy) policy.required = !isSend;
  if (received) received.required = !isSend;
  if (refunded) refunded.required = !isSend;
  if (mode) mode.required = !isSend;
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

$("transactionForm").addEventListener("submit", async e => {
  e.preventDefault();

  try {
    const f = Object.fromEntries(new FormData(e.currentTarget).entries());

    const record = addTransaction(activeDate, f);

    await persist();

    $("transactionDialog").close();
    renderAll(activeDate);

    toast(`Transaction #${record["Daily Serial"]} recorded.`);

  } catch (err) {
    toast(err.message || "Could not save transaction.");
  }
});

$("openingForm").addEventListener("submit", async e => {
  e.preventDefault();

  try {
    const f = Object.fromEntries(new FormData(e.currentTarget).entries());

    activeDate = f.businessDate;

    startDay(
      activeDate,
      f.openingCash,
      f.openingCoins,
      { P1: f.p1Limit, P2: f.p2Limit, P3: f.p3Limit }
    );

    await persist();

    setDaySetupVisible(false);
    renderAll(activeDate);
    toast("Business day started.");

  } catch (err) {
    toast(err.message);
  }
});

$("closingForm").addEventListener("input", updateClosingDifference);

$("closingForm").addEventListener("submit", async e => {
  e.preventDefault();

  try {
    const f = Object.fromEntries(new FormData(e.currentTarget).entries());

    const result = closeDay(
      activeDate,
      f.closingCash,
      f.closingCoins,
      f.remarks
    );

    await persist();

    renderAll(activeDate);

    toast(`Day closed. Cash difference: ₹${result.difference.toFixed(2)}.`);

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
