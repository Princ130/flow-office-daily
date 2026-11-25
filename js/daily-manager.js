import { rows, appendRow, replaceRows, logAction } from "./excel-engine.js";
import { setPlatformLimits } from "./platform-manager.js";
import { calculateDay, num } from "./cash-manager.js";

export function getOpenBusinessDays() {
  return rows("Daily Summary")
    .filter(row => String(row.Status || "Open") !== "Closed")
    .sort((a, b) => String(a.Date).slice(0, 10).localeCompare(String(b.Date).slice(0, 10)));
}

export function getSuggestedOpening(date) {
  const previous = rows("Daily Summary")
    .filter(row => String(row.Date).slice(0, 10) < date && row.Status === "Closed")
    .sort((a, b) => String(b.Date).slice(0, 10).localeCompare(String(a.Date).slice(0, 10)))[0];
  if (!previous) return { cash: "", coins: "", source: null };
  return {
    cash: num(previous["Closing Cash"]),
    coins: num(previous["Closing Coins"]),
    source: String(previous.Date).slice(0, 10)
  };
}

export function startDay(date, cash, coins, limits, platformBalances = {}) {
  const otherOpenDay = getOpenBusinessDays().find(row => String(row.Date).slice(0, 10) !== date);
  if (otherOpenDay) {
    throw new Error(`Business day ${String(otherOpenDay.Date).slice(0, 10)} is still open. Resume and close it before starting ${date}.`);
  }
  if (rows("Daily Summary").some(row => String(row.Date).slice(0, 10) === date)) {
    throw new Error("A business-day record already exists for this date.");
  }
  appendRow("Daily Summary", {
    Date: date, "Opening Cash": num(cash), "Opening Coins": num(coins),
    "Closing Cash": "", "Closing Coins": "", "Expected Closing": "",
    "Cash Difference": "", "Expected Cash": "", "Cash Difference Separate": "",
    "Expected Coins": "", "Coins Difference": "", "Total Cash In": 0,
    "Total Cash Out": 0, Status: "Open", Remarks: ""
  });
  appendRow("Denominations", {
    Date: date, "Opening Notes": num(cash), "Opening Coins": num(coins),
    "Closing Notes": "", "Closing Coins": ""
  });
  setPlatformLimits(date, limits, platformBalances);
  logAction("DAY_OPENED", date, `Opening cash ${cash}; coins ${coins}`);
}

export function closeDay(date, closingCash, closingCoins, remarks) {
  const day = rows("Daily Summary").find(row => String(row.Date) === date);
  if (!day) throw new Error("No open business day found.");
  if (day.Status === "Closed") throw new Error("This business day is already closed.");

  const calc = calculateDay(date);
  const actualCash = num(closingCash);
  const actualCoins = num(closingCoins);
  const cashDifference = actualCash - calc.expectedCash;
  const coinsDifference = actualCoins - calc.expectedCoins;
  const expected = calc.expectedCash + calc.expectedCoins;
  const actual = actualCash + actualCoins;
  const difference = cashDifference + coinsDifference;

  replaceRows("Daily Summary", rows("Daily Summary").map(row => String(row.Date) === date ? {
    ...row,
    "Closing Cash": actualCash,
    "Closing Coins": actualCoins,
    "Expected Closing": expected,
    "Cash Difference": difference,
    "Expected Cash": calc.expectedCash,
    "Cash Difference Separate": cashDifference,
    "Expected Coins": calc.expectedCoins,
    "Coins Difference": coinsDifference,
    "Total Cash In": calc.cashIn,
    "Total Cash Out": calc.cashOut,
    Status: "Closed",
    Remarks: remarks || ""
  } : row));

  replaceRows("Denominations", rows("Denominations").map(row => String(row.Date) === date ? {
    ...row, "Closing Notes": actualCash, "Closing Coins": actualCoins
  } : row));

  logAction("DAY_CLOSED", date,
    `Expected cash ${calc.expectedCash}; actual cash ${actualCash}; cash difference ${cashDifference}; expected coins ${calc.expectedCoins}; actual coins ${actualCoins}; coins difference ${coinsDifference}`);

  return { expected, actual, difference, expectedCash: calc.expectedCash, actualCash, cashDifference,
    expectedCoins: calc.expectedCoins, actualCoins, coinsDifference };
}
