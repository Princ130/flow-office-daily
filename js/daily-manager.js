import {rows,appendRow,replaceRows,logAction} from "./excel-engine.js";
import {setPlatformLimits} from "./platform-manager.js";
import {calculateDay,num} from "./cash-manager.js";
export function getSuggestedOpening(date) {
  const previous = rows("Daily Summary")
    .filter(r =>
      String(r.Date).slice(0, 10) < date &&
      r.Status === "Closed"
    )
    .sort((a, b) =>
      String(b.Date).slice(0, 10).localeCompare(String(a.Date).slice(0, 10))
    )[0];

  if (!previous) {
    return { cash: "", coins: "", source: null };
  }

  return {
    cash: num(previous["Closing Cash"]),
    coins: num(previous["Closing Coins"]),
    source: String(previous.Date).slice(0, 10)
  };
}
export function startDay(date,cash,coins,limits){
  if(rows("Daily Summary").some(r=>String(r.Date)===date)) throw new Error("A business-day record already exists for this date.");
  appendRow("Daily Summary",{"Date":date,"Opening Cash":num(cash),"Opening Coins":num(coins),"Closing Cash":"","Closing Coins":"","Expected Closing":"","Cash Difference":"","Total Cash In":0,"Total Cash Out":0,"Status":"Open","Remarks":""});
  appendRow("Denominations",{"Date":date,"Opening Notes":num(cash),"Opening Coins":num(coins),"Closing Notes":"","Closing Coins":""});
  setPlatformLimits(date,limits); logAction("DAY_OPENED",date,`Opening cash ${cash}; coins ${coins}`);
}
export function closeDay(date,closingCash,closingCoins,remarks){
  const day=rows("Daily Summary").find(r=>String(r.Date)===date);
  if(!day) throw new Error("No open business day found.");
  if(day.Status==="Closed") throw new Error("This business day is already closed.");
  const calc=calculateDay(date),actual=num(closingCash)+num(closingCoins),difference=actual-calc.balance;
  const all=rows("Daily Summary").map(r=>String(r.Date)===date?{...r,"Closing Cash":num(closingCash),"Closing Coins":num(closingCoins),"Expected Closing":calc.balance,"Cash Difference":difference,"Total Cash In":calc.cashIn,"Total Cash Out":calc.cashOut,"Status":"Closed","Remarks":remarks||""}:r);
  replaceRows("Daily Summary",all);
  const den=rows("Denominations").map(r=>String(r.Date)===date?{...r,"Closing Notes":num(closingCash),"Closing Coins":num(closingCoins)}:r);
  replaceRows("Denominations",den); logAction("DAY_CLOSED",date,`Expected ${calc.balance}; actual ${actual}; difference ${difference}`);
  return {expected:calc.balance,actual,difference};
}
