import {rows} from "./excel-engine.js";
const n = v => Number(v)||0;
export function currentDay(date){
  return rows("Daily Summary").find(r=>String(r.Date)===date) || null;
}
export function dayTransactions(date){
  return rows("Transactions").filter(r=>String(r.Date)===date);
}
export function calculateDay(date){
  const day=currentDay(date);
  if(!day) return {opening:0,coins:0,cashIn:0,cashOut:0,balance:0,pending:0,transactions:[]};
  const tx=dayTransactions(date);
  const cashIn=tx.reduce((s,t)=>s+(t["Payment Mode"]==="Cash"?Math.max(0,n(t["Net Cash Flow"])):0),0);
  const cashOut=tx.reduce((s,t)=>s+(t["Payment Mode"]==="Cash"?Math.max(0,-n(t["Net Cash Flow"])):0),0);
  const opening=n(day["Opening Cash"])+n(day["Opening Coins"]);
  return {opening,coins:n(day["Opening Coins"]),cashIn,cashOut,balance:opening+cashIn-cashOut,pending:tx.reduce((s,t)=>s+n(t["Balance Due"]),0),transactions:tx};
}
export const money = v => new Intl.NumberFormat("en-IN",{style:"currency",currency:"INR",maximumFractionDigits:2}).format(Number(v)||0);
export const num = v => Number(v)||0;
