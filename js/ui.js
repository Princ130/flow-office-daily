import {calculateDay,money,num} from "./cash-manager.js";
import {platformUsage} from "./platform-manager.js";
import {getRecentTransactions} from "./reports.js";
import {rows} from "./excel-engine.js";
let currentPage=1, pageSize=50;
const $=id=>document.getElementById(id);
export function toast(message){const el=$("toast");el.textContent=message;el.classList.add("show");setTimeout(()=>el.classList.remove("show"),2800);}
export function showApp(){ $("setupScreen").classList.add("hidden"); $("appShell").classList.remove("hidden"); }
export function showSetup(){ $("appShell").classList.add("hidden"); $("setupScreen").classList.remove("hidden"); }
export function setWorkbookStatus(name){$("workbookStatus").textContent=name;}
export function setDaySetupVisible(visible){$("daySetupBanner").classList.toggle("hidden",!visible);}
export function fillOpeningSuggestions(date,suggestion,limits){
  const f=$("openingForm");
  f.elements.businessDate.value=date;
  f.elements.openingCash.value=suggestion.cash;
  f.elements.openingCoins.value=suggestion.coins;
  f.elements.p1Limit.value=limits.P1||"";
  f.elements.p2Limit.value=limits.P2||"";
  f.elements.p3Limit.value=limits.P3||"";
  $("carryForwardNote").textContent=suggestion.source?`Suggested from ${suggestion.source}'s closing count. Physically count and confirm or edit these values.`:"First working day: enter the opening cash and coin count manually.";
}
export function setActiveView(name){
  document.querySelectorAll(".view").forEach(v=>v.classList.add("hidden"));
  $(`${name}View`).classList.remove("hidden");
  document.querySelectorAll(".nav-link").forEach(b=>b.classList.toggle("active",b.dataset.view===name));
  $("pageTitle").textContent=({dashboard:"Dashboard",transactions:"Transactions",limits:"Platform limits",closing:"Day closing"})[name]||"Dashboard";
}
function badge(status){
  const cls=status==="Settled Transaction"?"status-settled":status==="Pending Due"?"status-pending":"status-expense";
  return `<span class="status-tag ${cls}">${status}</span>`;
}
function platformCard(p){
  const percent=p.limit?Math.min(100,(p.used/p.limit)*100):0;
  return `<article class="platform-card"><h4>${p.platform==="P1"?"Platform 1":p.platform==="P2"?"Platform 2":"Platform 3"}</h4><div class="platform-amounts"><span>Used ${money(p.used)}</span><span>Limit ${money(p.limit)}</span></div><div class="progress"><span class="${percent>85?"warning":""}" style="width:${percent}%"></span></div><div class="platform-bottom"><span>${percent.toFixed(1)}% used</span><b>${money(p.remaining)} left</b></div></article>`;
}
export function renderAll(date){
  $("todayLabel").textContent=new Date(`${date}T12:00:00`).toLocaleDateString("en-IN",{weekday:"short",day:"numeric",month:"short",year:"numeric"});
  const calc=calculateDay(date);
  $("cashBalance").textContent=money(calc.balance); $("cashIn").textContent=money(calc.cashIn);
  $("cashOut").textContent=money(calc.cashOut); $("pendingDues").textContent=money(calc.pending);
  const usage=platformUsage(date);
  $("platformCards").innerHTML=usage.map(platformCard).join("");
  $("limitsDetail").innerHTML=usage.map(platformCard).join("");
  const recent=getRecentTransactions(date);
  $("recentTransactions").innerHTML=recent.length?recent.map(t=>`<tr><td>${t["Daily Serial"]}</td><td>${escapeHtml(t["Policy / Name"]||"—")}</td><td>${t.Platform}</td><td>${money(t["Net Cash Flow"])}</td><td>${badge(t.Status)}</td></tr>`).join(""):`<tr><td colspan="5" class="empty-state">No transactions recorded today.</td></tr>`;
  renderTransactions(date);
  $("expectedClosing").textContent=money(calc.balance);
  updateClosingDifference();
}
function escapeHtml(v){return String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));}
export function changeTransactionPage(delta,date){currentPage=Math.max(1,currentPage+delta);renderTransactions(date);}
export function renderTransactions(date){
  const search=$("transactionSearch").value.toLowerCase(), pf=$("platformFilter").value, sf=$("statusFilter").value;
  let data=rows("Transactions").filter(r=>String(r.Date)===date);
  data=data.filter(t=>(!search||[t["Policy / Name"],t.Details,t["Transaction ID"]].join(" ").toLowerCase().includes(search))&&(!pf||t.Platform===pf)&&(!sf||t.Status===sf));
  const total=data.length,pages=Math.max(1,Math.ceil(total/pageSize));currentPage=Math.min(currentPage,pages);
  const slice=data.slice((currentPage-1)*pageSize,currentPage*pageSize);
  $("transactionsBody").innerHTML=slice.length?slice.map(t=>`<tr><td>${t["Daily Serial"]}</td><td>${escapeHtml(t["Policy / Name"]||"—")}</td><td>${escapeHtml(t.Details)}</td><td>${t.Type}</td><td>${t.Platform}</td><td>${money(t["Required Amount"])}</td><td>${money(t["Received Amount"])}</td><td>${money(t["Refunded Amount"])}</td><td>${money(t["Net Cash Flow"])}</td><td>${money(t["Balance Due"])}</td><td>${badge(t.Status)}</td></tr>`).join(""):`<tr><td colspan="11" class="empty-state">No matching transactions.</td></tr>`;
  $("recordCount").textContent=`${total} records`; $("pageInfo").textContent=`Page ${currentPage} of ${pages}`;
  $("prevPage").disabled=currentPage<=1; $("nextPage").disabled=currentPage>=pages;
}
export function updateClosingDifference(){
  const form=$("closingForm"); if(!form) return;
  const a=num(form.elements.closingCash.value)+num(form.elements.closingCoins.value);
  const expected=num($("expectedClosing").textContent.replace(/[₹,]/g,""));
  $("closingDifference").textContent=money(a-expected);
}
export function showClosingForm(day){
  const f=$("closingForm");
  f.elements.closingCash.value=day?.["Closing Cash"]??"";
  f.elements.closingCoins.value=day?.["Closing Coins"]??"";
}
