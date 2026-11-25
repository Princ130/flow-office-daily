import {rows,replaceRows} from "./excel-engine.js";
import {num} from "./cash-manager.js";

const PLATFORMS = ["P1", "P2", "P3"];

export function platformUsage(date){
  const limits=rows("Platform Limits").filter(r=>String(r.Date)===date);
  const tx=rows("Transactions").filter(r=>String(r.Date)===date);
  const head=rows("Head Ledger").filter(r=>String(r.Date)===date);
  const customerReceipts=rows("Customer Receipts").filter(r=>String(r.Date)===date);
  return PLATFORMS.map(p=>{
    const item=limits.find(r=>r.Platform===p)||{};
    const used=tx.filter(t=>t.Platform===p).reduce((s,t)=>s+num(t["Limit Usage"]),0);
    const limit=num(item["Daily Limit"]);
    const openingBalance=num(item["Opening Balance"]);
    const customerVirtualIn=tx.filter(t=>t.Type==="receive" && t.Platform===p)
      .reduce((s,t)=>s+num(t["CHQ/QR Received"]),0)
      + customerReceipts.filter(r=>r.Platform===p).reduce((s,r)=>s+num(r["CHQ/QR Amount"]),0);
    const headVirtualIn=head.filter(t=>t.Mode==="Virtual" && t.Platform===p && t.Action==="Receive")
      .reduce((s,t)=>s+num(t.Amount),0);
    const headVirtualOut=head.filter(t=>t.Mode==="Virtual" && t.Platform===p && t.Action==="Send")
      .reduce((s,t)=>s+num(t.Amount),0);
    const balance=openingBalance+customerVirtualIn+headVirtualIn-headVirtualOut;
    return {platform:p,limit,used,remaining:limit-used,openingBalance,
      customerVirtualIn,headVirtualIn,headVirtualOut,balance};
  });
}

export function setPlatformLimits(date,values,balances={}){
  const all=rows("Platform Limits").filter(r=>String(r.Date)!==date);
  for(const p of PLATFORMS){
    const limit=num(values[p]);
    all.push({"Date":date,"Platform":p,"Daily Limit":limit,"Used Limit":0,
      "Remaining Limit":limit,"Opening Balance":num(balances[p])});
  }
  replaceRows("Platform Limits",all);
}

export function setPlatformOpeningBalances(date,balances={}){
  const all=rows("Platform Limits");
  for(const p of PLATFORMS){
    const index=all.findIndex(r=>String(r.Date)===date && r.Platform===p);
    if(index<0) throw new Error(`Set the daily limit for ${p} before entering its opening balance.`);
    all[index]={...all[index],"Opening Balance":num(balances[p])};
  }
  replaceRows("Platform Limits",all);
}

export function validateLimit(date,platform,usage){
  if(!PLATFORMS.includes(platform)) return {ok:true};
  const p=platformUsage(date).find(x=>x.platform===platform);
  if(!p || p.limit<=0) return {ok:false,message:`Set a daily limit for ${platform} before recording this transaction.`};
  if(p.used+usage>p.limit) return {ok:false,message:`${platform} limit exceeded. Remaining: ₹${Math.max(0,p.remaining).toFixed(2)}.`};
  return {ok:true};
}

export function validatePlatformBalance(date,platform,amount){
  if(!PLATFORMS.includes(platform)) return {ok:false,message:"Choose a valid platform."};
  const p=platformUsage(date).find(x=>x.platform===platform);
  if(!p) return {ok:false,message:`Enter today's opening balance for ${platform} first.`};
  if(amount>p.balance) return {ok:false,message:`Insufficient ${platform} balance. Available: ₹${Math.max(0,p.balance).toFixed(2)}.`};
  return {ok:true};
}
