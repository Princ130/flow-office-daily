import {rows,replaceRows} from "./excel-engine.js";
import {num} from "./cash-manager.js";
export function platformUsage(date){
  const limits=rows("Platform Limits").filter(r=>String(r.Date)===date);
  const tx=rows("Transactions").filter(r=>String(r.Date)===date);
  return ["P1","P2","P3"].map(p=>{
    const item=limits.find(r=>r.Platform===p)||{};
    const used=tx.filter(t=>t.Platform===p).reduce((s,t)=>s+num(t["Limit Usage"]),0);
    const limit=num(item["Daily Limit"]);
    return {platform:p,limit,used,remaining:limit-used};
  });
}
export function setPlatformLimits(date,values){
  const all=rows("Platform Limits").filter(r=>String(r.Date)!==date);
  for(const p of ["P1","P2","P3"]){
    const limit=num(values[p]);
    all.push({"Date":date,"Platform":p,"Daily Limit":limit,"Used Limit":0,"Remaining Limit":limit});
  }
  replaceRows("Platform Limits",all);
}
export function validateLimit(date,platform,usage){
  if(!["P1","P2","P3"].includes(platform)) return {ok:true};
  const p=platformUsage(date).find(x=>x.platform===platform);
  if(!p || p.limit<=0) return {ok:false,message:`Set a daily limit for ${platform} before recording this transaction.`};
  if(p.used+usage>p.limit) return {ok:false,message:`${platform} limit exceeded. Remaining: ₹${Math.max(0,p.remaining).toFixed(2)}.`};
  return {ok:true};
}
