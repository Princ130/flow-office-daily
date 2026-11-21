// Browser-only temporary state. The Excel workbook remains the persistent source of truth.
const KEY = "cashflow-office-session-v1";
export function saveSession(data){ try{ sessionStorage.setItem(KEY, JSON.stringify(data)); }catch(e){ console.warn("Session recovery unavailable",e); } }
export function loadSession(){ try{ return JSON.parse(sessionStorage.getItem(KEY) || "null"); }catch{return null;} }
export function clearSession(){ sessionStorage.removeItem(KEY); }
