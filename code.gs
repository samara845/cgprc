// CGPRC backend: Google Sheet + Apps Script
// Deploy: Extensions > Apps Script > paste > Deploy > New deployment > Web app
// Execute as: Me | Who has access: Anyone. Paste the /exec URL in index.html -> CFG.API
const PASS = 'cgprc@123'; // change this, and also CFG.ADMIN_PASS in index.html
const BH = ['id','name','mobile','age','service','date','slot','concern','utr','st','ts'];

function T(n, h) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(n);
  if (!s) { s = ss.insertSheet(n); s.appendRow(h); s.getRange('A:K').setNumberFormat('@'); }
  return s;
}
function rows(s) {
  const v = s.getDataRange().getValues(), h = v.shift();
  return v.map(r => { const o = {}; h.forEach((k, i) => o[k] = String(r[i])); return o; });
}
function adm(p) { if (p.pw !== PASS) throw 'Wrong password'; }

function doPost(e) {
  let out;
  try { const q = JSON.parse(e.postData.contents); out = run(q.a, q.p || {}); }
  catch (x) { out = { err: String(x) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function run(a, p) {
  const B = T('Bookings', BH), N = T('Notices', ['id', 'text']), S = T('Settings', ['k', 'v']);
  const br = rows(B);
  if (a === 'pub') { const s = {}; rows(S).forEach(r => s[r.k] = r.v); return { n: rows(N).reverse(), s: s }; }
  if (a === 'slots') return br.filter(x => x.date === p.date && x.st !== 'Rejected').map(x => x.slot);
  if (a === 'add') {
    if (br.some(x => x.utr === p.utr)) throw 'This transaction number is already used';
    if (br.some(x => x.date === p.date && x.slot === p.slot && x.st !== 'Rejected')) throw 'This slot was just taken, choose another';
    const id = 'CG' + Date.now().toString(36).toUpperCase().slice(-6);
    B.appendRow(BH.map(k => k === 'id' ? id : k === 'st' ? 'Pending Verification' : k === 'ts' ? new Date().toISOString() : (p[k] || '')));
    return { id: id };
  }
  if (a === 'track') {
    const q = String(p.q).trim().toUpperCase();
    return br.filter(x => x.id === q || x.mobile === q).map(x => { delete x.utr; return x; });
  }
  if (a === 'login') { adm(p); return { ok: 1 }; }
  if (a === 'list') { adm(p); return br; }
  if (a === 'status') {
    adm(p); const i = br.findIndex(x => x.id === p.id);
    if (i < 0) throw 'Not found';
    B.getRange(i + 2, 10).setValue(p.st); return { ok: 1 };
  }
  if (a === 'nadd') { adm(p); N.appendRow([String(Date.now()), p.text]); return { ok: 1 }; }
  if (a === 'ndel') {
    adm(p); const i = rows(N).findIndex(x => x.id === String(p.id));
    if (i >= 0) N.deleteRow(i + 2); return { ok: 1 };
  }
  if (a === 'sset') {
    adm(p); S.clear(); S.appendRow(['k', 'v']);
    Object.keys(p.s).forEach(k => S.appendRow([k, p.s[k]])); return { ok: 1 };
  }
  throw 'Unknown action';
}