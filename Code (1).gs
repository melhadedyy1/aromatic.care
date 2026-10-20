// AROMATIC CARE - الخادم (Google Apps Script) مربوط بملف Google Sheet
const CFG = {
  BRANCH_PASS: '0000',          // كلمة سر الفروع: بلاغ جديد فقط
  ADMIN_PASS: 'A987654321',     // كلمة سر الإدارة: كل الشاشات
  SHEET: 'Complaints',
  FOLDER: 'AromaticCare_Images'
};
const F = ['no','createdAt','bc','bn','name','phone','type','product','buy','recv','comps','desc','img1','img2','finding','dec','reason','done','cons','decDate','status'];

function doPost(e) {
  const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  try {
    const r = JSON.parse(e.postData.contents);
    const role = r.pass === CFG.ADMIN_PASS ? 'admin' : r.pass === CFG.BRANCH_PASS ? 'branch' : null;
    if (!role) return out({ ok: false, err: 'auth' });
    if (r.action === 'login') return out({ ok: true, role: role });
    const lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      const sh = sheet();
      if (r.action === 'submit') {
        const o = r.data || {};
        if (role !== 'admin') ['finding','reason','done','cons','decDate'].forEach(k => o[k] = ''), o.dec = 'pending';
        o.no = nextNo(sh); o.createdAt = new Date().toISOString();
        o.img1 = img(o.img1, o.no + '-1'); o.img2 = img(o.img2, o.no + '-2');
        o.status = status(o);
        sh.appendRow(F.map(k => o[k] == null ? '' : String(o[k])));
        return out({ ok: true, no: o.no, img1: o.img1, img2: o.img2 });
      }
      if (role !== 'admin') return out({ ok: false, err: 'forbidden' });
      if (r.action === 'list') {
        const v = sh.getDataRange().getValues().slice(1);
        return out({ ok: true, items: v.map(row => { const o = {}; F.forEach((k, i) => o[k] = row[i]); return o; }).reverse() });
      }
      const idx = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().findIndex(x => x[0] === (r.no || (r.data && r.data.no)));
      if (idx < 1) return out({ ok: false, err: 'notfound' });
      if (r.action === 'delete') { sh.deleteRow(idx + 1); return out({ ok: true }); }
      if (r.action === 'update') {
        const old = sh.getRange(idx + 1, 1, 1, F.length).getValues()[0], o = {};
        F.forEach((k, i) => o[k] = old[i]);
        Object.keys(r.data).forEach(k => { if (k !== 'no' && k !== 'createdAt') o[k] = r.data[k]; });
        o.img1 = img(o.img1, o.no + '-1'); o.img2 = img(o.img2, o.no + '-2');
        o.status = status(o);
        sh.getRange(idx + 1, 1, 1, F.length).setValues([F.map(k => o[k] == null ? '' : String(o[k]))]);
        return out({ ok: true, img1: o.img1, img2: o.img2 });
      }
      return out({ ok: false, err: 'action' });
    } finally { lock.releaseLock(); }
  } catch (err) { return out({ ok: false, err: String(err) }); }
}
function sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CFG.SHEET);
  if (!sh) { sh = ss.insertSheet(CFG.SHEET); sh.getRange('A:U').setNumberFormat('@'); sh.appendRow(F); sh.setFrozenRows(1); sh.setRightToLeft(true); }
  return sh;
}
function nextNo(sh) {
  const y = new Date().getFullYear(), p = 'AS-' + y + '-';
  let m = 0;
  sh.getRange(1, 1, sh.getLastRow(), 1).getValues().forEach(x => { if (String(x[0]).indexOf(p) === 0) m = Math.max(m, parseInt(String(x[0]).slice(p.length), 10) || 0); });
  return p + String(m + 1).padStart(4, '0');
}
function status(o) { return o.dec === 'pending' || !o.dec ? 'قيد الفحص' : (o.done === 'تم الاستبدال' || o.dec === 'no') ? 'تم إغلاق الشكوى' : 'تم اتخاذ القرار'; }
function img(v, name) {
  if (!v || String(v).indexOf('data:') !== 0) return v || '';
  const it = DriveApp.getFoldersByName(CFG.FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(CFG.FOLDER);
  const f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(String(v).split(',')[1]), 'image/jpeg', name + '.jpg'));
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return f.getId();
}
