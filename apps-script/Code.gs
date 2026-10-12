const PIN = 'CHANGE-ME-1234';

const TEN = {sheet: 'Tenants', r0: 5, r1: 34};
const PAY = {sheet: 'Payments', r0: 5, r1: 1004};
const APT = {sheet: 'Apartments', r0: 5, r1: 34};
const OWN = {sheet: 'Owners', r0: 5, r1: 14};
const CHG = {sheet: 'Rent Changes', r0: 5, r1: 504};
const SET = {sheet: 'Settings', deleteRow: 5, warnRow: 6, col: 2};

function doGet() { return out({ok: true, msg: 'Rentbook is running'}); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const q = JSON.parse(e.postData.contents);
    if (q.pin !== PIN) return out({ok: false, error: 'bad pin'});
    return out({ok: true, data: handle(q)});
  } catch (err) {
    return out({ok: false, error: String(err.message || err)});
  } finally { lock.releaseLock(); }
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sh(t) { return SpreadsheetApp.getActive().getSheetByName(t.sheet); }
function grid(t, ncol) { return sh(t).getRange(t.r0, 1, t.r1 - t.r0 + 1, ncol).getValues(); }
function dstr(v) { return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v || ''); }
function dval(s) { const p = String(s).split('-').map(Number); return new Date(p[0], p[1] - 1, p[2]); }
function emptyRow(t) {
  const col = sh(t).getRange(t.r0, 1, t.r1 - t.r0 + 1, 1).getValues();
  for (let i = 0; i < col.length; i++) if (col[i][0] === '') return t.r0 + i;
  throw new Error('The sheet is full. Ask the family IT person to add rows.');
}

function readSettings() {
  const s = sh(SET);
  if (!s) return {allowDelete: false, warnDays: 7};
  const del = String(s.getRange(SET.deleteRow, SET.col).getValue()).toLowerCase() === 'yes';
  const warn = Number(s.getRange(SET.warnRow, SET.col).getValue());
  return {allowDelete: del, warnDays: warn > 0 ? warn : 7};
}

function handle(q) {
  if (q.action === 'all') {
    const tenants = [], payments = [], changes = [];
    grid(TEN, 7).forEach((r, i) => { if (r[0] !== '') tenants.push({row: TEN.r0 + i, name: String(r[0]), apt: String(r[1]), phone: String(r[2]), rent: Number(r[3]) || 0, first: dstr(r[4]), status: r[5] || 'Active', notes: String(r[6])}); });
    grid(PAY, 7).forEach((r, i) => { if (r[0] !== '') payments.push({row: PAY.r0 + i, date: dstr(r[0]), tenant: String(r[1]), amount: Number(r[2]) || 0, method: String(r[3]), ref: String(r[4]), by: String(r[5]), notes: String(r[6])}); });
    if (sh(CHG)) grid(CHG, 4).forEach((r, i) => { if (r[0] !== '' && r[1] !== '') changes.push({row: CHG.r0 + i, date: dstr(r[0]), tenant: String(r[1]), rent: Number(r[2]) || 0, note: String(r[3])}); });
    const names = t => grid(t, 1).map(r => r[0]).filter(v => v !== '').map(String);
    return {tenants, payments, changes, settings: readSettings(), owners: names(OWN), apts: names(APT)};
  }
  if (q.action === 'addPayment') {
    if (!q.tenant || !(Number(q.amount) > 0) || !q.date) throw new Error('Missing tenant, amount or date');
    if (q.ref && grid(PAY, 7).some(r => String(r[4]) === String(q.ref))) throw new Error('This reference number is already saved (duplicate)');
    const row = emptyRow(PAY), s = sh(PAY);
    s.getRange(row, 5).setNumberFormat('@');
    s.getRange(row, 1, 1, 7).setValues([[dval(q.date), q.tenant, Number(q.amount), q.method || '', String(q.ref || ''), q.by || '', q.notes || '']]);
    return {row};
  }
  if (q.action === 'deletePayment') {
    if (!readSettings().allowDelete) throw new Error('Removing payments is switched off in Settings');
    if (q.row < PAY.r0 || q.row > PAY.r1) throw new Error('Bad row');
    sh(PAY).getRange(q.row, 1, 1, 7).clearContent();
    return {};
  }
  if (q.action === 'changeRent') {
    if (!q.tenant || !(Number(q.rent) > 0) || !q.date) throw new Error('Missing tenant, new rent or start date');
    if (!grid(TEN, 1).some(r => String(r[0]) === q.tenant)) throw new Error('Tenant not found');
    const rows = grid(CHG, 2), when = dstr(dval(q.date));
    const same = rows.findIndex(r => r[1] !== '' && String(r[1]) === q.tenant && dstr(r[0]) === when);
    const s = sh(CHG);
    if (same >= 0) {
      s.getRange(CHG.r0 + same, 3).setValue(Number(q.rent));
      return {row: CHG.r0 + same};
    }
    const row = emptyRow(CHG);
    s.getRange(row, 1, 1, 4).setValues([[dval(q.date), q.tenant, Number(q.rent), q.note || '']]);
    return {row};
  }
  if (q.action === 'setSetting') {
    const s = sh(SET);
    if (!s) throw new Error('The Settings tab is missing from the sheet');
    if (q.key === 'allowDelete') s.getRange(SET.deleteRow, SET.col).setValue(q.value ? 'Yes' : 'No');
    else if (q.key === 'warnDays') {
      const n = Math.round(Number(q.value));
      if (!(n >= 1 && n <= 60)) throw new Error('Choose 1 to 60 days');
      s.getRange(SET.warnRow, SET.col).setValue(n);
    } else throw new Error('Unknown setting');
    return {};
  }
  if (q.action === 'addCollector') {
    const name = String(q.name || '').trim();
    if (!name) throw new Error('Missing name');
    if (grid(OWN, 1).some(r => String(r[0]) === name)) throw new Error('That name is already on the list');
    const row = emptyRow(OWN);
    sh(OWN).getRange(row, 1).setValue(name);
    return {row};
  }
  if (q.action === 'removeCollector') {
    const i = grid(OWN, 1).findIndex(r => String(r[0]) === String(q.name));
    if (i < 0) throw new Error('Name not found');
    sh(OWN).getRange(OWN.r0 + i, 1, 1, 4).clearContent();
    return {};
  }
  if (q.action === 'saveTenant') {
    if (!q.name || !q.first) throw new Error('Missing name or first due date');
    const s = sh(TEN), all = grid(TEN, 1).map(r => String(r[0]));
    let row = Number(q.row) || 0, rent = Number(q.rent);
    if (all.some((n, i) => n === q.name && TEN.r0 + i !== row)) throw new Error('That name is already used');
    if (row) {
      const old = String(s.getRange(row, 1).getValue());
      rent = Number(s.getRange(row, 4).getValue()) || rent;
      if (old && old !== q.name) {
        const pr = sh(PAY).getRange(PAY.r0, 2, PAY.r1 - PAY.r0 + 1, 1);
        pr.setValues(pr.getValues().map(r => [r[0] === old ? q.name : r[0]]));
        if (sh(CHG)) {
          const cr = sh(CHG).getRange(CHG.r0, 2, CHG.r1 - CHG.r0 + 1, 1);
          cr.setValues(cr.getValues().map(r => [r[0] === old ? q.name : r[0]]));
        }
      }
    } else {
      if (!(rent > 0)) throw new Error('Missing rent');
      row = emptyRow(TEN);
    }
    s.getRange(row, 1, 1, 7).setValues([[q.name, q.apt || '', q.phone || '', rent, dval(q.first), q.status || 'Active', q.notes || '']]);
    if (q.apt && !grid(APT, 1).some(r => String(r[0]) === q.apt)) sh(APT).getRange(emptyRow(APT), 1).setValue(q.apt);
    return {row};
  }
  throw new Error('Unknown action');
}
