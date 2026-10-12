const PIN = 'CHANGE-ME-1234';

const TEN = {sheet: 'Tenants', r0: 5, r1: 34};
const PAY = {sheet: 'Payments', r0: 5, r1: 1004};
const TIMECOL = 11;   // Payments column K: Time Paid (added automatically)
const APT = {sheet: 'Apartments', r0: 5, r1: 34};
const OWN = {sheet: 'Owners', r0: 5, r1: 14};
const CHG = {sheet: 'Rent Changes', r0: 5, r1: 504};
const SET = {sheet: 'Settings', deleteRow: 5, warnRow: 6, hourRow: 7, stopRow: 8, col: 2};
const DEV = {sheet: 'Devices', r0: 5, r1: 54};   // columns: A Phone, B Reminders, C Added, D Last reminder, E Key

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

// "2:30 PM", "14:30", a time cell, or a fraction of a day -> "14:30". Empty / not understood -> "".
function normTime_(v) {
  const two = n => (n < 10 ? '0' : '') + n;
  if (v === '' || v === null || v === undefined) return '';
  if (v instanceof Date) { const t = Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm'); return t === '00:00' ? '' : t; }
  if (typeof v === 'number') { if (!(v > 0 && v < 1)) return ''; const m = Math.round(v * 1440); return two(Math.floor(m / 60) % 24) + ':' + two(m % 60); }
  const m = String(v).trim().match(/^(\d{1,2})(?:[:.](\d{2}))?\s*([ap])?\.?m?\.?$/i);
  if (!m || (!m[2] && !m[3])) return '';
  let h = +m[1]; const mi = m[2] ? +m[2] : 0, ap = (m[3] || '').toLowerCase();
  if (ap === 'p' && h < 12) h += 12;
  if (ap === 'a' && h === 12) h = 0;
  if (h > 23 || mi > 59) return '';
  return two(h) + ':' + two(mi);
}
function fmt12_(t) { const p = t.split(':').map(Number); return (p[0] % 12 || 12) + ':' + (p[1] < 10 ? '0' : '') + p[1] + ' ' + (p[0] < 12 ? 'AM' : 'PM'); }

// Adds the "Time Paid" column (K) to the Payments tab the first time it is needed. Nothing existing is moved.
function ensureTimeColumn_() {
  const s = sh(PAY);
  if (!s || String(s.getRange(4, TIMECOL).getValue()) !== '') return;
  const n = PAY.r1 - PAY.r0 + 1;
  s.getRange(3, TIMECOL).setValue('▼ TYPE HERE ▼');
  s.getRange(4, TIMECOL).setValue('Time Paid\n(optional)');
  s.getRange(3, 7, 2, 1).copyTo(s.getRange(3, TIMECOL, 2, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  s.getRange(PAY.r0, 7, n, 1).copyTo(s.getRange(PAY.r0, TIMECOL, n, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  s.getRange(PAY.r0, TIMECOL, n, 1).setNumberFormat('@');
  s.setColumnWidth(TIMECOL, 110);
}

function readSettings() {
  const out = {allowDelete: false, warnDays: 7, remindHour: 8, stopAfter: 30, pushReady: pushReady_()};
  const s = sh(SET);
  if (!s) return out;
  out.allowDelete = String(s.getRange(SET.deleteRow, SET.col).getValue()).toLowerCase() === 'yes';
  const warn = Number(s.getRange(SET.warnRow, SET.col).getValue());
  out.warnDays = warn > 0 ? warn : 7;
  const hv = s.getRange(SET.hourRow, SET.col).getValue(), sv = s.getRange(SET.stopRow, SET.col).getValue();
  if (hv !== '' && Number(hv) >= 0 && Number(hv) <= 23) out.remindHour = Math.floor(Number(hv));
  if (sv !== '' && Number(sv) >= 0) out.stopAfter = Math.floor(Number(sv));
  return out;
}

function readAll_() {
  const tenants = [], payments = [], changes = [];
  grid(TEN, 7).forEach((r, i) => { if (r[0] !== '') tenants.push({row: TEN.r0 + i, name: String(r[0]), apt: String(r[1]), phone: String(r[2]), rent: Number(r[3]) || 0, first: dstr(r[4]), status: r[5] || 'Active', notes: String(r[6])}); });
  grid(PAY, TIMECOL).forEach((r, i) => { if (r[0] !== '') payments.push({row: PAY.r0 + i, date: dstr(r[0]), time: normTime_(r[TIMECOL - 1]), tenant: String(r[1]), amount: Number(r[2]) || 0, method: String(r[3]), ref: String(r[4]), by: String(r[5]), notes: String(r[6])}); });
  if (sh(CHG)) grid(CHG, 4).forEach((r, i) => { if (r[0] !== '' && r[1] !== '') changes.push({row: CHG.r0 + i, date: dstr(r[0]), tenant: String(r[1]), rent: Number(r[2]) || 0, note: String(r[3])}); });
  return {tenants, payments, changes};
}

function handle(q) {
  if (q.action === 'all') {
    const d = readAll_();
    const names = t => grid(t, 1).map(r => r[0]).filter(v => v !== '').map(String);
    return {tenants: d.tenants, payments: d.payments, changes: d.changes, settings: readSettings(), owners: names(OWN), apts: names(APT)};
  }
  if (q.action === 'addPayment') {
    if (!q.tenant || !(Number(q.amount) > 0) || !q.date) throw new Error('Missing tenant, amount or date');
    if (q.ref && grid(PAY, 7).some(r => String(r[4]) === String(q.ref))) throw new Error('This reference number is already saved (duplicate)');
    let tm = '';
    if (q.time) { tm = normTime_(q.time); if (!tm) throw new Error('Time not understood. Use a time like 2:30 PM'); }
    ensureTimeColumn_();
    const row = emptyRow(PAY), s = sh(PAY);
    s.getRange(row, 5).setNumberFormat('@');
    s.getRange(row, 1, 1, 7).setValues([[dval(q.date), q.tenant, Number(q.amount), q.method || '', String(q.ref || ''), q.by || '', q.notes || '']]);
    s.getRange(row, TIMECOL).setNumberFormat('@').setValue(tm ? fmt12_(tm) : '');
    return {row};
  }
  if (q.action === 'deletePayment') {
    if (!readSettings().allowDelete) throw new Error('Removing payments is switched off in Settings');
    if (q.row < PAY.r0 || q.row > PAY.r1) throw new Error('Bad row');
    sh(PAY).getRange(q.row, 1, 1, 7).clearContent();
    sh(PAY).getRange(q.row, TIMECOL).clearContent();
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
    } else if (q.key === 'remindHour') {
      const n = Math.round(Number(q.value));
      if (!(n >= 0 && n <= 23)) throw new Error('Choose an hour from 0 to 23');
      ensureSetup_();
      s.getRange(SET.hourRow, SET.col).setValue(n);
      try { installTrigger_(n); }
      catch (err) { throw new Error('Time saved, but the daily timer was not moved. In Apps Script, run setupReminders once.'); }
    } else if (q.key === 'stopAfter') {
      const n = Math.round(Number(q.value));
      if (!(n >= 0 && n <= 365)) throw new Error('Choose 0 (never stop) up to 365 days');
      ensureSetup_();
      s.getRange(SET.stopRow, SET.col).setValue(n);
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
  if (q.action === 'registerDevice') {
    const token = String(q.token || '');
    if (token.length < 20) throw new Error('Missing phone code');
    ensureSetup_();
    const s = sh(DEV), rows = devRows_();
    if (q.old && q.old !== token) {
      const o = rows.findIndex(r => String(r[4]) === String(q.old));
      if (o >= 0) { s.getRange(DEV.r0 + o, 1, 1, 5).clearContent(); rows[o] = ['', '', '', '', '']; }
    }
    const i = rows.findIndex(r => String(r[4]) === token);
    if (i >= 0) {
      s.getRange(DEV.r0 + i, 2).setValue('Yes');
      return {row: DEV.r0 + i};
    }
    const base = String(q.label || 'Phone').slice(0, 30).trim() || 'Phone', used = rows.map(r => String(r[0]));
    let label = base, n = 1;
    while (used.indexOf(label) >= 0) label = base + ' ' + (++n);
    const row = emptyRow(DEV);
    s.getRange(row, 1, 1, 5).setValues([[label, 'Yes', new Date(), '', token]]);
    return {row};
  }
  if (q.action === 'unregisterDevice') {
    if (!sh(DEV)) return {};
    const i = devRows_().findIndex(r => String(r[4]) === String(q.token || '-'));
    if (i >= 0) sh(DEV).getRange(DEV.r0 + i, 2).setValue('No');
    return {};
  }
  if (q.action === 'testPush') {
    if (!q.token) throw new Error('Missing phone code');
    const r = fcmSend_(serviceAccount_(), String(q.token), 'Rentbook test', 'Reminders are working on this phone. ✓');
    if (!r.ok) throw new Error(r.gone ? 'This phone is not registered any more. Turn Reminders off and on again.' : 'Could not send: ' + r.error);
    return {};
  }
  throw new Error('Unknown action');
}

/* ---------------- Rent math (same as the app, so phone and reminders agree) ---------------- */

const DAY_ = 864e5;
function pd_(s) { const p = s.split('-').map(Number); return Date.UTC(p[0], p[1] - 1, p[2]) / DAY_; }
function edate_(s, n) {
  const p = s.split('-').map(Number), t = p[0] * 12 + p[1] - 1 + n, Y = Math.floor(t / 12), M = t % 12, L = new Date(Date.UTC(Y, M + 1, 0)).getUTCDate();
  return Y + '-' + String(M + 1).padStart(2, '0') + '-' + String(Math.min(p[2], L)).padStart(2, '0');
}
function rentAt_(t, chg, d) {
  let r = +t.rent || 0, best = '';
  for (const c of chg) { if (c.tenant === t.name && c.date <= d && c.date >= best) { best = c.date; r = +c.rent || 0; } }
  return r;
}
function calc_(t, pays, chg, today, warn) {
  let paid = 0, last = '';
  for (const p of pays) { if (p.tenant === t.name) { paid += +p.amount || 0; if (p.date > last) last = p.date; } }
  const out = t.status === 'Moved Out', rentNow = rentAt_(t, chg, today);
  if (!t.first) return {paid, last, months: 0, next: '', days: 9999, owes: 0, nextAmt: rentNow, rentNow, status: out ? 'out' : 'ok'};
  let cum = 0, dueSoFar = 0, m = -1, next = '', nextAmt = 0;
  for (let k = 0; k < 1500; k++) {
    const d = edate_(t.first, k);
    cum += rentAt_(t, chg, d);
    if (d <= today) dueSoFar = cum;
    if (m < 0 && cum > paid) { m = k; next = d; nextAmt = cum - paid; }
    if (m >= 0 && d > today) break;
  }
  if (m < 0) { m = 1500; next = edate_(t.first, 1500); nextAmt = rentNow; }
  const days = pd_(next) - pd_(today), owes = Math.max(0, dueSoFar - paid);
  const status = out ? 'out' : days < 0 ? 'late' : days === 0 ? 'today' : days <= warn ? 'soon' : 'ok';
  return {paid, last, months: m, next, days, owes, nextAmt, rentNow, status};
}
function peso_(n) { return '₱' + String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }

/* ---------------- The daily reminder ----------------
   ONE message per day, only if something needs attention:
   - late (until "stop after" days), due today, and a heads-up when rent is `warn` days away and again the day before. */
function digest_(tenants, pays, chg, today, warn, stopAfter) {
  const late = [], now = [], soon = [];
  tenants.forEach(t => {
    if (t.status === 'Moved Out' || !t.first) return;
    const c = calc_(t, pays, chg, today, warn);
    if (c.days < 0) { if (!(stopAfter > 0) || -c.days <= stopAfter) late.push({t, c}); }
    else if (c.days === 0) now.push({t, c});
    else if (c.days === warn || c.days === 1) soon.push({t, c});
  });
  const byName = (a, b) => String(a.t.name).localeCompare(String(b.t.name));
  late.sort((a, b) => a.c.days - b.c.days || byName(a, b));
  now.sort(byName);
  soon.sort((a, b) => a.c.days - b.c.days || byName(a, b));
  const all = late.concat(now, soon);
  if (!all.length) return null;
  const parts = [];
  if (late.length) parts.push(late.length + ' late');
  if (now.length) parts.push(now.length + ' due today');
  if (soon.length) parts.push(soon.length + ' due soon');
  const line = x => {
    const c = x.c, n = Math.abs(c.days);
    if (c.days < 0) return x.t.name + ': ' + n + (n === 1 ? ' day' : ' days') + ' late, owes ' + peso_(c.owes);
    if (c.days === 0) return x.t.name + ': due today, ' + peso_(c.owes > 0 ? c.owes : c.nextAmt);
    return x.t.name + ': due in ' + n + (n === 1 ? ' day' : ' days') + ', ' + peso_(c.nextAmt);
  };
  const MAX = 6;
  let body = all.slice(0, MAX).map(line).join('\n');
  if (all.length > MAX) body += '\n+' + (all.length - MAX) + ' more';
  return {title: 'Rent check: ' + parts.join(', '), body, count: all.length};
}

function todayStr_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }

function dailyReminders() { sendDaily_(false); }
function sendReminderNow() { Logger.log(sendDaily_(true)); }          // run by hand to test (ignores "already sent today")
function previewReminder() {                                          // run by hand: shows what would be sent, sends nothing
  const d = readAll_(), st = readSettings();
  Logger.log(JSON.stringify(digest_(d.tenants, d.payments, d.changes, todayStr_(), st.warnDays, st.stopAfter) || 'Nothing to report today.', null, 2));
}

function sendDaily_(force) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties(), t = todayStr_();
    if (!force && props.getProperty('LAST_SENT') === t) return 'already sent today';
    if (!pushReady_()) return 'reminders are not set up (FCM_KEY missing)';
    const d = readAll_(), st = readSettings();
    const m = digest_(d.tenants, d.payments, d.changes, t, st.warnDays, st.stopAfter);
    props.setProperty('LAST_SENT', t);
    if (!m) return 'nothing to report today';
    const s = sh(DEV);
    if (!s) return 'no phones registered';
    const sa = serviceAccount_();
    let sent = 0;
    devRows_().forEach((r, i) => {
      if (String(r[4]) === '' || String(r[1]) !== 'Yes') return;
      const res = fcmSend_(sa, String(r[4]), m.title, m.body), row = DEV.r0 + i;
      if (res.ok) { sent++; s.getRange(row, 4).setValue(new Date()); }
      else if (res.gone) s.getRange(row, 1, 1, 5).clearContent();
      else Logger.log('Could not send to ' + r[0] + ': ' + res.error);
    });
    return 'sent to ' + sent + ' phone(s): ' + m.title;
  } finally { lock.releaseLock(); }
}

/* ---------------- Phones (Devices tab) ---------------- */

function devRows_() { return sh(DEV).getRange(DEV.r0, 1, DEV.r1 - DEV.r0 + 1, 5).getValues(); }

// Safe to run any time. Adds what the reminders need to the sheet, only if it is missing.
function ensureSetup_() {
  ensureTimeColumn_();
  const ss = SpreadsheetApp.getActive(), s = ss.getSheetByName(SET.sheet);
  if (s) {
    const add = [
      [SET.hourRow, 'Send the daily reminder around this hour (0-23)', 8, 'Uses a 24-hour clock: 8 = 8 AM, 12 = noon, 17 = 5 PM. Manila time. You can also change it in the app.'],
      [SET.stopRow, 'Stop reminding about a tenant after this many days late', 30, 'So you are not told about the same very-late tenant forever. They still show in red in the app. 0 = never stop.']
    ];
    add.forEach(a => {
      if (s.getRange(a[0], 1).getValue() === '' && s.getRange(a[0], SET.col).getValue() === '') {
        s.getRange(a[0], 1, 1, 3).setValues([[a[1], a[2], a[3]]]);
        s.getRange(SET.warnRow, 1, 1, 3).copyTo(s.getRange(a[0], 1, 1, 3), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      }
    });
    s.getRange(SET.hourRow, SET.col).setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(0, 23).setAllowInvalid(false).build());
    s.getRange(SET.stopRow, SET.col).setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(0, 365).setAllowInvalid(false).build());
  }
  if (!ss.getSheetByName(DEV.sheet)) {
    const d = ss.insertSheet(DEV.sheet);
    d.getRange(1, 1).setValue('Phones that get reminders').setFontWeight('bold').setFontSize(14).setFontColor('#1F4E3D');
    d.getRange(2, 1).setValue('The app fills this in by itself when someone turns on Reminders. You may rename a phone in column A, or set Reminders to No to stop one phone. Please leave the long code alone.').setFontColor('#595959');
    d.getRange(4, 1, 1, 5).setValues([['Phone', 'Reminders', 'Added', 'Last reminder sent', 'Phone code (leave alone)']])
      .setBackground('#2E7D5B').setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true);
    d.getRange(DEV.r0, 2, DEV.r1 - DEV.r0 + 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build());
    d.getRange(DEV.r0, 3, DEV.r1 - DEV.r0 + 1, 2).setNumberFormat('d mmm yyyy h:mm AM/PM');
    d.getRange(DEV.r0, 5, DEV.r1 - DEV.r0 + 1, 1).setFontColor('#9E9E9E').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    [150, 100, 150, 160, 200].forEach((w, i) => d.setColumnWidth(i + 1, w));
    d.setFrozenRows(4);
  }
}

function installTrigger_(hour) {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'dailyReminders') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyReminders').timeBased().everyDays(1).atHour(hour).inTimezone(Session.getScriptTimeZone()).create();
}

// Run this ONCE from the Apps Script editor (see the guide). It is safe to run again.
function setupReminders() {
  ensureSetup_();
  const hour = readSettings().remindHour;
  installTrigger_(hour);
  Logger.log('Done. Daily reminder timer is set for about ' + hour + ':00 (' + Session.getScriptTimeZone() + ').');
  Logger.log(pushReady_() ? 'Phone reminders key found. Everything is ready.' : 'NEXT STEP: add the Script Property FCM_KEY (see the guide, Part 2).');
}

/* ---------------- Sending to phones (Firebase Cloud Messaging) ---------------- */

function serviceAccount_() {
  const raw = PropertiesService.getScriptProperties().getProperty('FCM_KEY');
  if (!raw) throw new Error('Reminders are not set up yet: the FCM_KEY Script Property is missing.');
  let sa;
  try { sa = JSON.parse(raw); } catch (e) { throw new Error('FCM_KEY is not the whole JSON file. Paste everything from { to }.'); }
  if (!sa.client_email || !sa.private_key || !sa.project_id) throw new Error('FCM_KEY is missing client_email, private_key or project_id.');
  return sa;
}
function pushReady_() { try { serviceAccount_(); return true; } catch (e) { return false; } }

function b64url_(s) { return Utilities.base64EncodeWebSafe(s).replace(/=+$/, ''); }
function fcmToken_(sa, force) {
  const cache = CacheService.getScriptCache();
  if (!force) { const hit = cache.get('fcm_access'); if (hit) return hit; }
  const now = Math.floor(Date.now() / 1000);
  const unsigned = b64url_(JSON.stringify({alg: 'RS256', typ: 'JWT'})) + '.' + b64url_(JSON.stringify({
    iss: sa.client_email, scope: 'https://www.googleapis.com/auth/firebase.messaging',
    aud: 'https://oauth2.googleapis.com/token', iat: now, exp: now + 3600
  }));
  const sig = b64url_(Utilities.computeRsaSha256Signature(unsigned, sa.private_key));
  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post', muteHttpExceptions: true,
    payload: {grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: unsigned + '.' + sig}
  });
  const j = JSON.parse(res.getContentText() || '{}');
  if (!j.access_token) throw new Error('Google refused the key: ' + (j.error_description || j.error || res.getResponseCode()));
  cache.put('fcm_access', j.access_token, 3000);
  return j.access_token;
}

// Returns {ok:true} or {ok:false, gone:true|false, error:'...'}. gone = that phone is no longer reachable.
function fcmSend_(sa, token, title, body) {
  const url = 'https://fcm.googleapis.com/v1/projects/' + sa.project_id + '/messages:send';
  const payload = JSON.stringify({message: {token, data: {title, body, tag: 'rentbook'}, webpush: {headers: {TTL: '43200', Urgency: 'high'}}}});
  const go = force => UrlFetchApp.fetch(url, {method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: {Authorization: 'Bearer ' + fcmToken_(sa, force)}, payload});
  let res = go(false);
  if (res.getResponseCode() === 401) res = go(true);
  const code = res.getResponseCode();
  if (code >= 200 && code < 300) return {ok: true};
  let msg = res.getContentText(), status = '';
  try { const j = JSON.parse(msg); status = (j.error && j.error.status) || ''; msg = (j.error && j.error.message) || msg; } catch (e) {}
  const gone = code === 404 || status === 'NOT_FOUND' || status === 'UNREGISTERED' || (status === 'INVALID_ARGUMENT' && /registration token/i.test(msg));
  return {ok: false, gone, error: status + ' ' + msg};
}
