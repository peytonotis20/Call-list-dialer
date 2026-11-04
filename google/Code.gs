/**
 * Call List CRM: the Google Sheet behind the Call List Dialer.
 *
 * Paste this into Extensions > Apps Script of an empty Google Sheet, set TEAM_KEY,
 * run setUp() once, then Deploy > New deployment > Web app (Execute as: Me,
 * Who has access: Anyone). The dialer talks to the web app address; nothing
 * works without TEAM_KEY. Full steps: https://peytonotis20.github.io/Call-list-dialer/crm-setup.html
 */

// ---- Settings -------------------------------------------------------------
const TEAM_KEY = 'PASTE-YOUR-TEAM-KEY-HERE'; // the key everyone types into the dialer
const REPORT_TO = '';      // who gets the 5pm email; blank = you. Several: 'a@x.com, b@y.com'
const REPORT_HOUR = 17;    // 17 = 5pm, in this project's time zone (Project Settings)
const SKIP_EMPTY_DAYS = true; // no email on days with no calls or changes

// ---- Sheet layout ---------------------------------------------------------
const CONTACTS = 'Contacts', ACTIVITY = 'Activity', REPORTS = 'Daily Reports';
const CRM_COLS = ['ID', 'Stage', 'Owner', 'Follow-up At', 'Follow-up Note', 'Follow-up Done', 'Requested', 'Offer',
  'Funded Amount', 'Lender', 'Commission', 'Funded Date', 'Last Outcome', 'Numbers Tried', 'Attempts', 'Last Called',
  'Call Notes', 'Match Key', 'Deleted', 'Created At', 'Updated At', 'Updated By'];
const ACT_COLS = ['Time', 'Date', 'User', 'Contact ID', 'Contact', 'Business', 'Type', 'Number', 'Number Label',
  'Outcome', 'Detail', 'Event ID', 'Synced At'];
const REPORT_COLS = ['Date', 'Calls', 'Contacts Called', 'Connected', 'App Outs', 'Call Backs', 'Funded Deals', 'Funded $'];
// Dialer field name -> Contacts column
const FIELDS = { stage: 'Stage', owner: 'Owner', fuAt: 'Follow-up At', fuNote: 'Follow-up Note', fuDone: 'Follow-up Done',
  requested: 'Requested', offer: 'Offer', funded: 'Funded Amount', lender: 'Lender', commission: 'Commission',
  fundedOn: 'Funded Date', status: 'Last Outcome', tried: 'Numbers Tried', attempts: 'Attempts', lastCalled: 'Last Called',
  notes: 'Call Notes', matchKey: 'Match Key', deleted: 'Deleted' };
const TIME_FIELDS = { fuAt: 1, lastCalled: 1 };
const TEXT_COLS = ['ID', 'Numbers Tried', 'Call Notes', 'Match Key', 'Follow-up Note', 'Owner', 'Lender', 'Funded Date'];
const ACTIVITY_DAYS_ON_FIRST_SYNC = 120;

// ---- Web app --------------------------------------------------------------
function doGet() {
  return ContentService.createTextOutput('The Call List CRM is running. Paste this address into the dialer\'s settings.');
}

function doPost(e) {
  let req = {};
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return reply({ ok: false, error: 'Bad request.' }); }
  if (TEAM_KEY === 'PASTE-YOUR-TEAM-KEY-HERE') return reply({ ok: false, error: 'Set TEAM_KEY at the top of the Apps Script code, then deploy again.' });
  if (String(req.key || '') !== TEAM_KEY) return reply({ ok: false, error: 'Wrong team key.' });
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const user = String(req.user || 'Someone').slice(0, 60);
    if (req.action === 'ping') return reply(ping());
    if (req.action === 'pull') return reply(pull(req.since || null));
    if (req.action === 'push') return reply(push(req.events || [], user));
    if (req.action === 'import') return reply(importContacts(req.contacts || [], req.headers || [], user));
    return reply({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    return reply({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (err) {}
  }
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---- Setup ----------------------------------------------------------------
/** Run once from the Apps Script editor: creates the tabs, the 5pm trigger and sends a test email. */
function setUp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  ensureSheets();
  ensureTriggers();
  const tz = Session.getScriptTimeZone();
  MailApp.sendEmail({
    to: reportRecipients(),
    subject: 'Your Call List CRM is set up',
    htmlBody: '<p>Your Call List CRM is ready.</p><p>The daily call report will arrive at about ' + hourLabel(REPORT_HOUR) +
      ' (' + tz + ') on days with calls.</p><p>Next: Deploy &gt; New deployment &gt; Web app, then paste its address into the dialer.</p>',
  });
  Logger.log('Set up. Report time: ' + hourLabel(REPORT_HOUR) + ' ' + tz + '. Email: ' + reportRecipients());
}

/** Sends today's report right now (for testing). */
function sendReportNow() { sendDailyReport(true); }

function ensureTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const has = name => triggers.some(t => t.getHandlerFunction() === name);
  // Always rebuilt, so running setUp again picks up a new REPORT_HOUR or project time zone.
  triggers.filter(t => t.getHandlerFunction() === 'sendDailyReport').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendDailyReport').timeBased().atHour(REPORT_HOUR).nearMinute(0).everyDays(1)
    .inTimezone(Session.getScriptTimeZone()).create();
  if (!has('onSheetEdit')) ScriptApp.newTrigger('onSheetEdit').forSpreadsheet(book()).onEdit().create();
}

function book() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function ensureSheets() {
  const ss = book();
  const make = (name, cols) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      ensureSize(sh, 1, cols.length);
      sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    return sh;
  };
  const c = make(CONTACTS, CRM_COLS);
  TEXT_COLS.forEach(h => { const k = CRM_COLS.indexOf(h); if (k >= 0) c.getRange(1, k + 1, c.getMaxRows(), 1).setNumberFormat('@'); });
  const a = make(ACTIVITY, ACT_COLS);
  ['Date', 'Number', 'Detail', 'Event ID', 'Contact ID'].forEach(h => a.getRange(1, ACT_COLS.indexOf(h) + 1, a.getMaxRows(), 1).setNumberFormat('@'));
  make(REPORTS, REPORT_COLS);
  const first = ss.getSheets()[0];
  if (first.getName() === 'Sheet1' && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
}

// Edits made by hand in the Contacts tab are stamped so every dialer picks them up.
function onSheetEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== CONTACTS || e.range.getRow() < 2) return;
    const head = headerRow(sh), upd = head.indexOf('Updated At') + 1, by = head.indexOf('Updated By') + 1;
    if (e.range.getColumn() === upd || e.range.getColumn() === by) return;
    const n = e.range.getNumRows(), now = new Date();
    sh.getRange(e.range.getRow(), upd, n, 1).setValues(Array.from({ length: n }, () => [now]));
    sh.getRange(e.range.getRow(), by, n, 1).setValues(Array.from({ length: n }, () => ['Google Sheet']));
    touch();
  } catch (err) {}
}

// ---- Sync -----------------------------------------------------------------
function ping() {
  ensureSheets();
  try { ensureTriggers(); } catch (err) {}
  const c = book().getSheetByName(CONTACTS);
  return { ok: true, contacts: Math.max(0, c.getLastRow() - 1), reportTo: reportRecipients(),
    timeZone: Session.getScriptTimeZone(), reportTime: hourLabel(REPORT_HOUR), sheet: book().getName() };
}

function touch() { PropertiesService.getScriptProperties().setProperty('LAST_CHANGE', new Date().toISOString()); }

function pull(since) {
  ensureSheets();
  const now = new Date().toISOString();
  const last = PropertiesService.getScriptProperties().getProperty('LAST_CHANGE');
  if (since && last && since > last) return { ok: true, now, contacts: [], activity: [], unchanged: true };
  const ss = book(), csh = ss.getSheetByName(CONTACTS), ash = ss.getSheetByName(ACTIVITY);
  const head = headerRow(csh), leadHeads = head.slice(CRM_COLS.length);
  const sinceMs = since ? Date.parse(since) : 0;
  const rows = csh.getLastRow() > 1 ? csh.getRange(2, 1, csh.getLastRow() - 1, head.length).getValues() : [];
  const ci = name => head.indexOf(name);
  const contacts = [];
  const owners = {};
  rows.forEach((r, k) => {
    if (!r[ci('ID')]) return;
    if (r[ci('Owner')]) owners[r[ci('Owner')]] = 1;
    const upd = r[ci('Updated At')] instanceof Date ? r[ci('Updated At')].getTime() : Date.parse(r[ci('Updated At')]) || 0;
    if (since && upd < sinceMs) return;
    const c = { id: String(r[ci('ID')]), row: k + 2, updatedAt: iso(r[ci('Updated At')]), updatedBy: r[ci('Updated By')] || '' };
    Object.keys(FIELDS).forEach(f => { c[f] = fromCell(f, r[ci(FIELDS[f])]); });
    c.lead = leadHeads.map((h, j) => cellText(r[CRM_COLS.length + j]));
    contacts.push(c);
  });
  const acts = [];
  const users = {};
  const alast = ash.getLastRow();
  if (alast > 1) {
    const av = ash.getRange(2, 1, alast - 1, ACT_COLS.length).getValues();
    const minMs = since ? sinceMs : Date.now() - ACTIVITY_DAYS_ON_FIRST_SYNC * 864e5;
    av.forEach(r => {
      if (r[2]) users[r[2]] = 1;
      const synced = r[12] instanceof Date ? r[12].getTime() : Date.parse(r[12]) || 0;
      const at = r[0] instanceof Date ? r[0].getTime() : Date.parse(r[0]) || 0;
      if (since ? synced < minMs : at < minMs) return;
      acts.push({ at: iso(r[0]), user: r[2], cid: String(r[3]), type: r[6], num: cellText(r[7]), label: r[8], outcome: r[9], detail: cellText(r[10]), id: String(r[11]) });
    });
  }
  return { ok: true, now, full: !since, headers: leadHeads, contacts, activity: acts, team: Object.keys(Object.assign(users, owners)).sort() };
}

function push(events, user) {
  ensureSheets();
  const ss = book(), csh = ss.getSheetByName(CONTACTS), ash = ss.getSheetByName(ACTIVITY);
  const seen = recentEventIds(ash);
  let head = headerRow(csh);
  const idCol = head.indexOf('ID');
  const ids = csh.getLastRow() > 1 ? csh.getRange(2, idCol + 1, csh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  const rowOf = {}; ids.forEach((id, k) => { if (id) rowOf[id] = k + 2; });
  const now = new Date(), actRows = [];
  let applied = 0, skipped = 0;
  events.forEach(ev => {
    if (!ev || !ev.id || seen[ev.id]) { skipped++; return; }
    seen[ev.id] = 1;
    const who = String(ev.user || user);
    if (ev.lead) head = addLeadHeaders(csh, head, Object.keys(ev.lead));
    let row = rowOf[ev.cid];
    if (!row && ev.type === 'Added') {
      row = csh.getLastRow() + 1; ensureSize(csh, row, head.length);
      textFormat(csh, row, 1, head.length);
      csh.getRange(row, idCol + 1).setValue(ev.cid);
      csh.getRange(row, head.indexOf('Stage') + 1).setValue('New');
      csh.getRange(row, head.indexOf('Created At') + 1).setValue(now);
      rowOf[ev.cid] = row;
    }
    if (row) {
      const cur = csh.getRange(row, 1, 1, head.length).getValues()[0];
      Object.keys(ev.set || {}).forEach(f => { if (FIELDS[f]) cur[head.indexOf(FIELDS[f])] = toCell(f, ev.set[f]); });
      Object.keys(ev.lead || {}).forEach(h => { cur[head.indexOf(h)] = ev.lead[h]; });
      cur[head.indexOf('Updated At')] = now; cur[head.indexOf('Updated By')] = who;
      csh.getRange(row, 1, 1, head.length).setValues([cur]);
    }
    applied++;
    if (ev.quiet) return; // e.g. call notes saved while typing: update the contact, no activity line
    const a = ev.act || {};
    actRows.push([toDate(ev.at) || now, dateKey(toDate(ev.at) || now), who, ev.cid || '', a.contact || '', a.business || '', ev.type || '',
      a.num || '', a.label || '', a.outcome || '', a.detail || '', ev.id, now]);
  });
  appendRows(ash, actRows);
  if (applied) touch();
  return { ok: true, applied, skipped, now: now.toISOString() };
}

function importContacts(list, headers, user) {
  ensureSheets();
  const ss = book(), csh = ss.getSheetByName(CONTACTS), ash = ss.getSheetByName(ACTIVITY);
  let head = addLeadHeaders(csh, headerRow(csh), headers);
  const last = csh.getLastRow();
  const existing = last > 1 ? csh.getRange(2, 1, last - 1, head.length).getValues() : [];
  const mk = head.indexOf('Match Key'), idc = head.indexOf('ID');
  const byKey = {}, byId = {};
  existing.forEach(r => {
    String(r[mk] || '').split(' ').filter(Boolean).forEach(k => { byKey[k] = String(r[idc]); });
    byId[String(r[idc])] = 1;
  });
  const seen = recentEventIds(ash);
  const now = new Date(), rows = [], acts = [], mapping = {};
  let added = 0, skipped = 0;
  list.forEach(c => {
    const keys = String(c.matchKey || '').split(' ').filter(Boolean);
    const dup = byId[c.id] ? c.id : keys.map(k => byKey[k]).find(Boolean);
    if (dup) { mapping[c.id] = dup; skipped++; return; }
    const row = head.map(() => '');
    row[idc] = c.id;
    row[mk] = keys.join(' ');
    Object.keys(c.set || {}).forEach(f => { if (FIELDS[f]) row[head.indexOf(FIELDS[f])] = toCell(f, c.set[f]); });
    if (!row[head.indexOf('Stage')]) row[head.indexOf('Stage')] = 'New';
    Object.keys(c.lead || {}).forEach(h => { const k = head.indexOf(h); if (k >= 0) row[k] = c.lead[h]; });
    row[head.indexOf('Created At')] = now; row[head.indexOf('Updated At')] = now; row[head.indexOf('Updated By')] = user;
    rows.push(row);
    keys.forEach(k => { byKey[k] = c.id; });
    byId[c.id] = 1;
    (c.events || []).forEach(ev => {
      if (!ev.id || seen[ev.id]) return;
      seen[ev.id] = 1;
      const a = ev.act || {};
      acts.push([toDate(ev.at) || now, dateKey(toDate(ev.at) || now), ev.user || user, c.id, a.contact || '', a.business || '', ev.type || '',
        a.num || '', a.label || '', a.outcome || '', a.detail || '', ev.id, now]);
    });
    added++;
  });
  appendRows(csh, rows);
  appendRows(ash, acts);
  if (added) touch();
  return { ok: true, added, skipped, mapping };
}

// ---- 5pm report -----------------------------------------------------------
function sendDailyReport(force) {
  force = force === true;   // the 5pm timer passes an event object here; only sendReportNow forces a send
  const ss = book(), tz = Session.getScriptTimeZone();
  const today = dateKey(new Date());
  const tomorrow = dateKey(new Date(Date.now() + 864e5));
  const ash = ss.getSheetByName(ACTIVITY), csh = ss.getSheetByName(CONTACTS);
  const acts = ash.getLastRow() > 1 ? ash.getRange(2, 1, ash.getLastRow() - 1, ACT_COLS.length).getValues() : [];
  // Filter on each row's exact timestamp, so the spreadsheet's own time zone setting can't shift a call to another day.
  const todays = acts.filter(r => r[0] instanceof Date ? dateKey(r[0]) === today : String(r[1]).slice(0, 10) === today);
  const calls = todays.filter(r => r[6] === 'Call');
  const head = headerRow(csh);
  const cRows = csh.getLastRow() > 1 ? csh.getRange(2, 1, csh.getLastRow() - 1, head.length).getValues() : [];
  const col = name => head.indexOf(name);
  const live = cRows.filter(r => r[col('ID')] && r[col('Deleted')] !== true && String(r[col('Deleted')]).toUpperCase() !== 'TRUE');
  const funded = live.filter(r => r[col('Funded Date')] && (r[col('Funded Date')] instanceof Date ? dateKey(r[col('Funded Date')]) : String(r[col('Funded Date')]).slice(0, 10)) === today);
  const dueTomorrow = live.filter(r => r[col('Follow-up At')] instanceof Date && dateKey(r[col('Follow-up At')]) === tomorrow && String(r[col('Follow-up Done')]).toUpperCase() !== 'TRUE');
  const stageMoves = todays.filter(r => r[6] === 'Stage');
  if (!force && SKIP_EMPTY_DAYS && !calls.length && !stageMoves.length && !funded.length) return;

  const is = o => r => r[9] === o;
  const appOuts = calls.filter(is('App Out')), connected = calls.filter(is('Connected')), callbacks = calls.filter(is('Call back'));
  const reached = calls.filter(r => r[9] === 'Connected' || r[9] === 'App Out');
  const contactsCalled = Object.keys(calls.reduce((m, r) => { m[r[3]] = 1; return m; }, {})).length;
  const fundedTotal = funded.reduce((s, r) => s + (Number(r[col('Funded Amount')]) || 0), 0);
  const people = {};
  calls.forEach(r => {
    const p = people[r[2]] = people[r[2]] || { calls: 0, reached: 0, appOuts: 0, callbacks: 0 };
    p.calls++; if (r[9] === 'Connected' || r[9] === 'App Out') p.reached++; if (r[9] === 'App Out') p.appOuts++; if (r[9] === 'Call back') p.callbacks++;
  });
  const stages = {};
  live.forEach(r => { const s = r[col('Stage')] || 'New'; stages[s] = (stages[s] || 0) + 1; });
  // Display names come from the Activity tab (the dialer always fills them in), else from the lead's own columns.
  const names = {};
  acts.forEach(r => { if (r[4] || r[5]) names[String(r[3])] = [r[4], r[5]]; });
  const who = r => names[String(r[col('ID')])] || [r[col('Contact Name')] || r[col('Name')] || '', r[col('Business')] || r[col('Company')] || ''];

  const money = n => '$' + Math.round(n).toLocaleString('en-US');
  const time = d => d instanceof Date ? Utilities.formatDate(d, tz, 'h:mm a') : '';
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const th = 'style="text-align:left;padding:6px 10px;border-bottom:1px solid #d3dcd8;font-size:12px;color:#5a6a64"';
  const td = 'style="padding:6px 10px;border-bottom:1px solid #eef2f0;vertical-align:top"';
  const table = (cols, rows, none) => rows.length ? '<table style="border-collapse:collapse;font:14px Arial,sans-serif;width:100%"><tr>' +
    cols.map(c => '<th ' + th + '>' + esc(c) + '</th>').join('') + '</tr>' +
    rows.map(r => '<tr>' + r.map(v => '<td ' + td + '>' + esc(v) + '</td>').join('') + '</tr>').join('') + '</table>' : '<p style="color:#5a6a64">' + (none || 'None today.') + '</p>';
  const callRows = list => list.map(r => [time(r[0]), r[4], r[5], r[7], r[2], r[10]]);
  const tile = (n, l) => '<td style="padding:10px 14px;border:1px solid #d3dcd8;border-radius:8px;text-align:center"><div style="font:600 22px Arial">' + n + '</div><div style="font:12px Arial;color:#5a6a64">' + l + '</div></td>';
  const dayLabel = Utilities.formatDate(new Date(), tz, 'EEEE, MMMM d');
  const html = '<div style="font:14px Arial,sans-serif;color:#15201c;max-width:760px">' +
    '<h2 style="margin:0 0 4px">Call report: ' + esc(dayLabel) + '</h2>' +
    '<table style="border-spacing:8px;margin:8px -8px"><tr>' + tile(calls.length, 'calls') + tile(contactsCalled, 'contacts called') +
    tile(reached.length, 'connected') + tile(appOuts.length, 'App Outs') + tile(callbacks.length, 'call backs') +
    tile(funded.length ? money(fundedTotal) : '0', 'funded today') + '</tr></table>' +
    (Object.keys(people).length
      ? '<h3>By person</h3>' + table(['Person', 'Calls', 'Connected', 'App Outs', 'Call backs'],
        Object.keys(people).sort().map(n => [n, people[n].calls, people[n].reached, people[n].appOuts, people[n].callbacks])) : '') +
    '<h3>App Outs</h3>' + table(['Time', 'Contact', 'Business', 'Number', 'By', 'Notes'], callRows(appOuts)) +
    '<h3>Connected</h3>' + table(['Time', 'Contact', 'Business', 'Number', 'By', 'Notes'], callRows(connected)) +
    '<h3>Call backs</h3>' + table(['Time', 'Contact', 'Business', 'Number', 'By', 'Notes'], callRows(callbacks)) +
    (funded.length ? '<h3>Funded today</h3>' + table(['Contact', 'Business', 'Funded', 'Lender', 'Owner'],
      funded.map(r => [who(r)[0], who(r)[1], money(Number(r[col('Funded Amount')]) || 0), r[col('Lender')], r[col('Owner')]])) : '') +
    (stageMoves.length ? '<h3>Stage changes</h3>' + table(['Time', 'Contact', 'Business', 'Change', 'By'], stageMoves.map(r => [time(r[0]), r[4], r[5], r[10], r[2]])) : '') +
    '<h3>Follow-ups due tomorrow</h3>' + table(['Time', 'Contact', 'Business', 'Owner', 'Note'],
      dueTomorrow.sort((a, b) => a[col('Follow-up At')] - b[col('Follow-up At')]).map(r => [time(r[col('Follow-up At')]), who(r)[0], who(r)[1], r[col('Owner')], r[col('Follow-up Note')]]), 'None scheduled.') +
    '<h3>Pipeline now</h3>' + table(['Stage', 'Contacts'], ['New', 'Contacted', 'App Out', 'Docs In', 'Submitted', 'Approved', 'Funded', 'Dead'].filter(s => stages[s]).map(s => [s, stages[s]])) +
    '<p style="color:#5a6a64;font-size:12px">Every call from today is in the attached spreadsheet and in the Activity tab of ' + esc(ss.getName()) + '.</p></div>';
  const csvCell = v => { const s = v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm') : String(v == null ? '' : v); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
  const csv = [ACT_COLS.slice(0, 11)].concat(todays.map(r => r.slice(0, 11))).map(r => r.map(csvCell).join(',')).join('\r\n');
  MailApp.sendEmail({
    to: reportRecipients(),
    subject: 'Call report ' + Utilities.formatDate(new Date(), tz, 'EEE, MMM d') + ': ' + calls.length + ' calls, ' + appOuts.length + ' App Out' + (appOuts.length === 1 ? '' : 's'),
    htmlBody: html,
    attachments: [Utilities.newBlob('﻿' + csv, 'text/csv', 'call-report-' + today + '.csv')],
  });
  appendRows(ss.getSheetByName(REPORTS), [[today, calls.length, contactsCalled, reached.length, appOuts.length, callbacks.length, funded.length, fundedTotal]]);
}

// ---- Helpers --------------------------------------------------------------
function reportRecipients() { return REPORT_TO || Session.getEffectiveUser().getEmail(); }
function hourLabel(h) { return (h % 12 || 12) + ':00 ' + (h < 12 ? 'AM' : 'PM'); }
function dateKey(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function toDate(v) { if (!v) return null; const d = v instanceof Date ? v : new Date(v); return isNaN(d.getTime()) ? null : d; }
function iso(v) { const d = toDate(v); return d ? d.toISOString() : ''; }
function cellText(v) { return v instanceof Date ? dateKey(v) : v == null ? '' : String(v); }
function headerRow(sh) { return sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String); }

function toCell(f, v) {
  if (v === null || v === undefined) return '';
  if (TIME_FIELDS[f]) return toDate(v) || '';
  if (f === 'tried') return typeof v === 'string' ? v : JSON.stringify(v);
  if (f === 'fuDone' || f === 'deleted') return v === true || v === 'true';
  if (['requested', 'offer', 'funded', 'commission', 'attempts'].indexOf(f) >= 0) return v === '' ? '' : Number(v) || 0;
  return v;
}
function fromCell(f, v) {
  if (TIME_FIELDS[f]) return iso(v);
  if (f === 'fundedOn') return v instanceof Date ? dateKey(v) : String(v || '');
  if (f === 'tried') { try { return v ? JSON.parse(v) : {}; } catch (e) { return {}; } }
  if (f === 'fuDone' || f === 'deleted') return v === true || String(v).toUpperCase() === 'TRUE';
  if (['requested', 'offer', 'funded', 'commission', 'attempts'].indexOf(f) >= 0) return v === '' || v == null ? '' : Number(v) || 0;
  return v == null ? '' : String(v);
}

function ensureSize(sh, rows, cols) {
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
}
function appendRows(sh, rows) {
  if (!rows.length) return;
  const start = sh.getLastRow() + 1, width = rows[0].length;
  ensureSize(sh, start + rows.length - 1, width);
  textFormat(sh, start, rows.length, width);
  sh.getRange(start, 1, rows.length, width).setValues(rows);
}
// Plain-text format for ID, phone, ZIP and note columns, so Sheets never turns "07442" into 7442.
function textFormat(sh, row, n, width) {
  const name = sh.getName(), cols = [];
  if (name === CONTACTS) {
    TEXT_COLS.forEach(h => cols.push(CRM_COLS.indexOf(h) + 1));
    for (let c = CRM_COLS.length + 1; c <= width; c++) cols.push(c);
  } else if (name === ACTIVITY) {
    ['Date', 'Number', 'Detail', 'Event ID', 'Contact ID'].forEach(h => cols.push(ACT_COLS.indexOf(h) + 1));
  }
  cols.filter(c => c >= 1 && c <= width).forEach(c => sh.getRange(row, c, n, 1).setNumberFormat('@'));
}
function addLeadHeaders(sh, head, names) {
  const fresh = names.filter(h => h && head.indexOf(h) < 0 && CRM_COLS.indexOf(h) < 0);
  if (!fresh.length) return head;
  const at = head.length + 1;
  ensureSize(sh, 1, head.length + fresh.length);
  sh.getRange(1, at, 1, fresh.length).setValues([fresh]).setFontWeight('bold');
  sh.getRange(1, at, sh.getMaxRows(), fresh.length).setNumberFormat('@'); // keep ZIPs and phone numbers as typed
  return head.concat(fresh);
}
function recentEventIds(ash) {
  const last = ash.getLastRow(), seen = {};
  if (last < 2) return seen;
  const from = Math.max(2, last - 5000);
  ash.getRange(from, 12, last - from + 1, 1).getValues().forEach(r => { if (r[0]) seen[String(r[0])] = 1; });
  return seen;
}
