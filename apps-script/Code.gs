/**
 * Energy Management Market Map — shared team backend (Google Apps Script)
 *
 * Setup (once):
 *   1. Create a Google Sheet, then Extensions → Apps Script. Replace the sample code with this file.
 *   2. Change TEAM_PASSCODE below to a passcode for your group, then Save.
 *   3. Deploy → New deployment → type "Web app". Execute as: Me. Who has access: Anyone. Deploy and authorise.
 *   4. Copy the Web app URL (ends in /exec) into the dashboard: Investor Notes & Settings → Team sharing.
 *
 * Anyone with the web-app URL can READ the map data. Only requests carrying the passcode can CHANGE it.
 * The script creates three tabs: "Companies", "To Research" and "Settings". You can also edit rows
 * directly in the sheet (leave the id column blank for new rows); extra columns you add are preserved.
 */
var TEAM_PASSCODE = 'change-me';

var COMPANY_TAB = 'Companies';
var RESEARCH_TAB = 'To Research';
var SETTINGS_TAB = 'Settings';
var COMPANY_FIELDS = ['id', 'name', 'url', 'description', 'businessModel', 'segment', 'quadrant', 'x', 'y', 'tier', 'stage', 'hq', 'founded',
  'employees', 'revenueM', 'fundingM', 'impactNote', 'dataNote', 'sources', 'updatedAt', 'updatedBy'];
var RESEARCH_FIELDS = ['id', 'name', 'url', 'note', 'foundVia', 'quadrant', 'priority', 'status', 'addedAt', 'updatedAt', 'mapId', 'addedBy'];

function doGet() {
  try {
    return json_({
      ok: true,
      companies: readRows_(COMPANY_TAB, COMPANY_FIELDS),
      research: readRows_(RESEARCH_TAB, RESEARCH_FIELDS),
      settings: readSettings_(),
      sheetUrl: ss_().getUrl(),
      serverTime: new Date().toISOString()
    });
  } catch (err) {
    // Report problems as readable JSON (an uncaught error returns a page the dashboard cannot read).
    return json_({ ok: false, error: 'Script error: ' + String((err && err.message) || err) });
  }
}

/* The spreadsheet this script stores data in. Works whether the script was created from inside a Sheet
   (Extensions → Apps Script) or on its own at script.google.com — in that case it creates a Sheet once. */
function ss_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  var created = SpreadsheetApp.create('EMS Market Map — Team data');
  props.setProperty('SHEET_ID', created.getId());
  return created;
}

/* Optional: run this once from the editor (select "authorize" → Run) to approve permissions and
   see the sheet's address in the log. */
function authorize() {
  Logger.log('Team data sheet: ' + ss_().getUrl());
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (TEAM_PASSCODE === 'change-me') return json_({ ok: false, error: 'Set TEAM_PASSCODE in the Apps Script, then redeploy.' });
    if (String(body.key || '') !== TEAM_PASSCODE) return json_({ ok: false, error: 'Wrong team passcode.' });
    var who = String(body.who || '').slice(0, 60);
    if (body.action === 'replaceAll') {
      replaceRows_(COMPANY_TAB, COMPANY_FIELDS, stamp_(body.companies, 'updatedBy', who, true));
      replaceRows_(RESEARCH_TAB, RESEARCH_FIELDS, stamp_(body.research, 'addedBy', who, false));
      if (body.settings) writeSettings_(body.settings);
    } else if (body.action === 'sync') {
      var c = body.companies || {}, r = body.research || {};
      applyChanges_(COMPANY_TAB, COMPANY_FIELDS, stamp_(c.upsert, 'updatedBy', who, true), c['delete']);
      applyChanges_(RESEARCH_TAB, RESEARCH_FIELDS, stamp_(r.upsert, 'addedBy', who, false), r['delete']);
      if (body.settings) writeSettings_(body.settings);
    } else {
      return json_({ ok: false, error: 'Unknown action.' });
    }
    SpreadsheetApp.flush();
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    lock.releaseLock();
  }
}

function stamp_(items, field, who, overwrite) {
  items = items || [];
  if (!who) return items;
  items.forEach(function (it) { if (overwrite || !it[field]) it[field] = who; });
  return items;
}

function tab_(name, fields) {
  var ss = ss_();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, fields.length).setValues([fields]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* Header row as an array; adds any missing standard columns at the end. */
function header_(sh, fields) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var head = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function (h) { return String(h).trim(); });
  fields.forEach(function (f) {
    if (head.indexOf(f) < 0) {
      var empty = head.indexOf('');
      var col = empty >= 0 ? empty : head.length;
      head[col] = f;
      sh.getRange(1, col + 1).setValue(f).setFontWeight('bold');
    }
  });
  return head;
}

function readRows_(name, fields) {
  var sh = tab_(name, fields);
  var head = header_(sh, fields);
  if (sh.getLastRow() < 2) return [];
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getDisplayValues();
  return values
    .filter(function (row) { return row.some(function (v) { return String(v).trim() !== ''; }); })
    .map(function (row) {
      var o = {};
      head.forEach(function (h, i) { if (h) o[h] = row[i]; });
      return o;
    });
}

function cell_(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Array]') return v.join(' ; ');
  return String(v);
}

/* Builds a row; standard fields come from the item, other (user-added) columns keep their old value. */
function toRow_(head, fields, item, oldRow) {
  return head.map(function (h, k) {
    if (fields.indexOf(h) >= 0) return cell_(item[h]);
    return oldRow ? oldRow[k] : '';
  });
}

function applyChanges_(name, fields, upserts, deletes) {
  var sh = tab_(name, fields);
  var head = header_(sh, fields);
  var idCol = head.indexOf('id'), nameCol = head.indexOf('name');
  var n = Math.max(sh.getLastRow() - 1, 0);
  var data = n ? sh.getRange(2, 1, n, head.length).getDisplayValues() : [];
  function norm(s) { return String(s || '').trim().toLowerCase(); }
  function find(item) {
    var i;
    if (item.id) for (i = 0; i < data.length; i++) if (data[i][idCol] === String(item.id)) return i;
    for (i = 0; i < data.length; i++) if (!data[i][idCol] && norm(data[i][nameCol]) && norm(data[i][nameCol]) === norm(item.name)) return i;
    return -1;
  }
  (deletes || []).forEach(function (d) {
    var i = find(d);
    if (i >= 0) { sh.deleteRow(i + 2); data.splice(i, 1); }
  });
  (upserts || []).forEach(function (item) {
    var i = find(item);
    var row = toRow_(head, fields, item, i >= 0 ? data[i] : null);
    var r = i >= 0 ? i + 2 : data.length + 2;
    sh.getRange(r, 1, 1, head.length).setNumberFormat('@').setValues([row]);
    if (i >= 0) data[i] = row; else data.push(row);
  });
}

function replaceRows_(name, fields, items) {
  var sh = tab_(name, fields);
  var head = header_(sh, fields);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), head.length)).clearContent();
  if (!items || !items.length) return;
  var rows = items.map(function (it) { return toRow_(head, fields, it, null); });
  sh.getRange(2, 1, rows.length, head.length).setNumberFormat('@').setValues(rows);
}

function readSettings_() {
  var sh = ss_().getSheetByName(SETTINGS_TAB);
  if (!sh) return null;
  var v = String(sh.getRange('A2').getValue() || '');
  try { return v ? JSON.parse(v) : null; } catch (e) { return null; }
}

function writeSettings_(settings) {
  var ss = ss_();
  var sh = ss.getSheetByName(SETTINGS_TAB) || ss.insertSheet(SETTINGS_TAB);
  sh.getRange('A1').setValue('Map labels (JSON, managed by the dashboard — do not edit by hand)').setFontWeight('bold');
  sh.getRange('A2').setNumberFormat('@').setValue(JSON.stringify(settings));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
