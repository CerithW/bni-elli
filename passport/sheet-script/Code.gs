/**
 * Passport to Success – chapter store
 *
 * Lives inside a Google Sheet (Extensions → Apps Script) and lets the Passport web page
 * read and write that sheet. Three tabs:
 *   Members  – one row per new member in the programme (name, password, dates, mentors)
 *   Archive  – members whose certificate has been presented
 *   Settings – chapter name, coordinator password, mentors per session
 *
 * One-time setup: paste this file in, run `setup` once (approve the permissions it asks for),
 * then Deploy → New deployment → Web app, "Execute as: Me", "Who has access: Anyone".
 * Paste the web app URL into CHAPTER.apiUrl in the Passport page.
 */

var MEMBERS = 'Members', ARCHIVE = 'Archive', SETTINGS = 'Settings';
var SESSION_KEYS = ['1','2','3','4','5','6','7','8','9','10'];
var ITEM_KEYS = ['1','2','3','4','LP','5','6','7','8','9','10','RT','MSP'];
var BASE = ['Name','Password','Classification','Inducted'];
var TAIL = ['Completed','Notes'];

function label(k){ return /^\d+$/.test(k) ? 'S' + k : k; }          // '1' → 'S1', 'MSP' → 'MSP'
function keyFromLabel(h){ var m = /^S(\d+)$/.exec(h); return m ? m[1] : h; }

/* ---------------- One-time setup ---------------- */
function setup(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var members = ss.getSheetByName(MEMBERS) || ss.insertSheet(MEMBERS);
  var archive = ss.getSheetByName(ARCHIVE) || ss.insertSheet(ARCHIVE);
  var settings = ss.getSheetByName(SETTINGS) || ss.insertSheet(SETTINGS);

  var headers = BASE.concat(['S1','S2','S3','S4','S5','S6','S7','S8','S9','S10','MSP'])
                    .concat(['Mentor S1','Mentor S2','Mentor S3','Mentor S4','Mentor S5','Mentor S6','Mentor S7','Mentor S8','Mentor S9','Mentor S10','Mentor MSP'])
                    .concat(TAIL);
  [members, archive].forEach(function(sh){
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1); sh.setFrozenColumns(1);
      sh.getRange(2, 1, 500, headers.length).setNumberFormat('@');   // keep dates as plain text, e.g. 2026-10-22
    }
  });

  if (settings.getLastRow() === 0) {
    var rows = [
      ['Setting', 'Value', 'Notes'],
      ['Chapter', 'BNI Elli', 'Shown at the top of the page'],
      ['Coordinator password', 'coord-' + rand4(), 'Change this to whatever you like. Needed to see every member.'],
      ['Extra items', '', 'Leave blank for 10 sessions + MSP. Put LP and/or RT here to also track Launch Pad and the Regional Team sign-off.'],
      ['Session 1 mentors', 'Matthew, Eleanor', 'Usual mentor first, then the alternative'],
      ['Session 2 mentors', 'Rhys D, Cerith', ''],
      ['Session 3 mentors', 'James R, David D', ''],
      ['Session 4 mentors', 'Des, Carys', ''],
      ['Session 5 mentors', 'Phil, Cerith', ''],
      ['Session 6 mentors', 'David, Cerith', ''],
      ['Session 7 mentors', 'Cerith, David', ''],
      ['Session 8 mentors', 'Eleanor, David', ''],
      ['Session 9 mentors', 'John L, Rachel', ''],
      ['Session 10 mentors', 'Eleanor, Barrie H', '']
    ];
    settings.getRange(1, 1, rows.length, 3).setValues(rows);
    settings.getRange(1, 1, 1, 3).setFontWeight('bold');
    settings.setColumnWidth(1, 180); settings.setColumnWidth(2, 220); settings.setColumnWidth(3, 420);
  }

  if (members.getLastRow() <= 1) {   // the five members from the tracking sheet, Oct 2026
    [
      ['Luke E',  '2026-05-01', 'yes'],
      ['Suzzane', '2026-07-01', 'yes'],
      ['Anthony', '2026-07-01', ''],
      ['James',   '2026-10-15', ''],
      ['Jon G',   '2026-10-15', '']
    ].forEach(function(r){
      var items = {}; if (r[2]) items.MSP = { done: 'yes', mentor: '' };
      writeMember(members, { name: r[0], password: 'elli-' + rand4(), trade: '', inducted: r[1], items: items });
    });
  }
  SpreadsheetApp.flush();
}
function rand4(){ return String(Math.floor(1000 + Math.random() * 9000)); }

/* ---------------- Web app entry points ---------------- */
function doGet(e){
  var action = (e && e.parameter && e.parameter.action) || 'names';
  if (action === 'names') return out(publicInfo());
  return out({ ok: false, error: 'Unknown action' });
}
function doPost(e){
  var lock = LockService.getScriptLock();
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var a = body.action;
    if (a === 'names') return out(publicInfo());
    if (a === 'login') return out(login(body));
    lock.waitLock(10000);
    if (a === 'save')    return out(saveOwn(body));
    if (a === 'list')    return out(listAll(body));
    if (a === 'upsert')  return out(upsert(body));
    if (a === 'archive') return out(archiveMember(body));
    if (a === 'remove')  return out(removeMember(body));
    return out({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return out({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}
function out(obj){ return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

/* ---------------- Actions ---------------- */
function publicInfo(){
  return { ok: true, names: readAll(sheet(MEMBERS)).map(function(m){ return m.name; }), settings: publicSettings() };
}
function login(b){
  var cur = authMember(b.name, b.password);
  return { ok: true, member: strip(cur.member), settings: publicSettings() };
}
function saveOwn(b){
  var cur = authMember(b.name, b.password), m = cur.member;
  if (typeof b.trade === 'string') m.trade = b.trade;
  if (b.items && typeof b.items === 'object') m.items = cleanItems(b.items);
  writeMember(sheet(MEMBERS), m, cur.row);
  return { ok: true, member: strip(m), settings: publicSettings() };
}
function listAll(b){
  authCoord(b.password);
  return listPayload();
}
function upsert(b){
  authCoord(b.password);
  var inc = b.member || {};
  if (!inc.name || !String(inc.name).trim()) throw new Error('A name is needed');
  var sh = sheet(MEMBERS);
  var existingName = inc.originalName || inc.name;
  var cur = findRow(sh, existingName);
  var m = cur ? cur.member : { name: '', password: '', trade: '', inducted: '', items: {}, completed: '', notes: '' };
  if (!cur && findRow(sh, inc.name)) throw new Error('There is already a member called ' + inc.name);
  m.name = String(inc.name).trim();
  if (typeof inc.password === 'string') m.password = inc.password;
  if (typeof inc.trade === 'string') m.trade = inc.trade;
  if (typeof inc.inducted === 'string') m.inducted = normDate(inc.inducted);
  if (inc.items && typeof inc.items === 'object') m.items = cleanItems(inc.items);
  writeMember(sh, m, cur ? cur.row : null);
  return listPayload();
}
function archiveMember(b){
  authCoord(b.password);
  var sh = sheet(MEMBERS), cur = findRow(sh, b.name);
  if (!cur) throw new Error('No member called ' + b.name);
  cur.member.completed = normDate(b.completed) || today();
  writeMember(sheet(ARCHIVE), cur.member);
  sh.deleteRow(cur.row);
  return listPayload();
}
function removeMember(b){
  authCoord(b.password);
  var sh = sheet(b.archived ? ARCHIVE : MEMBERS), cur = findRow(sh, b.name);
  if (!cur) { sh = sheet(b.archived ? MEMBERS : ARCHIVE); cur = findRow(sh, b.name); }
  if (!cur) throw new Error('No member called ' + b.name);
  sh.deleteRow(cur.row);
  return listPayload();
}
function listPayload(){
  return { ok: true, members: readAll(sheet(MEMBERS)), archived: readAll(sheet(ARCHIVE)), settings: publicSettings() };
}

/* ---------------- Auth ---------------- */
function authMember(name, password){
  var cur = findRow(sheet(MEMBERS), name);
  if (!cur || !cur.member.password || String(cur.member.password) !== String(password || '')) throw new Error('Name or password not recognised');
  return cur;   // { row, member }
}
function authCoord(password){
  var want = settings().coordinatorPassword;
  if (!want) throw new Error('No coordinator password is set. Put one in the Settings tab of the sheet.');
  if (String(password || '') !== String(want)) throw new Error('Coordinator password not recognised');
}
function strip(m){ return { name: m.name, trade: m.trade, inducted: m.inducted, items: m.items, completed: m.completed }; }

/* ---------------- Settings tab ---------------- */
function publicSettings(){ var s = settings(); return { chapter: s.chapter, mentors: s.mentors, extraItems: s.extraItems }; }
function settings(){
  var sh = sheet(SETTINGS), vals = sh.getLastRow() ? sh.getRange(1, 1, sh.getLastRow(), 2).getValues() : [];
  var map = {}; vals.forEach(function(r){ if (r[0]) map[String(r[0]).trim().toLowerCase()] = String(r[1] == null ? '' : r[1]).trim(); });
  var mentors = {};
  SESSION_KEYS.forEach(function(k){ var v = map['session ' + k + ' mentors']; if (v) mentors[k] = v.split(',').map(function(x){ return x.trim(); }).filter(String); });
  var extra = (map['extra items'] || '').split(',').map(function(x){ return x.trim().toUpperCase(); }).filter(function(x){ return x === 'LP' || x === 'RT'; });
  return { chapter: map['chapter'] || '', coordinatorPassword: map['coordinator password'] || '', mentors: mentors, extraItems: extra };
}

/* ---------------- Rows ↔ member objects ---------------- */
function sheet(name){ var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); if (!sh) throw new Error('Run setup first: the ' + name + ' tab is missing'); return sh; }
function headersOf(sh){ var n = sh.getLastColumn(); return n ? sh.getRange(1, 1, 1, n).getValues()[0].map(function(h){ return String(h).trim(); }) : []; }
function ensureHeaders(sh, needed){
  var have = headersOf(sh), added = false;
  needed.forEach(function(h){ if (have.indexOf(h) < 0) { have.push(h); added = true; } });
  if (added) { var tailIdx = have.indexOf('Completed'); if (tailIdx > -1) { have = have.filter(function(h){ return TAIL.indexOf(h) < 0; }).concat(TAIL); } sh.getRange(1, 1, 1, have.length).setValues([have]).setFontWeight('bold'); }
  return have;
}
function rowToMember(headers, row){
  var m = { name: '', password: '', trade: '', inducted: '', items: {}, completed: '', notes: '' };
  headers.forEach(function(h, i){
    var v = row[i]; if (v instanceof Date) v = fmtDate(v); v = (v == null) ? '' : String(v).trim();
    if (h === 'Name') m.name = v; else if (h === 'Password') m.password = v; else if (h === 'Classification') m.trade = v;
    else if (h === 'Inducted') m.inducted = normDate(v); else if (h === 'Completed') m.completed = normDate(v); else if (h === 'Notes') m.notes = v;
    else if (/^Mentor /.test(h)) { var k = keyFromLabel(h.slice(7)); if (ITEM_KEYS.indexOf(k) > -1) { m.items[k] = m.items[k] || { done: '', mentor: '' }; m.items[k].mentor = v; } }
    else { var k2 = keyFromLabel(h); if (ITEM_KEYS.indexOf(k2) > -1) { m.items[k2] = m.items[k2] || { done: '', mentor: '' }; m.items[k2].done = normDone(v); } }
  });
  return m;
}
function readAll(sh){
  var n = sh.getLastRow(); if (n < 2) return [];
  var headers = headersOf(sh), rows = sh.getRange(2, 1, n - 1, headers.length).getValues();
  return rows.filter(function(r){ return String(r[0] || '').trim(); }).map(function(r){ return rowToMember(headers, r); });
}
function findRow(sh, name){
  var n = sh.getLastRow(); if (n < 2 || !name) return null;
  var headers = headersOf(sh), rows = sh.getRange(2, 1, n - 1, headers.length).getValues();
  var want = String(name).trim().toLowerCase();
  for (var i = 0; i < rows.length; i++) if (String(rows[i][0] || '').trim().toLowerCase() === want) return { row: i + 2, member: rowToMember(headers, rows[i]) };
  return null;
}
function writeMember(sh, m, rowIndex){
  var needed = [];
  Object.keys(m.items || {}).forEach(function(k){ if (ITEM_KEYS.indexOf(k) > -1) { needed.push(label(k)); needed.push('Mentor ' + label(k)); } });
  var headers = ensureHeaders(sh, needed);
  var row = headers.map(function(h){
    if (h === 'Name') return m.name; if (h === 'Password') return m.password || ''; if (h === 'Classification') return m.trade || '';
    if (h === 'Inducted') return m.inducted || ''; if (h === 'Completed') return m.completed || ''; if (h === 'Notes') return m.notes || '';
    if (/^Mentor /.test(h)) { var k = keyFromLabel(h.slice(7)); return (m.items && m.items[k] && m.items[k].mentor) || ''; }
    var k2 = keyFromLabel(h); return (m.items && m.items[k2] && m.items[k2].done) || '';
  });
  var r = rowIndex || sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, row.length).setNumberFormat('@').setValues([row]);
}
function cleanItems(items){
  var o = {};
  Object.keys(items).forEach(function(k){ if (ITEM_KEYS.indexOf(k) > -1 && items[k] && typeof items[k] === 'object') o[k] = { done: normDone(String(items[k].done || '')), mentor: String(items[k].mentor || '').slice(0, 80) }; });
  return o;
}

/* ---------------- Dates ---------------- */
function fmtDate(d){ return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function today(){ return fmtDate(new Date()); }
function normDate(v){
  if (!v) return ''; if (v instanceof Date) return fmtDate(v);
  var s = String(v).trim(), m;
  if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s))) return m[1] + '-' + pad(m[2]) + '-' + pad(m[3]);
  if ((m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(s))) { var y = m[3].length === 2 ? '20' + m[3] : m[3]; return y + '-' + pad(m[2]) + '-' + pad(m[1]); }   // UK d/m/y
  return s;
}
function normDone(v){
  var s = String(v || '').trim(); if (!s) return '';
  var d = normDate(s); if (/^\d{4}-\d{2}-\d{2}$/.test(d)) return d;
  return 'yes';   // "Yes", "Y", "x", a tick: done, date not recorded
}
function pad(n){ n = String(n); return n.length === 1 ? '0' + n : n; }
