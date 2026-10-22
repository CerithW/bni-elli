/**
 * Passport to Success – chapter store (v2)
 *
 * Lives inside a Google Sheet (Extensions → Apps Script) and lets the Passport web page
 * read and write that sheet. Tabs:
 *   Members   – one row per new member in the programme (name, password, dates, mentors, bookings)
 *   Archive   – members whose certificate has been presented
 *   Settings  – chapter name, passwords, mentors per session, weekly email
 *   Dashboard – live formulas over Members (week, target, stamped, status)
 *
 * Setup: paste this file in, run `setup` (approve the permissions it asks for), then
 * Deploy → New deployment → Web app, "Execute as: Me", "Who has access: Anyone", and put the
 * web app URL into CHAPTER.apiUrl in the Passport page. `setup` is safe to run again at any
 * time: it only adds what is missing (new columns, Settings rows, the Dashboard, the weekly
 * email timer). After changing this code, Deploy → Manage deployments → edit → New version.
 */

var MEMBERS = 'Members', ARCHIVE = 'Archive', SETTINGS = 'Settings', DASHBOARD = 'Dashboard';
var SESSION_KEYS = ['1','2','3','4','5','6','7','8','9','10'];
var ITEM_KEYS = ['1','2','3','4','LP','5','6','7','8','9','10','RT','MSP'];
var BASE = ['Name','Password','Classification','Inducted'];
var TAIL = ['Completed','Notes'];
var TARGET_DAYS = 84;
var SESSION_SHORT = { '1':'Roles & commitment', '2':'Policies, PALMS & Traffic Lights', '3':'Bio, speakers, fees & renewal', '4':'Attendance & substitutes',
  '5':'Referrals vs leads', '6':'121 etiquette', '7':'Connect & the app', '8':'Inviting & Gold Club', '9':'Visitor host day', '10':'Training',
  'LP':'Launch Pad', 'RT':'Regional sign-off', 'MSP':'MSP' };
var DEFAULT_PAGE = 'https://cerithw.github.io/bni-elli/passport/';

function label(k){ return /^\d+$/.test(k) ? 'S' + k : k; }          // '1' → 'S1', 'MSP' → 'MSP'
function keyFromLabel(h){ var m = /^S(\d+)$/.exec(h); return m ? m[1] : h; }
function doneHeaders(){ return ['S1','S2','S3','S4','S5','S6','S7','S8','S9','S10','MSP']; }
function mentorHeaders(){ return doneHeaders().map(function(h){ return 'Mentor ' + h; }); }
function bookedHeaders(){ return doneHeaders().map(function(h){ return 'Booked ' + h; }); }

/* ---------------- Setup (safe to re-run) ---------------- */
function setup(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var members = ss.getSheetByName(MEMBERS) || ss.insertSheet(MEMBERS);
  var archive = ss.getSheetByName(ARCHIVE) || ss.insertSheet(ARCHIVE);
  var setSh = ss.getSheetByName(SETTINGS) || ss.insertSheet(SETTINGS);
  var fresh = members.getLastRow() === 0;

  var headers = BASE.concat(doneHeaders()).concat(mentorHeaders()).concat(bookedHeaders()).concat(TAIL);
  [members, archive].forEach(function(sh){
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1); sh.setFrozenColumns(1);
      sh.getRange(2, 1, 500, headers.length).setNumberFormat('@');   // keep dates as plain text, e.g. 2026-10-22
    } else {
      ensureHeaders(sh, headers);
    }
  });

  // Settings: add any rows that are missing, keep what is there
  var isElli = /elli/i.test(ss.getName()) || (setSh.getLastRow() > 0 && /elli/i.test(String(setSh.getRange(2, 2).getValue())));
  var wanted = [
    ['Chapter', isElli ? 'BNI Elli' : ss.getName().replace(/passport to success/i, '').replace(/[–\-:]+/g, ' ').trim() || 'BNI chapter', 'Shown at the top of the page'],
    ['Coordinator password', 'coord-' + rand4(), 'Change this to whatever you like. Needed to see every member.'],
    ['Extra items', '', 'Leave blank for 10 sessions + MSP. Put LP and/or RT here to also track Launch Pad and the Regional Team sign-off.'],
    ['Session 1 mentors', isElli ? 'Matthew, Eleanor' : '', 'Usual mentor first, then the alternative'],
    ['Session 2 mentors', isElli ? 'Rhys D, Cerith' : '', ''],
    ['Session 3 mentors', isElli ? 'James R, David D' : '', ''],
    ['Session 4 mentors', isElli ? 'Des, Carys' : '', ''],
    ['Session 5 mentors', isElli ? 'Phil, Cerith' : '', ''],
    ['Session 6 mentors', isElli ? 'David, Cerith' : '', ''],
    ['Session 7 mentors', isElli ? 'Cerith, David' : '', ''],
    ['Session 8 mentors', isElli ? 'Eleanor, David' : '', ''],
    ['Session 9 mentors', isElli ? 'John L, Rachel' : '', ''],
    ['Session 10 mentors', isElli ? 'Eleanor, Barrie H' : '', ''],
    ['Page link', DEFAULT_PAGE, 'Used in the weekly email and welcome messages'],
    ['Weekly email', 'on', 'on / off. Sent every Wednesday at about 6pm with the summary and each mentor’s outstanding sessions.'],
    ['Weekly email to', ownerEmail(), 'Who receives the weekly email. Separate several addresses with commas.']
  ];
  if (setSh.getLastRow() === 0) {
    setSh.getRange(1, 1, 1, 3).setValues([['Setting', 'Value', 'Notes']]).setFontWeight('bold');
    setSh.setColumnWidth(1, 200); setSh.setColumnWidth(2, 240); setSh.setColumnWidth(3, 460);
  }
  var have = settingsMap();
  wanted.forEach(function(r){ if (!(r[0].toLowerCase() in have)) setSh.appendRow(r); });
  // one password per mentor named in the roster
  var s = settings();
  distinctMentors(s.mentors).forEach(function(name){
    var key = 'mentor password: ' + name.toLowerCase();
    if (!(key in settingsMap())) setSh.appendRow(['Mentor password: ' + name, 'm-' + rand4(), 'Lets ' + name + ' sign in on the Mentors tab to see and sign off their sessions']);
  });

  if (fresh && isElli) {   // the five members from the tracking sheet, Oct 2026
    [['Luke E','2026-05-01','yes'], ['Suzzane','2026-07-01','yes'], ['Anthony','2026-07-01',''], ['James','2026-10-15',''], ['Jon G','2026-10-15','']]
    .forEach(function(r){ var items = {}; if (r[2]) items.MSP = { done: 'yes', mentor: '', booked: '' };
      writeMember(members, { name: r[0], password: 'elli-' + rand4(), trade: '', inducted: r[1], items: items }); });
  }

  applyFormatting(members); applyFormatting(archive);
  buildDashboard(ss);
  ensureWeeklyTrigger();
  SpreadsheetApp.flush();
}
function rand4(){ return String(Math.floor(1000 + Math.random() * 9000)); }
function ownerEmail(){ try { return Session.getEffectiveUser().getEmail() || ''; } catch (e) { return ''; } }
function distinctMentors(mentors){
  var seen = {}, out = [];
  Object.keys(mentors || {}).forEach(function(k){ (mentors[k] || []).forEach(function(n){ var key = n.toLowerCase(); if (n && !seen[key]) { seen[key] = true; out.push(n); } }); });
  return out;
}

/* Green when a session is signed off, amber when it is booked */
function applyFormatting(sh){
  var headers = headersOf(sh); if (!headers.length) return;
  var rules = [];
  var doneCols = headers.map(function(h, i){ return doneHeaders().indexOf(h) > -1 || h === 'LP' || h === 'RT' ? i + 1 : 0; }).filter(Boolean);
  var bookedCols = headers.map(function(h, i){ return /^Booked /.test(h) ? i + 1 : 0; }).filter(Boolean);
  if (doneCols.length) rules.push(SpreadsheetApp.newConditionalFormatRule().whenCellNotEmpty().setBackground('#C8E6C9')
      .setRanges(doneCols.map(function(c){ return sh.getRange(2, c, 499, 1); })).build());
  if (bookedCols.length) rules.push(SpreadsheetApp.newConditionalFormatRule().whenCellNotEmpty().setBackground('#FFE8B8')
      .setRanges(bookedCols.map(function(c){ return sh.getRange(2, c, 499, 1); })).build());
  sh.setConditionalFormatRules(rules);
}

/* Dashboard tab: live formulas over Members, so the sheet reads well without the page */
function buildDashboard(ss){
  var dash = ss.getSheetByName(DASHBOARD) || ss.insertSheet(DASHBOARD);
  var h = headersOf(sheet(MEMBERS));
  var L = function(name){ return colLetter(h.indexOf(name) + 1); };
  var ind = L('Inducted'), s1 = L('S1'), s10 = L('S10'), msp = L('MSP');
  var M = "'" + MEMBERS + "'!";
  var notBlank = "IF(" + M + "A2:A=\"\",\"\",";
  var rows = [
    ['Name', 'Inducted', 'Days', 'Week', 'Target', 'Sessions done (of 10)', 'Stamped (of 11)', 'Status'],
    ["=ARRAYFORMULA(" + notBlank + M + "A2:A))",
     "=ARRAYFORMULA(" + notBlank + M + ind + "2:" + ind + "))",
     "=ARRAYFORMULA(" + notBlank + "IFERROR(TODAY()-DATEVALUE(" + M + ind + "2:" + ind + "),\"\")))",
     "=ARRAYFORMULA(IF(C2:C=\"\",\"\",INT(C2:C/7)+1))",
     "=ARRAYFORMULA(" + notBlank + "IFERROR(TEXT(DATEVALUE(" + M + ind + "2:" + ind + ")+" + TARGET_DAYS + ",\"dd mmm yyyy\"),\"\")))",
     "=ARRAYFORMULA(" + notBlank + "MMULT(--(" + M + s1 + "2:" + s10 + "<>\"\"),TRANSPOSE(COLUMN(" + M + s1 + "2:" + s10 + "2)^0))))",
     "=ARRAYFORMULA(" + notBlank + "MMULT(--(" + M + s1 + "2:" + msp + "<>\"\"),TRANSPOSE(COLUMN(" + M + s1 + "2:" + msp + "2)^0))))",
     "=ARRAYFORMULA(IF(A2:A=\"\",\"\",IF(G2:G>=11,\"Complete\",IF(C2:C=\"\",\"No induction date\",IF(C2:C>" + TARGET_DAYS + ",\"Overdue\",IF(F2:F<IF(INT(C2:C/7)>10,10,INT(C2:C/7))-1,\"Behind pace\",\"On track\"))))))"]
  ];
  dash.getRange(1, 1, 2, rows[0].length).setValues(rows);
  dash.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  dash.setFrozenRows(1);
  var status = dash.getRange(2, 8, 499, 1);
  dash.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('On track').setBackground('#C8E6C9').setRanges([status]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Behind pace').setBackground('#FFE8B8').setRanges([status]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Overdue').setBackground('#F8C9CC').setRanges([status]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Complete').setBackground('#E0E0E0').setRanges([status]).build()
  ]);
}
function colLetter(n){ var s = ''; while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }

/* ---------------- Weekly email ---------------- */
function ensureWeeklyTrigger(){
  var exists = ScriptApp.getProjectTriggers().some(function(t){ return t.getHandlerFunction() === 'weeklyEmail'; });
  if (!exists) ScriptApp.newTrigger('weeklyEmail').timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(18).create();
}
function weeklyEmail(){
  var s = settings();
  if ((s.weeklyEmail || 'on').toLowerCase() === 'off') return 'off';
  var to = s.weeklyEmailTo || ownerEmail(); if (!to) return 'no address';
  var body = summaryText(s);
  MailApp.sendEmail({ to: to, subject: 'Passport to Success – ' + s.chapter + ' – ' + fmtLong(today()), body: body });
  return to;
}
function summaryText(s){
  s = s || settings();
  var members = readAll(sheet(MEMBERS)), extra = s.extraItems || [];
  var order = orderFor(extra);
  var L = [];
  L.push('Passport to Success – ' + s.chapter + ' – ' + fmtLong(today()));
  L.push('');
  L.push('NEW MEMBERS');
  if (!members.length) L.push('Nobody in the programme at the moment.');
  members.sort(function(a, b){ return (a.inducted || '9') < (b.inducted || '9') ? -1 : 1; }).forEach(function(m){
    var p = progressOf(m, order, s.mentors);
    var line = '• ' + m.name + (m.trade ? ' (' + m.trade + ')' : '') + ' – ';
    if (p.status === 'Complete') line += 'complete, ' + p.done + '/' + p.total + '. Certificate due.';
    else {
      line += (p.week ? 'week ' + p.week + ', ' : '') + p.done + '/' + p.total + ' stamped, ' + p.status.toLowerCase();
      if (p.next) {
        var it = m.items[p.next] || {};
        line += '. Next: S' + p.next + ' ' + SESSION_SHORT[p.next];
        if (it.booked) line += ' – booked ' + fmtLong(it.booked) + (it.mentor ? ' with ' + it.mentor : '');
        else line += ' – not booked' + (it.mentor ? ' (' + it.mentor + ')' : (s.mentors[p.next] ? ' (' + s.mentors[p.next].join(' or ') + ')' : ''));
      }
    }
    L.push(line);
  });
  L.push('');
  L.push('MENTORS – OUTSTANDING SESSIONS');
  var chase = chaseList(members, s.mentors, order);
  var names = Object.keys(chase).sort();
  if (!names.length) L.push('Nothing outstanding.');
  names.forEach(function(n){ L.push(n + ': ' + chase[n].join('; ')); });
  L.push('');
  L.push('Open the tracker: ' + (s.pageLink || DEFAULT_PAGE) + '#tracker');
  return L.join('\n');
}
/* Who owes what this week: each member's next session, plus anything already booked. The paired mentor
   if one is set, otherwise the usual mentor for that session. */
function chaseList(members, mentors, order){
  var out = {};
  members.forEach(function(m){
    var p = progressOf(m, order, mentors);
    SESSION_KEYS.forEach(function(k){
      var it = m.items[k] || {}; if (it.done) return;
      if (k !== p.next && !it.booked) return;
      var who = it.mentor || (mentors[k] && mentors[k][0]) || ''; if (!who) return;
      (out[who] = out[who] || []).push('S' + k + ' for ' + m.name + (it.booked ? ' (booked ' + fmtShort(it.booked) + ')' : ' (not booked)'));
    });
  });
  return out;
}
function orderFor(extra){ var x = extra || []; return ['1','2','3','4'].concat(x.indexOf('LP') > -1 ? ['LP'] : []).concat(['5','6','7','8','9','10']).concat(x.indexOf('RT') > -1 ? ['RT'] : []).concat(['MSP']); }
function progressOf(m, order, mentors){
  var items = m.items || {}, asOf = today();
  var done = order.filter(function(k){ return items[k] && items[k].done; }).length;
  var sessionsDone = SESSION_KEYS.filter(function(k){ return items[k] && items[k].done; }).length;
  var days = m.inducted ? daysBetween(m.inducted, asOf) : null;
  var week = days == null ? null : Math.floor(Math.max(0, days) / 7) + 1;
  var rec = ['7'].concat(order.filter(function(k){ return k !== '7'; }));
  var next = null; for (var i = 0; i < rec.length; i++) { var k = rec[i]; if (SESSION_KEYS.indexOf(k) > -1 && !(items[k] && items[k].done)) { next = k; break; } }
  var status;
  if (done === order.length) status = 'Complete';
  else if (days == null) status = 'No induction date';
  else if (days > TARGET_DAYS) status = 'Overdue';
  else { var expected = Math.min(10, Math.floor(days / 7)); status = sessionsDone < expected - 1 ? 'Behind pace' : 'On track'; }
  return { done: done, total: order.length, sessionsDone: sessionsDone, days: days, week: week, next: next, status: status };
}

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
    if (a === 'mentorLogin') return out(mentorLogin(body));
    lock.waitLock(10000);
    if (a === 'save')       return out(saveOwn(body));
    if (a === 'mentorSign') return out(mentorSign(body));
    if (a === 'list')       return out(listAll(body));
    if (a === 'upsert')     return out(upsert(body));
    if (a === 'archive')    return out(archiveMember(body));
    if (a === 'remove')     return out(removeMember(body));
    if (a === 'emailNow')   return out(emailNow(body));
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
  var s = publicSettings();
  return { ok: true, names: readAll(sheet(MEMBERS)).map(function(m){ return m.name; }), mentorNames: distinctMentors(s.mentors), settings: s };
}
function login(b){
  var cur = authMember(b.name, b.password);
  return { ok: true, member: strip(cur.member), settings: publicSettings() };
}
function saveOwn(b){
  var cur = authMember(b.name, b.password), m = cur.member;
  if (typeof b.trade === 'string') m.trade = b.trade;
  if (b.items && typeof b.items === 'object') m.items = Object.assign(m.items || {}, cleanItems(b.items));
  writeMember(sheet(MEMBERS), m, cur.row);
  return { ok: true, member: strip(m), settings: publicSettings() };
}
function mentorLogin(b){
  var name = authMentor(b.name, b.password);
  return { ok: true, mentor: name, members: readAll(sheet(MEMBERS)).map(strip), settings: publicSettings() };
}
function mentorSign(b){
  var name = authMentor(b.name, b.password);
  var sh = sheet(MEMBERS), cur = findRow(sh, b.member);
  if (!cur) throw new Error('No member called ' + b.member);
  if (ITEM_KEYS.indexOf(String(b.item)) < 0) throw new Error('Unknown session');
  var it = cur.member.items[b.item] || { done: '', mentor: '', booked: '' };
  if ('done' in b) it.done = normDone(String(b.done || ''));
  if ('booked' in b) it.booked = normDate(String(b.booked || ''));
  if (it.done || it.booked) it.mentor = name;
  cur.member.items[b.item] = it;
  writeMember(sh, cur.member, cur.row);
  return { ok: true, mentor: name, members: readAll(sh).map(strip), settings: publicSettings() };
}
function listAll(b){ authCoord(b.password); return listPayload(); }
function upsert(b){
  authCoord(b.password);
  var inc = b.member || {};
  if (!inc.name || !String(inc.name).trim()) throw new Error('A name is needed');
  var sh = sheet(MEMBERS);
  var cur = findRow(sh, inc.originalName || inc.name);
  var m = cur ? cur.member : { name: '', password: '', trade: '', inducted: '', items: {}, completed: '', notes: '' };
  if (!cur && findRow(sh, inc.name)) throw new Error('There is already a member called ' + inc.name);
  m.name = String(inc.name).trim();
  if (typeof inc.password === 'string') m.password = inc.password;
  if (typeof inc.trade === 'string') m.trade = inc.trade;
  if (typeof inc.inducted === 'string') m.inducted = normDate(inc.inducted);
  if (inc.items && typeof inc.items === 'object') m.items = Object.assign(m.items || {}, cleanItems(inc.items));
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
function emailNow(b){
  authCoord(b.password);
  var s = settings(); var to = s.weeklyEmailTo || ownerEmail();
  if (!to) throw new Error('No address in the Settings tab (Weekly email to)');
  MailApp.sendEmail({ to: to, subject: 'Passport to Success – ' + s.chapter + ' – ' + fmtLong(today()), body: summaryText(s) });
  return { ok: true, sentTo: to };
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
function authMentor(name, password){
  var rec = settings().mentorPasswords[String(name || '').trim().toLowerCase()];
  if (!rec || String(rec.password) !== String(password || '')) throw new Error('Name or password not recognised');
  return rec.name;
}
function authCoord(password){
  var want = settings().coordinatorPassword;
  if (!want) throw new Error('No coordinator password is set. Put one in the Settings tab of the sheet.');
  if (String(password || '') !== String(want)) throw new Error('Coordinator password not recognised');
}
function strip(m){ return { name: m.name, trade: m.trade, inducted: m.inducted, items: m.items, completed: m.completed }; }

/* ---------------- Settings tab ---------------- */
function settingsMap(){
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS); if (!sh || !sh.getLastRow()) return {};
  var vals = sh.getRange(1, 1, sh.getLastRow(), 2).getValues(), map = {};
  vals.forEach(function(r){ if (r[0]) map[String(r[0]).trim().toLowerCase()] = String(r[1] == null ? '' : r[1]).trim(); });
  return map;
}
function publicSettings(){ var s = settings(); return { chapter: s.chapter, mentors: s.mentors, extraItems: s.extraItems, pageLink: s.pageLink }; }
function settings(){
  var map = settingsMap(), mentors = {}, mentorPasswords = {};
  SESSION_KEYS.forEach(function(k){ var v = map['session ' + k + ' mentors']; if (v) mentors[k] = v.split(',').map(function(x){ return x.trim(); }).filter(String); });
  Object.keys(map).forEach(function(key){ var m = /^mentor password:\s*(.+)$/.exec(key); if (m && map[key]) mentorPasswords[m[1].trim()] = { password: map[key], name: titleFromKey(m[1], mentors) }; });
  var extra = (map['extra items'] || '').split(',').map(function(x){ return x.trim().toUpperCase(); }).filter(function(x){ return x === 'LP' || x === 'RT'; });
  return { chapter: map['chapter'] || '', coordinatorPassword: map['coordinator password'] || '', mentors: mentors, mentorPasswords: mentorPasswords, extraItems: extra,
           pageLink: map['page link'] || DEFAULT_PAGE, weeklyEmail: map['weekly email'] || 'on', weeklyEmailTo: map['weekly email to'] || '' };
}
/* The Settings key is lower-cased; recover the mentor's name as written in the roster */
function titleFromKey(lowerName, mentors){
  var hit = distinctMentors(mentors).filter(function(n){ return n.toLowerCase() === lowerName.trim(); })[0];
  return hit || lowerName.trim().replace(/\b\w/g, function(c){ return c.toUpperCase(); });
}

/* ---------------- Rows ↔ member objects ---------------- */
function sheet(name){ var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); if (!sh) throw new Error('Run setup first: the ' + name + ' tab is missing'); return sh; }
function headersOf(sh){ var n = sh.getLastColumn(); return n ? sh.getRange(1, 1, 1, n).getValues()[0].map(function(h){ return String(h).trim(); }) : []; }
function ensureHeaders(sh, needed){
  var have = headersOf(sh), added = false;
  needed.forEach(function(h){ if (have.indexOf(h) < 0) { have.push(h); added = true; } });
  if (added) {
    have = have.filter(function(h){ return TAIL.indexOf(h) < 0; }).concat(TAIL);   // keep Completed / Notes at the end
    sh.getRange(1, 1, 1, have.length).setValues([have]).setFontWeight('bold');
    sh.getRange(2, 1, 500, have.length).setNumberFormat('@');
  }
  return have;
}
function rowToMember(headers, row){
  var m = { name: '', password: '', trade: '', inducted: '', items: {}, completed: '', notes: '' };
  var item = function(k){ return m.items[k] = m.items[k] || { done: '', mentor: '', booked: '' }; };
  headers.forEach(function(h, i){
    var v = row[i]; if (v instanceof Date) v = fmtDate(v); v = (v == null) ? '' : String(v).trim();
    if (h === 'Name') m.name = v; else if (h === 'Password') m.password = v; else if (h === 'Classification') m.trade = v;
    else if (h === 'Inducted') m.inducted = normDate(v); else if (h === 'Completed') m.completed = normDate(v); else if (h === 'Notes') m.notes = v;
    else if (/^Mentor /.test(h)) { var k = keyFromLabel(h.slice(7)); if (ITEM_KEYS.indexOf(k) > -1) item(k).mentor = v; }
    else if (/^Booked /.test(h)) { var k3 = keyFromLabel(h.slice(7)); if (ITEM_KEYS.indexOf(k3) > -1) item(k3).booked = normDate(v); }
    else { var k2 = keyFromLabel(h); if (ITEM_KEYS.indexOf(k2) > -1) item(k2).done = normDone(v); }
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
  Object.keys(m.items || {}).forEach(function(k){ if (ITEM_KEYS.indexOf(k) > -1) { needed.push(label(k)); needed.push('Mentor ' + label(k)); needed.push('Booked ' + label(k)); } });
  var headers = ensureHeaders(sh, needed);
  var row = headers.map(function(h){
    if (h === 'Name') return m.name; if (h === 'Password') return m.password || ''; if (h === 'Classification') return m.trade || '';
    if (h === 'Inducted') return m.inducted || ''; if (h === 'Completed') return m.completed || ''; if (h === 'Notes') return m.notes || '';
    if (/^Mentor /.test(h)) { var k = keyFromLabel(h.slice(7)); return (m.items && m.items[k] && m.items[k].mentor) || ''; }
    if (/^Booked /.test(h)) { var k3 = keyFromLabel(h.slice(7)); return (m.items && m.items[k3] && m.items[k3].booked) || ''; }
    var k2 = keyFromLabel(h); return (m.items && m.items[k2] && m.items[k2].done) || '';
  });
  var r = rowIndex || sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, row.length).setNumberFormat('@').setValues([row]);
}
function cleanItems(items){
  var o = {};
  Object.keys(items).forEach(function(k){ if (ITEM_KEYS.indexOf(k) > -1 && items[k] && typeof items[k] === 'object')
    o[k] = { done: normDone(String(items[k].done || '')), mentor: String(items[k].mentor || '').slice(0, 80), booked: normDate(String(items[k].booked || '')) }; });
  return o;
}

/* ---------------- Dates ---------------- */
var MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
function fmtDate(d){ return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function today(){ return fmtDate(new Date()); }
function parseISO(s){ var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || '')); return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null; }
function fmtLong(s){ var d = parseISO(s); return d ? d.getDate() + ' ' + MONTHS[d.getMonth()] + ' ' + d.getFullYear() : String(s || ''); }
function fmtShort(s){ var d = parseISO(s); return d ? d.getDate() + ' ' + MONTHS[d.getMonth()] : String(s || ''); }
function daysBetween(a, b){ var da = parseISO(a), db = parseISO(b); return da && db ? Math.round((db - da) / 86400000) : null; }
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
