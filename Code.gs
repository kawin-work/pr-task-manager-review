/**
 * PR Task Manager · Review comments API
 * Google Sheet = shared store for the comment layer on the GitHub Pages prototype.
 * GET  ?action=list                         -> {ok, comments, replies, now}
 * POST {action:'add'|'update'|'remove', sheet:'comments'|'replies', id?, data?}  (Content-Type: text/plain)
 */
var ST = {open: 'รอตัดสิน', ready: 'พร้อมพัฒนา', done: 'เสร็จแล้ว'};
var CC = [
  ['id', 'รหัส'], ['createdAt', 'สร้างเมื่อ'], ['authorName', 'ผู้เขียน'], ['screen', 'หน้าจอ'], ['pathText', 'ตำแหน่ง'],
  ['note', 'โน้ตต้นฉบับ'], ['statusTh', 'สถานะ'], ['title', 'หัวข้อ (สำหรับ dev)'], ['change', 'เปลี่ยน'], ['reason', 'เหตุผล'],
  ['affects', 'กระทบส่วนอื่น'], ['doneWhen', 'ถือว่าเสร็จเมื่อ'], ['questions', 'รอตัดสิน (คำถาม)'], ['role', 'มุมมอง'],
  ['updatedAt', 'แก้ไขล่าสุด'], ['selector', 'selector'], ['tag', 'tag'], ['elText', 'elText'], ['ox', 'ox'], ['oy', 'oy'],
  ['html', 'html'], ['author', 'authorId'], ['expandTh', 'ให้ Claude ขยาย']
];
var EX = {want: 'รอขยาย', no: 'ไม่ต้องขยาย', done: 'ขยายแล้ว'};
var RC = [['id', 'รหัส'], ['commentId', 'คอมเมนต์'], ['createdAt', 'เวลา'], ['authorName', 'ผู้เขียน'], ['text', 'ข้อความ'], ['author', 'authorId']];
var NOT_TEXT = {createdAt: 1, updatedAt: 1, ox: 1, oy: 1};
var WIDE = {note: 280, change: 320, reason: 240, affects: 240, doneWhen: 240, questions: 260, title: 260, pathText: 240, text: 360};

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name, cols) {
  var sh = ss_().getSheetByName(name);
  if (!sh) sh = ss_().insertSheet(name);
  var have = sh.getLastRow() === 0 ? 0 : sh.getLastColumn();
  if (have < cols.length) {
    // new sheet, or columns added in a later version: write the missing headers and format those columns
    if (sh.getMaxColumns() < cols.length) sh.insertColumnsAfter(sh.getMaxColumns(), cols.length - sh.getMaxColumns());
    sh.getRange(1, have + 1, 1, cols.length - have).setValues([cols.slice(have).map(function (c) { return c[1]; })]).setFontWeight('bold').setBackground('#EEEAFE');
    sh.setFrozenRows(1);
    var rows = sh.getMaxRows() - 1;
    cols.forEach(function (c, i) {
      if (i < have) return;
      var rg = sh.getRange(2, i + 1, rows, 1);
      if (!NOT_TEXT[c[0]]) rg.setNumberFormat('@');
      if (WIDE[c[0]]) { sh.setColumnWidth(i + 1, WIDE[c[0]]); rg.setWrap(true); }
      if (c[0] === 'createdAt' || c[0] === 'updatedAt') rg.setNumberFormat('d/m/yyyy hh:mm');
      if (c[0] === 'statusTh') rg.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([ST.open, ST.ready, ST.done], true).build());
      if (c[0] === 'expandTh') { rg.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([EX.want, EX.no, EX.done], true).setAllowInvalid(true).build()); sh.setColumnWidth(i + 1, 130); }
    });
  }
  return sh;
}
function setup() {
  sheet_('comments', CC); sheet_('replies', RC);
  ['Sheet1', 'ชีต1', 'แผ่นงาน1'].forEach(function (n) { var d = ss_().getSheetByName(n); if (d && ss_().getSheets().length > 1 && d.getLastRow() === 0) ss_().deleteSheet(d); });
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function toMs_(v) { return v instanceof Date ? v.getTime() : (Number(v) || 0); }
function lines_(v) { return String(v || '').split('\n').map(function (s) { return s.trim(); }).filter(String); }
function clean_(v) { if (v === undefined || v === null) return ''; if (typeof v === 'string' && v.charAt(0) === '=') return ' ' + v; return v; }
function rows_(name, cols) {
  var sh = sheet_(name, cols), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, cols.length).getValues().filter(function (r) { return r[0] !== ''; }).map(function (r) {
    var o = {}; cols.forEach(function (c, i) { o[c[0]] = r[i]; }); return o;
  });
}
function find_(sh, id) {
  var n = sh.getLastRow() - 1; if (n < 1) return -1;
  var ids = sh.getRange(2, 1, n, 1).getValues();
  for (var i = 0; i < n; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function rowObj_(sh, cols, r) { var v = sh.getRange(r, 1, 1, cols.length).getValues()[0], o = {}; cols.forEach(function (c, i) { o[c[0]] = v[i]; }); return o; }
function writeRow_(sh, cols, r, o) { if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 200); sh.getRange(r, 1, 1, cols.length).setValues([cols.map(function (c) { return clean_(o[c[0]]); })]); }

function commentOut_(o) {
  var hasSpec = String(o.title || '') !== '' || String(o.change || '') !== '';
  var st = 'open'; Object.keys(ST).forEach(function (k) { if (ST[k] === String(o.statusTh)) st = k; });
  var ex = ''; Object.keys(EX).forEach(function (k) { if (EX[k] === String(o.expandTh || '')) ex = k; });
  if (hasSpec) ex = 'done';
  return {
    id: String(o.id), createdAt: toMs_(o.createdAt), updatedAt: toMs_(o.updatedAt), authorName: String(o.authorName), author: String(o.author),
    screen: String(o.screen), role: String(o.role), path: String(o.pathText).split(' › ').filter(String), note: String(o.note), status: st,
    spec: hasSpec ? {title: String(o.title), change: String(o.change), reason: String(o.reason), affects: lines_(o.affects), done_when: lines_(o.doneWhen), questions: lines_(o.questions)} : null,
    selector: String(o.selector), tag: String(o.tag), elText: String(o.elText), ox: Number(o.ox) || 0, oy: Number(o.oy) || 0, html: String(o.html),
    expand: ex
  };
}
function commentIn_(d, prev) {
  var o = prev || {};
  function set(k, v) { if (v !== undefined) o[k] = v; }
  if (d.createdAt) o.createdAt = new Date(d.createdAt);
  o.updatedAt = d.updatedAt ? new Date(d.updatedAt) : new Date();
  set('authorName', d.authorName); set('author', d.author); set('screen', d.screen); set('role', d.role); set('note', d.note);
  if (d.path) o.pathText = d.path.join(' › ');
  if (d.status) o.statusTh = ST[d.status] || ST.open;
  if (d.spec !== undefined) {
    var s = d.spec || {};
    o.title = s.title || ''; o.change = s.change || ''; o.reason = s.reason || '';
    o.affects = (s.affects || []).join('\n'); o.doneWhen = (s.done_when || []).join('\n'); o.questions = (s.questions || []).join('\n');
  }
  ['selector', 'tag', 'elText', 'ox', 'oy', 'html'].forEach(function (k) { set(k, d[k]); });
  if (d.expand !== undefined) o.expandTh = EX[d.expand] || '';
  if (d.spec && (d.spec.title || d.spec.change)) o.expandTh = EX.done;
  if (!o.statusTh) o.statusTh = ST.open;
  return o;
}
function replyOut_(o) { return {id: String(o.id), commentId: String(o.commentId), createdAt: toMs_(o.createdAt), authorName: String(o.authorName), text: String(o.text), author: String(o.author)}; }
function replyIn_(d, prev) {
  var o = prev || {};
  ['commentId', 'authorName', 'text', 'author'].forEach(function (k) { if (d[k] !== undefined) o[k] = d[k]; });
  if (d.createdAt || !o.createdAt) o.createdAt = new Date(d.createdAt || Date.now());
  return o;
}

function doGet() {
  return out_({ok: true, comments: rows_('comments', CC).map(commentOut_), replies: rows_('replies', RC).map(replyOut_), now: Date.now()});
}
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var b = JSON.parse(e.postData.contents);
    var name = b.sheet === 'replies' ? 'replies' : 'comments', cols = name === 'replies' ? RC : CC, sh = sheet_(name, cols);
    var isR = name === 'replies';
    if (b.action === 'add') {
      var id = (isR ? 'r' : 'c') + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
      var o = isR ? replyIn_(b.data || {}) : commentIn_(b.data || {});
      o.id = id; if (!o.createdAt) o.createdAt = new Date();
      writeRow_(sh, cols, sh.getLastRow() + 1, o);
      return out_({ok: true, id: id});
    }
    var r = find_(sh, b.id);
    if (b.action === 'update') {
      if (r < 0) return out_({ok: false, error: 'not_found'});
      var prev = rowObj_(sh, cols, r);
      writeRow_(sh, cols, r, isR ? replyIn_(b.data || {}, prev) : commentIn_(b.data || {}, prev));
      return out_({ok: true});
    }
    if (b.action === 'remove') {
      if (r > 0) sh.deleteRow(r);
      if (!isR) {
        var rs = ss_().getSheetByName('replies');
        if (rs) for (var i = rs.getLastRow(); i >= 2; i--) if (String(rs.getRange(i, 2).getValue()) === String(b.id)) rs.deleteRow(i);
      }
      return out_({ok: true});
    }
    return out_({ok: false, error: 'bad_action'});
  } catch (err) {
    return out_({ok: false, error: String(err)});
  } finally {
    lock.releaseLock();
  }
}
