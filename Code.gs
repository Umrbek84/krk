/**
 * KRK / KRS — oʻqituvchilarning kasbiy rivojlanish kuni test tizimi (Code.gs)
 *
 * Varaqlar (Google Sheets):
 *   PIN_Codes   A2 = PIN-kod, B2 = "Ochiq" / "Yopiq"
 *   Testlar     A: Test (varaq nomi) | B: Holat (Ochiq/Yopiq) | C: Vaqt (daqiqa) | D: Oʻtish bali (%) | E: Izoh
 *   <Test nomi> har bir test alohida varaq. Ustunlar:
 *               A №  | B Savol | C A | D B | E C | F D | G Toʻgʻri javob (matn yoki A/B/C/D) | H Ball (ixtiyoriy, boʻsh = 1)
 *   Natijalar   natijalar jadvali (avtomatik toʻldiriladi)
 *   Javoblar    har bir urinishning batafsil javoblari (admin panel uchun)
 *
 * Ishga tushirish: setup() funksiyasini bir marta ishga tushiring, soʻng
 * Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 */

const SPREADSHEET_ID = '';   // Jadvalga bogʻlangan (Extensions → Apps Script) skriptda boʻsh qoldiring
const ADMIN_PIN = 'PIN-NI-SHU-YERGA-YOZING';    // Haqiqiy PIN faqat Apps Script ichida saqlanadi, repoga yozilmaydi
const SERVICE = ['Testlar', 'PIN_Codes', 'Natijalar', 'Javoblar'];
const SESSION_HOURS = 6;

function ss_() { return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function t_(v) { return v === null || v === undefined ? '' : String(v).trim(); }

function doGet(e) { return json_(route_((e && e.parameter) || {})); }
function doPost(e) {
  var p = {};
  try { p = JSON.parse(e.postData.contents); } catch (x) {}
  return json_(route_(p));
}

function route_(p) {
  try {
    switch (p.action) {
      case 'tests': return { success: true, data: listTests_(true) };
      case 'start': return start_(p);
      case 'submit': return submit_(p);
      case 'adminList': return adminList_(p);
      case 'adminAttempt': return adminAttempt_(p);
      case 'adminDelete': return adminDelete_(p);
      default: return { success: true, msg: 'KRK API ishlayapti' };
    }
  } catch (err) {
    return { success: false, msg: err.message };
  }
}

// ---------------------------------------------------------------- testlar roʻyxati
function listTests_(withCount) {
  var ss = ss_(), list = [];
  var sh = ss.getSheetByName('Testlar');
  var rows = (sh && sh.getLastRow() > 1) ? sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues() : [];
  var cfg = [];
  rows.forEach(function (r) {
    var n = t_(r[0]);
    if (n) cfg.push({ name: n, open: t_(r[1]).toLowerCase() !== 'yopiq', minutes: Number(r[2]) || 30, pass: Number(r[3]) || 0, note: t_(r[4]) });
  });
  if (!cfg.length) {
    ss.getSheets().forEach(function (s) {
      var n = s.getName();
      if (SERVICE.indexOf(n) < 0 && n.charAt(0) !== '_') cfg.push({ name: n, open: true, minutes: 30, pass: 0, note: '' });
    });
  }
  cfg.forEach(function (c) {
    var s = ss.getSheetByName(c.name);
    if (!s) return;
    if (withCount) {
      var cnt = 0;
      if (s.getLastRow() > 1) cnt = s.getRange(2, 2, s.getLastRow() - 1, 1).getValues().filter(function (r) { return t_(r[0]) !== ''; }).length;
      c.count = cnt;
    }
    list.push(c);
  });
  return list;
}

function checkPin_(pin) {
  var sh = ss_().getSheetByName('PIN_Codes');
  if (!sh) throw new Error('PIN_Codes varagʻi topilmadi!');
  var d = sh.getRange('A2:B2').getValues()[0];
  if (t_(d[0]) !== t_(pin) || t_(d[1]).toLowerCase() !== 'ochiq') throw new Error('PIN-kod xato yoki test hozircha yopiq!');
}

// ---------------------------------------------------------------- savollar
function norm_(s) { return String(s).toLowerCase().replace(/[\s"'ʻʼ‘’`´.,;:!?()\-]/g, ''); }

function findCorrect_(c, opt) {
  var nc = norm_(c);
  if (!nc) return -1;
  var i;
  for (i = 0; i < 4; i++) if (opt[i] && norm_(opt[i]) === nc) return i;           // 1) variant matni
  var lc = c.toLowerCase().replace(/[\s).]/g, '');
  var LET = { a: 0, b: 1, c: 2, d: 3, 'а': 0, 'в': 1, 'с': 2, 'д': 3 };
  if (lc.length === 1 && LET[lc] !== undefined) return LET[lc];                      // 2) A/B/C/D harfi
  if (/^[1-4]$/.test(lc)) return parseInt(lc, 10) - 1;                               // 3) 1-4 raqami
  return -1;
}

function loadQuestions_(name, problems) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error("'" + name + "' varagʻi topilmadi!");
  var data = sh.getDataRange().getValues(), out = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i], q = t_(r[1]);
    if (!q) continue;
    var opt = [r[2], r[3], r[4], r[5]].map(t_);
    var corr = findCorrect_(t_(r[6]), opt);
    if (corr < 0 || !opt[corr]) {
      if (problems) problems.push(name + ': ' + (i + 1) + '-qator — toʻgʻri javob variantlardan topilmadi');
      continue;
    }
    var pts = parseFloat(String(r[7]).replace(',', '.'));
    if (!(pts > 0)) pts = 1;
    out.push({ id: i, q: q, opt: opt, correct: corr, pts: pts, noShuffle: /barcha|hammasi|hech biri|все|всё/i.test(opt.join(' | ')) });
  }
  return out;
}

// ---------------------------------------------------------------- test boshlash
function start_(p) {
  checkPin_(p.pin);
  var name = t_(p.name);
  if (name.split(/\s+/).length < 2) throw new Error('F.I.Sh. ni toʻliq kiriting!');
  var cfg = null;
  listTests_(false).forEach(function (t) { if (t.name === p.test) cfg = t; });
  if (!cfg) throw new Error('Test topilmadi!');
  if (!cfg.open) throw new Error('Bu test hozircha yopiq.');
  var qs = loadQuestions_(cfg.name);
  if (!qs.length) throw new Error("'" + cfg.name + "' varagʻida savol yoʻq!");

  var token = Utilities.getUuid();
  var info = { test: cfg.name, name: name, region: t_(p.region), org: t_(p.org), start: Date.now(), minutes: cfg.minutes, pass: cfg.pass };
  CacheService.getScriptCache().put('s_' + token, JSON.stringify(info), SESSION_HOURS * 3600);

  // Toʻgʻri javob mijozga yuborilmaydi — tekshirish serverda
  return {
    success: true, token: token, minutes: cfg.minutes, note: cfg.note,
    questions: qs.map(function (q) { return { id: q.id, q: q.q, opt: q.opt, pts: q.pts, noShuffle: q.noShuffle }; })
  };
}

// ---------------------------------------------------------------- yakunlash
function ensureSheets_() {
  var ss = ss_();
  var n = ss.getSheetByName('Natijalar');
  if (!n) {
    n = ss.insertSheet('Natijalar');
    n.appendRow(['№', 'Vaqt', 'F.I.Sh.', 'Tuman', 'Muassasa', 'Test', 'Toʻgʻri', 'Jami savol', 'Ball', 'Foiz (%)', 'Natija', 'Vaqt sarfi', 'Chiqishlar', 'ID']);
    n.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#0b2a4a').setFontColor('#ffffff');
    n.setFrozenRows(1);
  }
  var j = ss.getSheetByName('Javoblar');
  if (!j) {
    j = ss.insertSheet('Javoblar');
    j.appendRow(['ID', 'Vaqt', 'F.I.Sh.', 'Tuman', 'Muassasa', 'Test', 'Ball', 'Maks. ball', 'Foiz (%)', 'Vaqt sarfi', 'Chiqishlar', 'Javoblar (JSON)', 'Savol vaqtlari (JSON)']);
    j.getRange(1, 1, 1, 13).setFontWeight('bold').setBackground('#0b2a4a').setFontColor('#ffffff');
    j.setFrozenRows(1);
  }
  return { n: n, j: j, tz: ss.getSpreadsheetTimeZone() };
}

function hms_(s) {
  s = Math.max(0, Math.round(s));
  var p = function (x) { return x < 10 ? '0' + x : '' + x; };
  return p(Math.floor(s / 3600)) + ':' + p(Math.floor(s % 3600 / 60)) + ':' + p(s % 60);
}

function submit_(p) {
  var cache = CacheService.getScriptCache();
  var raw = cache.get('s_' + p.token);
  if (!raw) throw new Error('Sessiya topilmadi yoki muddati tugagan. Testni qaytadan boshlang.');
  if (cache.get('d_' + p.token)) throw new Error('Bu urinish allaqachon yuborilgan.');
  var s = JSON.parse(raw), qs = loadQuestions_(s.test);
  var ans = p.answers || {}, times = p.times || {};

  var right = 0, score = 0, max = 0, review = [];
  qs.forEach(function (q) {
    max += q.pts;
    var a = ans[q.id];
    var ok = (a !== undefined && a !== null && Number(a) === q.correct);
    if (ok) { right++; score += q.pts; }
    review.push({ id: q.id, correct: q.correct, ok: ok });
  });
  score = Math.round(score * 10) / 10;
  max = Math.round(max * 10) / 10;
  var percent = max ? Math.round(score / max * 1000) / 10 : 0;
  var verdict = s.pass > 0 ? (percent >= s.pass ? 'Oʻtdi' : 'Oʻtmadi') : '';
  var elapsed = Math.min(Math.round((Date.now() - s.start) / 1000), s.minutes * 60);
  var timeStr = hms_(elapsed);

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = ensureSheets_();
    var stamp = Utilities.formatDate(new Date(), sh.tz, 'dd.MM.yyyy HH:mm:ss');
    var id = String(Date.now());
    sh.n.appendRow([sh.n.getLastRow(), stamp, s.name, s.region, s.org, s.test, right, qs.length, score, percent, verdict, timeStr, Number(p.cheats) || 0, id]);
    sh.j.appendRow([id, stamp, s.name, s.region, s.org, s.test, score, max, percent, timeStr, Number(p.cheats) || 0,
      JSON.stringify(ans).substring(0, 45000), JSON.stringify(times).substring(0, 45000)]);
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
  cache.put('d_' + p.token, '1', SESSION_HOURS * 3600);
  return { success: true, right: right, total: qs.length, score: score, max: max, percent: percent, verdict: verdict, time: timeStr, review: review };
}

// ---------------------------------------------------------------- ADMIN PANEL
function adminCheck_(p) {
  if (t_(p.pin) !== t_(ADMIN_PIN)) throw new Error('Admin PIN-kod notoʻgʻri!');
}
function stampText_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, ss_().getSpreadsheetTimeZone(), 'dd.MM.yyyy HH:mm:ss');
  return String(v);
}

function adminList_(p) {
  adminCheck_(p);
  var sh = ss_().getSheetByName('Javoblar');
  if (!sh || sh.getLastRow() < 2) return { success: true, data: [] };
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues();
  var out = rows.map(function (r) {
    return { id: String(r[0]), stamp: stampText_(r[1]), name: r[2], region: r[3], org: r[4], test: r[5], score: r[6], max: r[7], percent: r[8], time: String(r[9]), cheats: r[10] };
  }).reverse();
  return { success: true, data: out };
}

function adminAttempt_(p) {
  adminCheck_(p);
  var sh = ss_().getSheetByName('Javoblar');
  if (!sh || sh.getLastRow() < 2) throw new Error('Urinish topilmadi');
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 13).getValues(), r = null;
  for (var i = 0; i < rows.length; i++) { if (String(rows[i][0]) === String(p.id)) { r = rows[i]; break; } }
  if (!r) throw new Error('Urinish topilmadi');
  return {
    success: true,
    info: { stamp: stampText_(r[1]), name: r[2], region: r[3], org: r[4], test: r[5], score: r[6], max: r[7], percent: r[8], time: String(r[9]), cheats: r[10] },
    answers: JSON.parse(r[11] || '{}'),
    times: JSON.parse(r[12] || '{}'),
    questions: loadQuestions_(r[5])
  };
}

function adminDelete_(p) {
  adminCheck_(p);
  var ss = ss_(), done = 0;
  var j = ss.getSheetByName('Javoblar'), n = ss.getSheetByName('Natijalar');
  if (j && j.getLastRow() > 1) {
    var ids = j.getRange(2, 1, j.getLastRow() - 1, 1).getValues();
    for (var i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === String(p.id)) { j.deleteRow(i + 2); done++; }
  }
  if (n && n.getLastRow() > 1) {
    var nid = n.getRange(2, 14, n.getLastRow() - 1, 1).getValues();
    for (var k = nid.length - 1; k >= 0; k--) if (String(nid[k][0]) === String(p.id)) n.deleteRow(k + 2);
  }
  return { success: true, deleted: done };
}

// ---------------------------------------------------------------- yordamchi funksiyalar (qoʻlda ishga tushiriladi)
/** Xizmat varaqlarini yaratadi. Birinchi marta bir marta ishga tushiring. */
function setup() {
  var ss = ss_();
  var pin = ss.getSheetByName('PIN_Codes') || ss.insertSheet('PIN_Codes');
  if (pin.getLastRow() < 2) {
    pin.getRange('A1:C2').setValues([['PIN-kod', 'Holat', 'Izoh'], ['2026', 'Ochiq', 'Holat: Ochiq yoki Yopiq']]);
    pin.getRange(1, 1, 1, 3).setFontWeight('bold');
  }
  var t = ss.getSheetByName('Testlar') || ss.insertSheet('Testlar');
  if (t.getLastRow() < 1) {
    t.appendRow(['Test (varaq nomi)', 'Holat', 'Vaqt (daqiqa)', 'Oʻtish bali (%)', 'Izoh / koʻrsatma']);
    t.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#0b2a4a').setFontColor('#ffffff');
    t.setFrozenRows(1);
  }
  var have = t.getLastRow() > 1 ? t.getRange(2, 1, t.getLastRow() - 1, 1).getValues().map(function (r) { return t_(r[0]); }) : [];
  ss.getSheets().forEach(function (s) {
    var n = s.getName();
    if (SERVICE.indexOf(n) < 0 && n.charAt(0) !== '_' && have.indexOf(n) < 0) t.appendRow([n, 'Ochiq', 30, 60, '']);
  });
  ensureSheets_();
  SpreadsheetApp.flush();
  Logger.log('Tayyor. Testlar varagʻini tekshiring (vaqt, oʻtish bali).');
}

/** Barcha test varaqlarini tekshiradi: toʻgʻri javobi topilmagan qatorlarni koʻrsatadi. */
function tekshirish() {
  var problems = [], total = 0;
  listTests_(false).forEach(function (t) { total += loadQuestions_(t.name, problems).length; });
  var msg = problems.length ? ('Muammolar (' + problems.length + '):\n' + problems.join('\n')) : ('Hammasi joyida. Jami savollar: ' + total);
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('KRK').addItem('Varaqlarni tayyorlash (setup)', 'setup').addItem('Savollarni tekshirish', 'tekshirish').addToUi();
  } catch (e) {}
}
