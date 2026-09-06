/**
 * ============================================================================
 *  KUIZ DBS10072 — BACKEND (Google Apps Script Web App)
 * ============================================================================
 *  Fail ini adalah "otak" di sisi pelayan (server). Ia menyimpan API key
 *  dengan SELAMAT (key tidak pernah muncul dalam index.html yang awam).
 *
 *  FUNGSI:
 *   1) doPost  -> simpan keputusan kuiz. Setiap cubaan direkod dalam sheet
 *                 "Log". Leaderboard menyimpan MARKAH TERTINGGI setiap matrik
 *                 (satu baris satu matrik).
 *   2) doGet?action=leaderboard  -> pulangkan senarai leaderboard (JSON).
 *   3) doGet?action=stats        -> pulangkan statistik kelas (JSON).
 *   4) doGet?action=ai           -> proksi selamat ke Gemini AI Tutor.
 *
 * ----------------------------------------------------------------------------
 *  CARA PASANG (lakukan SEKALI sahaja):
 *  1. Buka Google Sheet pensyarah anda -> menu Extensions -> Apps Script.
 *  2. Padam kod lama, tampal SELURUH fail ini, kemudian Save.
 *  3. Menu (gear/Project Settings) -> "Script Properties" -> tambah:
 *        - GEMINI_KEY   = <API key Gemini anda>   (WAJIB untuk AI Tutor)
 *        - GEMINI_MODEL = gemini-2.5-flash        (pilihan; boleh tukar)
 *     >> JANGAN tulis key di dalam kod. Simpan dalam Script Properties sahaja.
 *  4. Deploy -> New deployment -> Type: "Web app".
 *        - Execute as: Me
 *        - Who has access: Anyone
 *     Salin URL /exec dan tampal ke index.html (pemalar BACKEND_URL).
 *  5. Setiap kali anda ubah kod ini, buat "Manage deployments" -> Edit ->
 *     Version: New version -> Deploy (supaya perubahan berkuat kuasa).
 * ============================================================================
 */

// ----- Tetapan -----
var SHEET_LEADERBOARD = 'Leaderboard';
var SHEET_LOG         = 'Log';
var SHEET_BANK        = 'Bank';
var LEADERBOARD_LIMIT = 200;
var DEFAULT_MODEL     = 'gemini-2.5-flash';

// ----- Utiliti -----
function prop_(key, fallback) {
  var v = PropertiesService.getScriptProperties().getProperty(key);
  return (v === null || v === '') ? fallback : v;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Halang "formula injection" (cth: =HYPERLINK, +, -, @) dalam Sheet/CSV. */
function safeCell_(val) {
  var s = (val === null || val === undefined) ? '' : String(val);
  s = s.replace(/[\u0000-\u001F\u007F]/g, '').trim();     // buang aksara kawalan
  if (/^[=\+\-@\t\r]/.test(s)) s = "'" + s;               // neutralkan formula
  return s.substring(0, 120);                             // had panjang
}

function sheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Baca badan permintaan sama ada JSON (text/plain) atau borang (parameter). */
function parseBody_(e) {
  if (e && e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch (err) { /* jatuh ke param */ }
  }
  if (e && e.parameter && e.parameter.id) return e.parameter;
  return null;
}

// ============================================================================
//  SIMPAN KEPUTUSAN (POST)
// ============================================================================
function doPost(e) {
  try {
    var data = parseBody_(e);
    // Tambah soalan ke Bank (dari borang pensyarah) — TIADA 'id' matrik.
    if (data && data.action === 'addq') return addQuestion_(data);
    // Simpan pilihan soalan (set kuiz) — TIADA 'id' matrik.
    if (data && data.action === 'setselection') return setSelection_(data);
    // Pengurusan berbilang kuiz — TIADA 'id' matrik.
    if (data && data.action === 'savequiz')      return saveQuiz_(data);
    if (data && data.action === 'setquizactive') return setQuizActive_(data);
    if (data && data.action === 'deletequiz')    return deleteQuiz_(data);
    if (!data || !data.id) return json_({ ok: false, error: 'Data tidak lengkap (tiada matrik).' });

    var name    = safeCell_(data.name);
    var id      = safeCell_(String(data.id).toUpperCase());
    var score   = Number(data.score)   || 0;
    var correct = Number(data.correct) || 0;
    var total   = Number(data.total)   || 0;
    var detail  = '';
    try { detail = data.detail ? JSON.stringify(data.detail).substring(0, 4000) : ''; } catch (er) {}
    var now = new Date();

    // (1) Rekod SETIAP cubaan ke Log
    var log = sheet_(SHEET_LOG, ['Masa', 'Nama', 'Matrik', 'Skor', 'Betul', 'Jumlah', 'Butiran(JSON)']);
    log.appendRow([now, name, id, score, correct, total, detail]);

    // (2) Upsert markah TERTINGGI ke Leaderboard (satu baris satu matrik)
    var lb = sheet_(SHEET_LEADERBOARD, ['Matrik', 'Nama', 'Skor Tertinggi', 'Betul', 'Jumlah', 'Bilangan Cubaan', 'Kemaskini']);
    upsertHighest_(lb, id, name, score, correct, total, now);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function upsertHighest_(sh, id, name, score, correct, total, now) {
  var last = sh.getLastRow();
  if (last < 2) {                       // kosong -> terus tambah
    sh.appendRow([id, name, score, correct, total, 1, now]);
    return;
  }
  var ids = sh.getRange(2, 1, last - 1, 1).getValues(); // lajur Matrik
  var rowIdx = -1;
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).toUpperCase() === id) { rowIdx = i + 2; break; }
  }
  if (rowIdx === -1) {                  // matrik baharu
    sh.appendRow([id, name, score, correct, total, 1, now]);
    return;
  }
  // Matrik sedia ada: naikkan kiraan cubaan, dan ganti HANYA jika skor lebih tinggi
  var attempts = Number(sh.getRange(rowIdx, 6).getValue()) || 0;
  sh.getRange(rowIdx, 6).setValue(attempts + 1);
  var existing = Number(sh.getRange(rowIdx, 3).getValue()) || 0;
  if (score > existing) {
    sh.getRange(rowIdx, 2, 1, 4).setValues([[name, score, correct, total]]); // Nama..Jumlah
    sh.getRange(rowIdx, 7).setValue(now);
  }
}

// ============================================================================
//  BANK SOALAN TERSUAI (ditambah oleh pensyarah)
// ============================================================================
/** Tambah satu soalan ke sheet "Bank". data.question ialah objek soalan penuh. */
function addQuestion_(data) {
  var q = data.question || {};
  if (!q.topic || !q.question || !q.options || !q.answer) {
    return json_({ ok: false, error: 'Soalan tidak lengkap (perlu topik, soalan, pilihan & jawapan).' });
  }
  if (!q.options[q.answer]) {
    return json_({ ok: false, error: 'Kunci jawapan tidak sepadan dengan pilihan.' });
  }
  var jsonStr = JSON.stringify(q).substring(0, 6000);
  var plainQ = String(q.question).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().substring(0, 150);
  var sh = sheet_(SHEET_BANK, ['Masa', 'Topik', 'Soalan', 'Jawapan', 'JSON(penuh)']);
  sh.appendRow([new Date(), safeCell_(q.topic), safeCell_(plainQ), safeCell_(q.answer), jsonStr]);
  return json_({ ok: true });
}

/** Pulangkan semua soalan tersuai daripada sheet "Bank". */
function readBank_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BANK);
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 5, sh.getLastRow() - 1, 1).getValues(); // lajur JSON(penuh)
  var out = [];
  rows.forEach(function (r, i) {
    try {
      var q = JSON.parse(r[0]);
      if (q && q.question && q.options && q.answer) { q.qid = 'C' + (i + 1); out.push(q); }
    } catch (er) { /* langkau baris rosak */ }
  });
  return out;
}

// ============================================================================
//  PILIHAN SOALAN (set kuiz yang ditetapkan pensyarah)
// ============================================================================
/** Simpan pilihan: { mode:'selected'|'random', ids:[qid,...] } dalam Script Properties. */
function setSelection_(data) {
  var mode = (data.mode === 'selected') ? 'selected' : 'random';
  var ids = [];
  if (data.ids && data.ids.length) {
    for (var i = 0; i < data.ids.length && ids.length < 500; i++) {
      if (typeof data.ids[i] === 'string') ids.push(data.ids[i]);
    }
  }
  PropertiesService.getScriptProperties().setProperty('QUIZ_SELECTION', JSON.stringify({ mode: mode, ids: ids }));
  return json_({ ok: true, mode: mode, count: ids.length });
}

/** Baca pilihan semasa. Default: mod rawak. */
function readSelection_() {
  var v = PropertiesService.getScriptProperties().getProperty('QUIZ_SELECTION');
  if (!v) return { mode: 'random', ids: [] };
  try {
    var o = JSON.parse(v);
    return { mode: (o.mode === 'selected') ? 'selected' : 'random', ids: Array.isArray(o.ids) ? o.ids : [] };
  } catch (e) { return { mode: 'random', ids: [] }; }
}

// ============================================================================
//  PENGURUSAN BERBILANG KUIZ (Quizzes)
// ============================================================================
var SHEET_QUIZZES = 'Quizzes';
var QUIZZES_HEADERS = ['id', 'Tajuk', 'Aktif', 'Soalan(qids JSON)', 'Tetapan(JSON)', 'Dicipta', 'Dikemaskini'];

function slugifyId_(title) {
  var base = String(title || 'kuiz').toLowerCase()
    .replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').substring(0, 30) || 'kuiz';
  return base + '-' + Math.random().toString(36).substring(2, 6);
}

function findQuizRow_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}

function saveQuiz_(data) {
  var title = safeCell_(data.title);
  if (!title) return json_({ ok: false, error: 'Tajuk kuiz diperlukan.' });
  var qids = Array.isArray(data.qids) ? data.qids.filter(function (x) { return typeof x === 'string'; }).slice(0, 500) : [];
  if (!qids.length) return json_({ ok: false, error: 'Pilih sekurang-kurangnya 1 soalan.' });
  var settings = (data.settings && typeof data.settings === 'object') ? data.settings : {};
  var now = new Date();
  var sh = sheet_(SHEET_QUIZZES, QUIZZES_HEADERS);
  var id = (data.id && typeof data.id === 'string') ? data.id : '';
  var active = (data.active === true);
  if (id) {
    var row = findQuizRow_(sh, id);
    if (row === -1) return json_({ ok: false, error: 'Kuiz tidak dijumpai.' });
    var created = sh.getRange(row, 6).getValue() || now;
    sh.getRange(row, 2, 1, 6).setValues([[title, active, JSON.stringify(qids), JSON.stringify(settings), created, now]]);
    return json_({ ok: true, id: id });
  }
  id = slugifyId_(title);
  sh.appendRow([id, title, active, JSON.stringify(qids), JSON.stringify(settings), now, now]);
  return json_({ ok: true, id: id });
}

function setQuizActive_(data) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUIZZES);
  if (!sh) return json_({ ok: false, error: 'Tiada kuiz.' });
  var row = findQuizRow_(sh, String(data.id || ''));
  if (row === -1) return json_({ ok: false, error: 'Kuiz tidak dijumpai.' });
  sh.getRange(row, 3).setValue(data.active === true);
  sh.getRange(row, 7).setValue(new Date());
  return json_({ ok: true });
}

function deleteQuiz_(data) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUIZZES);
  if (!sh) return json_({ ok: false, error: 'Tiada kuiz.' });
  var row = findQuizRow_(sh, String(data.id || ''));
  if (row === -1) return json_({ ok: false, error: 'Kuiz tidak dijumpai.' });
  sh.deleteRow(row);
  return json_({ ok: true });
}

function readQuizzes_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUIZZES);
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  return rows.map(function (r) {
    var qids = []; try { qids = JSON.parse(r[3]) || []; } catch (e) {}
    return { id: r[0], title: r[1], active: (r[2] === true || String(r[2]).toLowerCase() === 'true'), count: qids.length };
  });
}

function readQuiz_(id) {
  id = String(id || '');
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUIZZES);
  if (!sh || sh.getLastRow() < 2) return null;
  var row = findQuizRow_(sh, id);
  if (row === -1) return null;
  var v = sh.getRange(row, 1, 1, 5).getValues()[0];
  var qids = []; try { qids = JSON.parse(v[3]) || []; } catch (e) {}
  var settings = {}; try { settings = JSON.parse(v[4]) || {}; } catch (e) {}
  return { id: v[0], title: v[1], active: (v[2] === true || String(v[2]).toLowerCase() === 'true'), qids: qids, settings: settings };
}

// ============================================================================
//  BACA DATA (GET): leaderboard / stats / AI tutor
// ============================================================================
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'leaderboard';
  try {
    if (action === 'ai')    return aiTutor_(e);
    if (action === 'stats') return json_({ ok: true, stats: computeStats_() });
    if (action === 'bank')  return json_({ ok: true, bank: readBank_() });
    if (action === 'selection') return json_({ ok: true, selection: readSelection_() });
    if (action === 'log')       return json_({ ok: true, log: readLog_() });
    if (action === 'quizzes')   return json_({ ok: true, quizzes: readQuizzes_() });
    if (action === 'quiz')      return json_({ ok: true, quiz: readQuiz_(e.parameter.id) });
    return json_({ ok: true, leaderboard: readLeaderboard_() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function readLeaderboard_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LEADERBOARD);
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  var out = rows.map(function (r) {
    return { id: r[0], name: r[1], score: Number(r[2]) || 0, correct: Number(r[3]) || 0,
             total: Number(r[4]) || 0, attempts: Number(r[5]) || 0 };
  }).filter(function (x) { return !/^TEST[-_]/i.test(String(x.id || '')); }); // sembunyi rekod ujian
  out.sort(function (a, b) { return b.score - a.score; });
  return out.slice(0, LEADERBOARD_LIMIT);
}

/** Pulangkan SETIAP cubaan (untuk muat turun log penuh). Rekod ujian ditapis. */
function readLog_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues(); // Masa,Nama,Matrik,Skor,Betul,Jumlah
  var out = [];
  rows.forEach(function (r) {
    if (/^TEST[-_]/i.test(String(r[2] || ''))) return; // langkau rekod ujian
    out.push({ time: r[0], name: r[1], id: r[2], score: Number(r[3]) || 0, correct: Number(r[4]) || 0, total: Number(r[5]) || 0 });
  });
  return out;
}

function computeStats_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  var stats = { attempts: 0, students: 0, avgScore: 0, avgCorrect: 0, perQuestion: [] };
  if (!sh || sh.getLastRow() < 2) return stats;

  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  var sumScore = 0, sumCorrect = 0, uniq = {}, qMap = {};

  rows.forEach(function (r) {
    if (/^TEST[-_]/i.test(String(r[2] || ''))) return; // langkau rekod ujian (ghost)
    stats.attempts++;
    sumScore   += Number(r[3]) || 0;
    sumCorrect += Number(r[4]) || 0;
    uniq[String(r[2]).toUpperCase()] = true;
    if (r[6]) {
      try {
        JSON.parse(r[6]).forEach(function (d) {
          var k = d.qid;
          if (!qMap[k]) qMap[k] = { qid: k, label: d.label || k, seen: 0, wrong: 0, timeSum: 0 };
          qMap[k].seen++;
          if (!d.correct) qMap[k].wrong++;
          qMap[k].timeSum += Number(d.time) || 0;
        });
      } catch (er) {}
    }
  });

  stats.students   = Object.keys(uniq).length;
  stats.avgScore   = Math.round(sumScore / stats.attempts);
  stats.avgCorrect = Math.round((sumCorrect / stats.attempts) * 10) / 10;
  stats.perQuestion = Object.keys(qMap).map(function (k) {
    var q = qMap[k];
    return { qid: q.qid, label: q.label, seen: q.seen, wrong: q.wrong,
             missRate: Math.round((q.wrong / q.seen) * 100),
             avgTime: Math.round(q.timeSum / q.seen) };
  }).sort(function (a, b) { return b.missRate - a.missRate; });

  return stats;
}

// ============================================================================
//  AI TUTOR — proksi SELAMAT ke Gemini (key kekal di server)
// ============================================================================
function aiTutor_(e) {
  var key = prop_('GEMINI_KEY', '');
  if (!key) return json_({ ok: false, error: 'GEMINI_KEY belum ditetapkan dalam Script Properties.' });

  var model  = prop_('GEMINI_MODEL', DEFAULT_MODEL);
  var q      = (e.parameter.q   || '').substring(0, 1200);
  var sel    = (e.parameter.sel || '').substring(0, 200);
  var ans    = (e.parameter.ans || '').substring(0, 200);

  var prompt =
    'Anda ialah AI Tutor Fizik Sains untuk kursus diploma politeknik (DBS10072 Science). ' +
    'Soalan: ' + q + '\n' +
    'Pilihan pelajar: ' + (sel || 'Tiada (masa tamat)') + '\n' +
    'Jawapan betul: ' + ans + '\n' +
    'Berikan penerangan ringkas (maksimum 3-4 ayat), mesra pelajar, dalam Bahasa Melayu, ' +
    'dan 1 tip praktikal untuk mengingati konsep ini semasa peperiksaan.';

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent';
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': key },
    payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
    muteHttpExceptions: true
  });

  var body = {};
  try { body = JSON.parse(res.getContentText()); } catch (er) {
    return json_({ ok: false, error: 'Respons AI tidak sah.' });
  }
  if (body.error) return json_({ ok: false, error: body.error.message || 'Ralat Gemini.' });

  var text = '';
  try { text = body.candidates[0].content.parts[0].text; } catch (er) { text = ''; }
  return json_({ ok: true, text: text || 'AI Tutor tiada respons buat masa ini.' });
}

// ============================================================================
//  FUNGSI PERSEDIAAN — jalankan SEKALI di editor untuk beri kebenaran internet
// ============================================================================
/**
 * Jalankan fungsi ini SEKALI di editor Apps Script (pilih 'grantPermissions'
 * di dropdown -> Run). Google akan minta kebenaran "Connect to an external
 * service" — klik Allow. Selepas itu AI Tutor akan berfungsi.
 */
function grantPermissions() {
  var res = UrlFetchApp.fetch('https://www.googleapis.com/discovery/v1/apis', { muteHttpExceptions: true });
  Logger.log('Kebenaran internet OK. Status: ' + res.getResponseCode());
}

/**
 * Uji AI Tutor terus dari editor (pilih 'testAiTutor' -> Run, lihat Logs).
 * Ini juga akan mencetuskan permintaan kebenaran internet jika belum diberi.
 */
function testAiTutor() {
  var out = aiTutor_({ parameter: { q: 'Apakah unit SI bagi jisim?', sel: 'A', ans: 'Kilogram' } });
  Logger.log(out.getContent());
}
