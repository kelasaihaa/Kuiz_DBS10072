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
//  BACA DATA (GET): leaderboard / stats / AI tutor
// ============================================================================
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'leaderboard';
  try {
    if (action === 'ai')    return aiTutor_(e);
    if (action === 'stats') return json_({ ok: true, stats: computeStats_() });
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
  });
  out.sort(function (a, b) { return b.score - a.score; });
  return out.slice(0, LEADERBOARD_LIMIT);
}

function computeStats_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  var stats = { attempts: 0, students: 0, avgScore: 0, avgCorrect: 0, perQuestion: [] };
  if (!sh || sh.getLastRow() < 2) return stats;

  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  var sumScore = 0, sumCorrect = 0, uniq = {}, qMap = {};

  rows.forEach(function (r) {
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
