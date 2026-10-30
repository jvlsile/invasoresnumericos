/* Patrulla Numérica · registro de partidas en Google Sheets (versión reforzada)
   Pega este código en Extensiones > Apps Script de tu hoja de cálculo. */
const SHEET = 'Partidas';
const STUDENT_KEY = 'CAMBIA-ESTA-CLAVE';      // la misma que pondrás en config.js (key)
const TEACHER_PW  = 'CAMBIA-LA-CONTRASEÑA';   // la que escribirás en el Panel del profesor
const HEAD = ['id','fecha','apodo','curso','tipo','modo','puntos','nivel','oleada','naves','disparos','aciertos','segundos','falladas'];
const MAX_ROWS_PER_MIN = 150, MAX_ROWS_TOTAL = 20000, TOKEN_DAYS = 30;
const TYPES = ['mix','sum','res','mul','div','comb'], MODES = ['solo','local','host','guest'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET);
  if (!s) { s = ss.insertSheet(SHEET); s.appendRow(HEAD); }
  return s;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function same_(a, b) { a = String(a); b = String(b); let r = a.length ^ b.length; for (let i = 0; i < Math.max(a.length, b.length); i++) r |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0); return r === 0; }
function num_(v, max) { v = Number(v); return isFinite(v) ? Math.max(0, Math.min(max, Math.round(v))) : 0; }
function txt_(v, n) { let t = String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').slice(0, n); return /^[=+\-@]/.test(t) ? "'" + t : t; }
function sign_(exp) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(String(exp), TEACHER_PW + '|tok')); }
function tokenOk_(t) {
  const p = String(t || '').split('.'); if (p.length !== 2) return false;
  return Number(p[0]) > Date.now() && same_(p[1], sign_(p[0]));
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents), cache = CacheService.getScriptCache();
    if (d.action === 'login') {
      const bad = Number(cache.get('badlogin') || 0);
      if (bad >= 5) return out_({ ok: false, error: 'wait' });
      if (!same_(d.pw, TEACHER_PW)) { cache.put('badlogin', String(bad + 1), 600); return out_({ ok: false, error: 'pw' }); }
      const exp = Date.now() + TOKEN_DAYS * 86400000;
      return out_({ ok: true, token: exp + '.' + sign_(exp) });
    }
    if (d.action === 'stats') {
      if (!tokenOk_(d.token)) return out_({ ok: false, error: 'auth' });
      const s = sheet_(), last = s.getLastRow();
      if (last < 2) return out_({ ok: true, rows: [] });
      const from = Math.max(2, last - 1999);
      const v = s.getRange(from, 1, last - from + 1, HEAD.length).getValues();
      return out_({ ok: true, rows: v.map(r => ({ id: r[0], ts: r[1], nick: r[2], course: r[3], type: r[4], mode: r[5],
        score: r[6], level: r[7], wave: r[8], kills: r[9], shots: r[10], hits: r[11], secs: r[12], missed: r[13] })) });
    }
    // alta de partidas
    if (!same_(d.key, STUDENT_KEY)) return out_({ ok: false, error: 'key' });
    const lock = LockService.getScriptLock(); lock.waitLock(10000);
    try {
      const minute = 'rows' + Math.floor(Date.now() / 60000), used = Number(cache.get(minute) || 0);
      const s = sheet_(), last = s.getLastRow();
      if (last > MAX_ROWS_TOTAL) return out_({ ok: false, error: 'full' });
      const seen = new Set(last > 1 ? s.getRange(2, 1, last - 1, 1).getValues().flat() : []);
      const now = new Date();
      const rows = (Array.isArray(d.rows) ? d.rows : []).slice(0, 50)
        .filter(r => r && /^[a-z0-9]{6,24}$/.test(String(r.id)) && !seen.has(String(r.id)))
        .slice(0, Math.max(0, MAX_ROWS_PER_MIN - used))
        .map(r => [String(r.id), now, txt_(r.nick, 20), num_(r.course, 6), TYPES.includes(r.type) ? r.type : 'mix', MODES.includes(r.mode) ? r.mode : 'solo',
          num_(r.score, 1e6), num_(r.level, 99), num_(r.wave, 999), num_(r.kills, 9999), num_(r.shots, 9999), num_(r.hits, 9999), num_(r.secs, 36000), txt_(r.missed, 300)]);
      if (rows.length) { s.getRange(s.getLastRow() + 1, 1, rows.length, HEAD.length).setValues(rows); cache.put(minute, String(used + rows.length), 90); }
      return out_({ ok: true, added: rows.length });
    } finally { lock.releaseLock(); }
  } catch (err) { return out_({ ok: false, error: 'bad_request' }); }
}

function doGet() { return out_({ ok: false, error: 'use_post' }); }
