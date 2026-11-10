/**
 * Chișinău 2040 — бэкенд опроса жителей на Google Apps Script.
 * Хранит ответы в листе "responses" текущей Google-таблицы.
 *   POST  (тело — JSON)       → сохранить или обновить ответ (по id устройства)
 *   GET   ?action=stats       → агрегированные результаты для вкладки «Результаты»
 * Инструкция по установке — в README.md.
 */

const SHEET_NAME = 'responses';
const SHOW_TEXT_IN_RESULTS = true;   // false — не показывать свободные ответы публично
const MAX_QUOTES = 12;
const HEADERS = ['id', 'updatedAt', 'lang', 'sector', 'age', 'tenure', 'commute', 'mode',
  'thr', 'pairs', 'budget', 'keep', 'prob', 'identity', 'remove', 'diaspora', 'hidden'];
// Колонка "hidden": поставьте в ней 1, чтобы скрыть ответ из публичных результатов (модерация).

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Защита от формул в ячейках: текст, начинающийся с = + - @, сохраняем как текст.
function safeText_(v, max) {
  let s = String(v == null ? '' : v).slice(0, max).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}
function intOrNull_(v, min, max) {
  if (v === null || v === undefined || v === '') return null;
  const n = Math.round(Number(v));
  return isFinite(n) && n >= min && n <= max ? n : null;
}
function pins_(arr, max) {
  if (!Array.isArray(arr)) return [];
  return arr.slice(0, max).map(p => ({
    lat: Math.round(Number(p.lat) * 1e5) / 1e5,
    lng: Math.round(Number(p.lng) * 1e5) / 1e5,
    note: safeText_(p.note, 300)
  })).filter(p => p.lat > 46.7 && p.lat < 47.3 && p.lng > 28.4 && p.lng < 29.3);
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }
  if (d.website) return json_({ ok: true });                      // honeypot: боты заполняют скрытое поле
  const id = String(d.id || '');
  if (!/^[A-Za-z0-9.\-]{8,64}$/.test(id)) return json_({ ok: false, error: 'bad_id' });

  const thr = Array.isArray(d.thr) ? d.thr.slice(0, 4).map(v => intOrNull_(v, 0, 100)) : [];
  const pairs = Array.isArray(d.pairs) ? d.pairs.slice(0, 5).map(v => (v === 0 || v === 1) ? v : null) : [];
  let budget = Array.isArray(d.budget) ? d.budget.slice(0, 6).map(v => intOrNull_(v, 0, 100) || 0) : [];
  if (budget.reduce((a, b) => a + b, 0) > 100) budget = budget.map(() => 0);

  const row = [
    id, new Date(), d.lang === 'ru' ? 'ru' : 'ro',
    intOrNull_(d.sector, 0, 6), intOrNull_(d.age, 0, 4), intOrNull_(d.tenure, 0, 2),
    intOrNull_(d.commute, 0, 240), intOrNull_(d.mode, 0, 4),
    JSON.stringify(thr), JSON.stringify(pairs), JSON.stringify(budget),
    JSON.stringify(pins_(d.keep, 3)), JSON.stringify(pins_(d.prob, 5)),
    safeText_(d.identity, 1000), safeText_(d.remove, 1000), safeText_(d.diaspora, 1000)
  ];

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const found = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1)
      .createTextFinder(id).matchEntireCell(true).findNext();
    if (found && found.getRow() > 1) {
      sh.getRange(found.getRow(), 1, 1, row.length).setValues([row]);   // колонку hidden не трогаем
    } else {
      sh.appendRow(row.concat(['']));
    }
  } finally {
    lock.releaseLock();
  }
  CacheService.getScriptCache().remove('stats');
  return json_({ ok: true });
}

function doGet(e) {
  if (!e || !e.parameter || e.parameter.action !== 'stats') return json_({ ok: true, service: 'chisinau2040-survey' });
  const cache = CacheService.getScriptCache();
  const hit = cache.get('stats');
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  const out = JSON.stringify(computeStats_());
  try { cache.put('stats', out, 60); } catch (err) { /* слишком большой для кэша — просто не кэшируем */ }
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

function computeStats_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  const res = { ok: true, count: 0, keep: [], prob: [], thrMed: [null, null, null, null],
    pairs: [[0, 0], [0, 0], [0, 0], [0, 0], [0, 0]], budgetAvg: [0, 0, 0, 0, 0, 0], why: [], remove: [] };
  if (n < 1) return res;
  const rows = sh.getRange(2, 1, n, HEADERS.length).getValues();
  const col = name => HEADERS.indexOf(name);
  const parse = s => { try { return JSON.parse(s); } catch (err) { return []; } };
  const thrs = [[], [], [], []];
  const budgetSum = [0, 0, 0, 0, 0, 0];
  const texts = [];
  rows.forEach(r => {
    if (!r[0]) return;
    res.count++;
    const hidden = String(r[col('hidden')]) === '1';
    parse(r[col('keep')]).forEach(p => {
      res.keep.push([p.lat, p.lng]);
      if (!hidden && p.note) texts.push({ t: r[col('updatedAt')], why: String(p.note) });
    });
    parse(r[col('prob')]).forEach(p => res.prob.push([p.lat, p.lng]));
    parse(r[col('thr')]).forEach((v, i) => { if (typeof v === 'number' && i < 4) thrs[i].push(v); });
    parse(r[col('pairs')]).forEach((v, i) => { if (i < 5 && (v === 0 || v === 1)) res.pairs[i][v]++; });
    parse(r[col('budget')]).forEach((v, i) => { if (i < 6) budgetSum[i] += Number(v) || 0; });
    if (!hidden && r[col('remove')]) texts.push({ t: r[col('updatedAt')], remove: String(r[col('remove')]) });
  });
  res.thrMed = thrs.map(a => {
    if (!a.length) return null;
    a.sort((x, y) => x - y);
    const m = Math.floor(a.length / 2);
    return a.length % 2 ? a[m] : Math.round((a[m - 1] + a[m]) / 2);
  });
  res.budgetAvg = budgetSum.map(s => Math.round(s / res.count * 10) / 10);
  if (SHOW_TEXT_IN_RESULTS) {
    texts.sort((a, b) => new Date(b.t) - new Date(a.t));
    const clean = s => s.replace(/^'/, '').trim();
    res.why = texts.filter(x => x.why).slice(0, MAX_QUOTES).map(x => clean(x.why));
    res.remove = texts.filter(x => x.remove).slice(0, MAX_QUOTES).map(x => clean(x.remove));
  }
  return res;
}
