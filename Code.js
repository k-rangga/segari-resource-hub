/**
 * Segari Knowledge Base — server side.
 * Documents live in a spreadsheet (sheets "Dokumen" and "Versi"), created and filled from Seed.js on first load.
 * Mancu answers with Gemini, grounded on the latest revision of every document.
 *
 * Script Properties:
 *   GEMINI_API_KEY      required for Mancu (set it yourself; never commit it)
 *   GEMINI_MODEL        optional, default gemini-flash-latest
 *   MANCU_DAILY_LIMIT   optional, default 300 questions per day for the whole app
 */

const TZ = 'Asia/Jakarta';
const SS_PROP = 'ARTICLE_SPREADSHEET_ID';
const DOC_SHEET = 'Dokumen';
const VER_SHEET = 'Versi';
const DOC_COLS = ['ID', 'Kode', 'Judul', 'Nama pendek', 'Jenis', 'Area', 'Halaman', 'Kata kunci', 'Persetujuan', 'Status', 'Diperbarui'];
const VER_COLS = ['ID dokumen', 'Rev', 'Berlaku', 'Catatan', 'Isi'];
const TYPES = ['sop', 'ik', 'spec', 'safety'];
const AREAS = ['pc', 'dmg', 'fc', 'storage', 'outbound'];
const TYPE_CODE = { sop: 'SOP', ik: 'IK', spec: 'Specbook', safety: 'SOP' };
const DEFAULT_MODEL = 'gemini-flash-latest';
const DEFAULT_DAILY_LIMIT = 300;
const CONTEXT_CHARS = 60000;

// ---------------------------------------------------------------- web app

function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  t.appUrl = ScriptApp.getService().getUrl() || '';
  return t.evaluate()
    .setTitle('Segari Knowledge Base')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// ---------------------------------------------------------------- library

/** Every published document with its revisions, newest revision first. */
function getLibrary() {
  return readLibrary_().filter(d => d.status !== 'draf');
}

/**
 * Saves a draft or publishes an article from writer mode.
 * @param {{id?:string,title:string,type:string,areas:string[],body:string,versions:{rev:number,effectiveDate:string,note:string}[],keywords:string[]}} a
 * @param {'draft'|'published'} status
 */
function saveArticle(a, status) {
  a = clean_(a);
  if (status === 'published' && !a.title) throw new Error('Isi judul artikel dulu');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = library_();
    const id = a.id || 'artikel-' + Utilities.getUuid().slice(0, 8);
    const docRow = [id, TYPE_CODE[a.type], a.title, a.title, a.type, a.areas.join(', '), '', a.keywords.join(', '), '',
      status === 'published' ? 'terbit' : 'draf', now_()];
    upsertDoc_(ss.getSheetByName(DOC_SHEET), docRow);
    replaceVersions_(ss.getSheetByName(VER_SHEET), id, a.versions.map(v => [id, v.rev, v.effectiveDate, v.note, a.body]));
    return readLibrary_().find(d => d.id === id);
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- Mancu (Gemini)

/**
 * Answers a staff question from the latest revision of every document.
 * Returns null when no API key is set, so the page can fall back to a document search.
 * @return {{intro:string, items:string[], note:string, sources:{docId:string, sec:string, label:string}[]}|null}
 */
function askMancu(question) {
  question = String(question || '').trim().slice(0, 500);
  if (!question) throw new Error('Pertanyaan kosong');
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('GEMINI_API_KEY');
  if (!key) return null;
  countQuestion_(props);

  const chunks = pickChunks_(question, chunks_(getLibrary()));
  const context = chunks.map(c => '[' + c.id + '] ' + c.label + '\n' + c.text).join('\n\n---\n\n');
  const model = props.getProperty('GEMINI_MODEL') || DEFAULT_MODEL;
  const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': key },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      systemInstruction: { parts: [{ text: SYSTEM_PROMPT }] },
      contents: [{ role: 'user', parts: [{ text: 'DOKUMEN:\n\n' + context + '\n\nPERTANYAAN STAF:\n' + question }] }],
      generationConfig: { temperature: 0.2, responseMimeType: 'application/json', responseSchema: ANSWER_SCHEMA }
    })
  });
  if (res.getResponseCode() !== 200) {
    console.error('Gemini ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 500));
    throw new Error('Mancu sedang tidak tersedia. Coba lagi sebentar.');
  }
  let out;
  try {
    const body = JSON.parse(res.getContentText());
    out = JSON.parse(body.candidates[0].content.parts.map(p => p.text || '').join(''));
  } catch (e) {
    console.error('Gemini parse: ' + res.getContentText().slice(0, 500));
    throw new Error('Mancu sedang tidak tersedia. Coba lagi sebentar.');
  }
  const byId = {};
  chunks.forEach(c => { byId[c.id] = c; });
  const seen = {};
  return {
    intro: String(out.intro || ''),
    items: (out.items || []).map(String).filter(Boolean).slice(0, 8),
    note: String(out.note || ''),
    sources: (out.sourceIds || []).map(id => byId[id]).filter(c => c && !seen[c.id] && (seen[c.id] = true))
      .map(c => ({ docId: c.docId, sec: c.sec, label: c.label }))
  };
}

const SYSTEM_PROMPT = [
  'Kamu adalah Mancu, asisten SOP gudang Segari. Staf gudang (inbound, QC, TL/Captain) bertanya lewat aplikasi.',
  'Jawab HANYA dari DOKUMEN yang diberikan. Jangan mengarang aturan, angka, jam, atau nama grup.',
  'Jika dokumen tidak memuat jawabannya, katakan di intro bahwa kamu belum menemukan jawaban pasti di dokumen yang terindeks, dan sarankan bertanya ke TL/Captain shift.',
  'Bahasa Indonesia, kalimat pendek dan jelas, sentence case. Boleh memakai "Saya".',
  'intro: satu kalimat jawaban langsung. items: langkah atau poin penting secara berurutan, maksimal 6, tanpa nomor di depan.',
  'note: peringatan penting jika ada (mis. kondisi inbound ditolak), selain itu string kosong.',
  'sourceIds: id dokumen dalam kurung siku (mis. c3) yang benar-benar dipakai untuk menjawab.'
].join('\n');

const ANSWER_SCHEMA = {
  type: 'OBJECT',
  properties: {
    intro: { type: 'STRING' },
    items: { type: 'ARRAY', items: { type: 'STRING' } },
    note: { type: 'STRING' },
    sourceIds: { type: 'ARRAY', items: { type: 'STRING' } }
  },
  required: ['intro', 'items', 'note', 'sourceIds']
};

/** One chunk per "# " section of each document's latest revision. Section ids match the reader (s0, s1, …). */
function chunks_(docs) {
  const out = [];
  docs.forEach(d => {
    const v = d.versions[0];
    if (!v || !v.body) return;
    const head = d.code === d.title ? d.title : d.code + ' ' + d.title;
    let cur = null, i = -1;
    String(v.body).split('\n').forEach(line => {
      if (/^# /.test(line)) { i++; cur = { title: line.slice(2).trim(), lines: [] }; push(); return; }
      if (!cur) { i++; cur = { title: d.title, lines: [] }; push(); }
      cur.lines.push(line);
    });
    function push() {
      const c = cur, idx = i;
      out.push({ docId: d.id, sec: 's' + idx, label: d.code + ' · ' + c.title, keywords: d.keywords.join(' '),
        get text() { return head + ' (Rev ' + v.rev + ')\n' + c.title + '\n' + c.lines.join('\n').trim(); } });
    }
  });
  out.forEach((c, n) => { c.id = 'c' + (n + 1); });
  return out;
}

const STOP = 'apa yang harus jika dan untuk bagaimana berapa dari ada atau dengan saat tidak sudah pada ini itu bisa kalau mana kapan siapa'.split(' ');

/** Most relevant chunks first, within the context budget. */
function pickChunks_(question, chunks) {
  const words = question.toLowerCase().split(/[^a-z0-9]+/).filter(w => w.length >= 3 && STOP.indexOf(w) < 0);
  const scored = chunks.map(c => {
    const hay = (c.label + ' ' + c.keywords + ' ' + c.text).toLowerCase();
    return { c: c, n: words.reduce((s, w) => s + (hay.indexOf(w) >= 0 ? 1 : 0), 0) };
  }).sort((a, b) => b.n - a.n);
  const picked = [];
  let size = 0;
  scored.forEach(o => {
    const len = o.c.text.length;
    if (size + len <= CONTEXT_CHARS) { picked.push(o.c); size += len; }
  });
  return picked;
}

function countQuestion_(props) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const day = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd');
    const limit = Number(props.getProperty('MANCU_DAILY_LIMIT')) || DEFAULT_DAILY_LIMIT;
    const parts = String(props.getProperty('MANCU_COUNT') || '').split(':');
    const n = parts[0] === day ? Number(parts[1]) || 0 : 0;
    if (n >= limit) throw new Error('Batas pertanyaan Mancu hari ini sudah tercapai. Coba lagi besok.');
    props.setProperty('MANCU_COUNT', day + ':' + (n + 1));
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- spreadsheet

/** Opens the library spreadsheet, creating it and its sheets (and seeding them) when missing. */
function library_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(SS_PROP);
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create('Segari Knowledge Base – Dokumen');
  if (!id) props.setProperty(SS_PROP, ss.getId());
  let docs = ss.getSheetByName(DOC_SHEET), vers = ss.getSheetByName(VER_SHEET);
  if (!docs) docs = setupSheet_(ss, DOC_SHEET, DOC_COLS, ['A', 'G', 'K']);
  if (!vers) vers = setupSheet_(ss, VER_SHEET, VER_COLS, ['A', 'B', 'C']);
  const blank = ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 2 && blank.getLastRow() === 0) ss.deleteSheet(blank);
  if (docs.getLastRow() < 2) seed_(docs, vers);
  return ss;
}

function setupSheet_(ss, name, cols, textCols) {
  const sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
  sh.setFrozenRows(1);
  textCols.forEach(c => sh.getRange(c + ':' + c).setNumberFormat('@'));
  return sh;
}

function seed_(docs, vers) {
  const stamp = now_();
  docs.getRange(2, 1, SEED_DOCS.length, DOC_COLS.length).setValues(SEED_DOCS.map(d =>
    [d.id, d.code, d.title, d.short, d.type, d.areas.join(', '), String(d.pages || ''), d.keywords.join(', '), d.signoff, 'terbit', stamp]));
  const rows = [];
  SEED_DOCS.forEach(d => d.versions.forEach(v => rows.push([d.id, String(v.rev), v.effectiveDate, v.note, v.body])));
  if (rows.length) vers.getRange(vers.getLastRow() + 1, 1, rows.length, VER_COLS.length).setValues(rows);
}

function readLibrary_() {
  const ss = library_();
  const docs = values_(ss.getSheetByName(DOC_SHEET), DOC_COLS.length);
  const vers = values_(ss.getSheetByName(VER_SHEET), VER_COLS.length);
  const byDoc = {};
  vers.forEach(r => {
    const id = String(r[0]).trim();
    if (!id) return;
    (byDoc[id] = byDoc[id] || []).push({ rev: Number(r[1]) || 0, effectiveDate: dateText_(r[2]), note: String(r[3] || ''), body: String(r[4] || '') });
  });
  const list = s => String(s || '').split(',').map(x => x.trim()).filter(Boolean);
  return docs.filter(r => String(r[0]).trim()).map(r => {
    const id = String(r[0]).trim();
    const type = TYPES.indexOf(String(r[4]).trim()) >= 0 ? String(r[4]).trim() : 'sop';
    return {
      id: id, code: String(r[1] || TYPE_CODE[type]), title: String(r[2] || id), short: String(r[3] || r[2] || id), type: type,
      areas: list(r[5]).filter(a => AREAS.indexOf(a) >= 0), pages: String(r[6] || ''), keywords: list(r[7]),
      signoff: String(r[8] || ''), status: String(r[9] || 'terbit').trim().toLowerCase(),
      versions: (byDoc[id] || []).sort((a, b) => b.rev - a.rev || (b.effectiveDate > a.effectiveDate ? 1 : -1))
    };
  });
}

function values_(sh, width) {
  const n = sh.getLastRow();
  return n > 1 ? sh.getRange(2, 1, n - 1, width).getValues() : [];
}

function upsertDoc_(sh, row) {
  const ids = values_(sh, 1).map(r => String(r[0]));
  const i = ids.indexOf(row[0]);
  if (i >= 0) sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
}

function replaceVersions_(sh, id, rows) {
  const ids = values_(sh, 1).map(r => String(r[0]));
  for (let i = ids.length - 1; i >= 0; i--) if (ids[i] === id) sh.deleteRow(i + 2);
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, VER_COLS.length).setValues(rows.map(r => [r[0], String(r[1]), r[2], r[3], r[4]]));
}

function dateText_(v) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v || '');
}

function now_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
}

function clean_(a) {
  a = a || {};
  const str = v => String(v == null ? '' : v).trim();
  let versions = (a.versions || []).map(v => ({ rev: Number(v.rev) || 1, effectiveDate: str(v.effectiveDate), note: str(v.note) }));
  if (!versions.length) versions = [{ rev: 1, effectiveDate: '', note: 'Rilis pertama' }];
  return {
    id: str(a.id),
    title: str(a.title),
    type: TYPES.indexOf(a.type) >= 0 ? a.type : 'sop',
    areas: (a.areas || []).filter(x => AREAS.indexOf(x) >= 0),
    body: String(a.body || ''),
    versions: versions,
    keywords: (a.keywords || []).map(str).filter(Boolean)
  };
}
