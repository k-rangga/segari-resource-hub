/**
 * Segari Knowledge Base — server side.
 * Documents live in the Knowledge Base spreadsheet (sheets "Dokumen" and "Versi"), filled from Seed.js when empty.
 * Mancu answers with Gemini, grounded on the latest revision of every document.
 * Publishing a changed article adds a revision whose changelog Gemini writes; compileChangelogs() backfills empty ones.
 *
 * Script Properties:
 *   GEMINI_API_KEY      required for Mancu (set it yourself; never commit it)
 *   GEMINI_MODEL        optional, tried before the defaults (gemini-3.1-flash-lite, then gemini-3.5-flash)
 *   MANCU_DAILY_LIMIT   optional, default 300 questions per day for the whole app
 */

const TZ = 'Asia/Jakarta';
const SPREADSHEET_ID = '1HwcfPdSgeYgLKs_QQovOOnkfipgspQ9nnnDBPbIkMkY';
const DOC_SHEET = 'Dokumen';
const VER_SHEET = 'Versi';
const PROMPT_SHEET = 'Pertanyaan';
const DRAFT_SHEET = 'Draf';
const DRAFT_COLS = ['Kunci', 'ID dokumen', 'Disimpan', 'Judul', 'Jenis', 'Area', 'Berlaku', 'Kata kunci', 'Isi'];
const DRAFT_LIMIT = 100;     // the app is public: keeps the sheet from being flooded
const DRAFT_KEEP_DAYS = 30;
const BODY_LIMIT = 50000;    // characters a Google Sheets cell can hold
const FEEDBACK_SHEET = 'Masukan Mancu';
const FEEDBACK_COLS = ['Waktu', 'ID dokumen', 'Pertanyaan', 'Penilaian', 'Jawaban'];
const FEEDBACK_DAILY_LIMIT = 500; // the app is public: keeps the sheet from being flooded
const DOC_COLS = ['ID', 'Kode', 'Judul', 'Nama pendek', 'Jenis', 'Area', 'Halaman', 'Kata kunci', 'Persetujuan', 'Status', 'Diperbarui'];
const VER_COLS = ['ID dokumen', 'Rev', 'Berlaku', 'Catatan', 'Isi'];
const TYPES = ['sop', 'ik', 'spec', 'safety'];
const AREAS = ['pc', 'dmg', 'fc', 'storage', 'outbound'];
const TYPE_CODE = { sop: 'SOP', ik: 'IK', spec: 'Specbook', safety: 'SOP' };
// Fastest first. 3.1 Flash-Lite answered in ~1 s where the -latest aliases took 17-50 s; the others step in when it is busy (503).
const MODELS = ['gemini-3.1-flash-lite', 'gemini-3.5-flash', 'gemini-flash-lite-latest'];
const DEFAULT_DAILY_LIMIT = 300;
// Browser tab icon: segari.id's own icon, served by Google's favicon service (Apps Script needs an https image URL).
const FAVICON_URL = 'https://www.google.com/s2/favicons?domain=segari.id&sz=64&f=icon.png'; // must end in an image extension
const CONTEXT_CHARS = 60000;
const DOC_CONTEXT_CHARS = 16000; // on a document page: its most relevant sections, enough for a 27-page SOP's key parts
const CHUNK_TTL = 600;      // seconds the parsed library stays cached
const ANSWER_TTL = 21600;   // seconds an identical question reuses its answer (a saved document starts fresh anyway)
const SUGGEST_TTL = 21600;  // seconds a document keeps its suggested questions

// ---------------------------------------------------------------- web app

function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  t.appUrl = ScriptApp.getService().getUrl() || '';
  const page = t.evaluate()
    .setTitle('Segari Knowledge Base')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
  // The Segari mark in the browser tab; a rejected icon must never stop the page from loading.
  try { page.setFaviconUrl(FAVICON_URL); } catch (e) { console.warn('Favicon: ' + e.message); }
  return page;
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// ---------------------------------------------------------------- library

/** Published documents (revisions newest first) and the example questions, all from the spreadsheet. */
function getLibrary() {
  const ss = library_();
  const prompts = values_(ss.getSheetByName(PROMPT_SHEET), 1).map(r => String(r[0]).trim()).filter(Boolean);
  return { docs: publishedDocs_(), prompts: prompts };
}

function publishedDocs_() {
  return readLibrary_().filter(d => d.status !== 'draf');
}

/**
 * Saves a draft or publishes an article from writer mode.
 * Publishing a changed body adds a new revision whose changelog Gemini writes by comparing it with the previous one.
 * A draft (or a document never published) keeps a single revision that is replaced on each save.
 * @param {{id?:string,title:string,type:string,areas:string[],body:string,effectiveDate:string,keywords:string[]}} a
 * @param {'draft'|'published'} status
 */
function saveArticle(a, status) {
  a = clean_(a);
  if (status === 'published' && !a.title) throw new Error('Isi judul artikel dulu');
  checkBodySize_(a.body);
  library_();
  const old = a.id ? readLibrary_().find(d => d.id === a.id) : null;
  const live = old && old.status !== 'draf' && old.versions.length > 0;
  const prev = live ? old.versions[0] : null;
  const changed = !prev || prev.body.trim() !== a.body.trim();
  // Gemini runs before taking the lock so a slow answer never blocks other saves.
  const note = status === 'published' && changed ? changelog_(prev ? prev.body : '', a.body) : '';

  const ss = library_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const id = a.id || 'artikel-' + Utilities.getUuid().slice(0, 8);
    const docRow = [id, old ? old.code : TYPE_CODE[a.type], a.title, old && old.short !== old.title ? old.short : a.title, a.type, a.areas.join(', '),
      old ? old.pages : '', a.keywords.join(', '), old ? old.signoff : '', status === 'published' ? 'terbit' : 'draf', now_()];
    upsertDoc_(ss.getSheetByName(DOC_SHEET), docRow);
    const verSheet = ss.getSheetByName(VER_SHEET);
    if (live) {
      if (changed) verSheet.appendRow([id, String(Math.max.apply(null, old.versions.map(v => v.rev)) + 1), a.effectiveDate, note, a.body]);
    } else {
      const rev = old && old.versions.length ? old.versions[0].rev : 1;
      replaceVersions_(verSheet, id, [[id, rev, a.effectiveDate, note, a.body]]);
    }
    bumpLibrary_();
    return readLibrary_().find(d => d.id === id);
  } finally {
    lock.releaseLock();
  }
}

function checkBodySize_(body) {
  if (body.length > BODY_LIMIT) throw new Error('Isi artikel terlalu panjang (' + body.length + ' karakter, maksimal ' + BODY_LIMIT + '). Pecah menjadi dua dokumen.');
}

/** The rich-text editor, fetched only when Mode penulis opens so readers don't download it. Built from editor-src/. */
function getEditorBundle() {
  return HtmlService.createHtmlOutputFromFile('EditorBundle').getContent();
}

function draftSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(DRAFT_SHEET) || setupSheet_(ss, DRAFT_SHEET, DRAFT_COLS, ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']);
}

/** Autosave from Mode penulis: one row per document (or new article), overwritten on each save. */
function saveDraft(d) {
  d = d || {};
  const key = String(d.key || '');
  if (!/^[\w-]{1,80}$/.test(key)) throw new Error('Draf tidak valid');
  const a = clean_({ id: d.docId, title: d.title, type: d.type, areas: d.areas, body: d.body, effectiveDate: d.date, keywords: d.keywords });
  checkBodySize_(a.body);
  const row = [key, a.id, String(Date.now()), a.title, a.type, a.areas.join(', '), a.effectiveDate, a.keywords.join(', '), a.body];
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = draftSheet_();
    const at = values_(sh, 1).findIndex(r => String(r[0]) === key);
    if (at >= 0) { sh.getRange(at + 2, 1, 1, row.length).setValues([row]); return; }
    if (sh.getLastRow() - 1 >= DRAFT_LIMIT) throw new Error('Draf penuh');
    sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  } finally {
    lock.releaseLock();
  }
}

/** Saved drafts from the last 30 days, newest first. */
function listDrafts() {
  const cutoff = Date.now() - DRAFT_KEEP_DAYS * 86400000;
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DRAFT_SHEET);
  if (!sh) return [];
  return values_(sh, DRAFT_COLS.length).filter(r => String(r[0]) && Number(r[2]) > cutoff).map(r => ({
    key: String(r[0]), docId: String(r[1]), savedAt: Number(r[2]), title: String(r[3]), type: String(r[4]),
    areas: String(r[5]).split(',').map(x => x.trim()).filter(Boolean), date: String(r[6]),
    keywords: String(r[7]).split(',').map(x => x.trim()).filter(Boolean), body: String(r[8])
  })).sort((a, b) => b.savedAt - a.savedAt);
}

function deleteDraft(key) {
  key = String(key || '');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DRAFT_SHEET);
    if (!sh) return;
    const at = values_(sh, 1).findIndex(r => String(r[0]) === key);
    if (at >= 0) sh.deleteRow(at + 2);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Fills every empty changelog in the Versi sheet by comparing each revision with the one before it.
 * Run from the editor after adding revisions by hand. It skips revisions that already have a note and stops
 * before the 6-minute limit, so run it again until the log says nothing is left.
 */
function compileChangelogs() {
  const started = Date.now();
  const sh = library_().getSheetByName(VER_SHEET);
  const rows = values_(sh, VER_COLS.length);
  const byDoc = {};
  rows.forEach((r, i) => { const id = String(r[0]).trim(); if (id) (byDoc[id] = byDoc[id] || []).push({ i: i, rev: Number(r[1]) || 0, date: dateText_(r[2]), note: String(r[3] || '').trim(), body: String(r[4] || '') }); });
  let n = 0, left = 0;
  Object.keys(byDoc).forEach(id => {
    const vs = byDoc[id].sort((a, b) => a.rev - b.rev || (a.date > b.date ? 1 : -1));
    vs.forEach((v, k) => {
      if (v.note || !v.body.trim()) return;
      if (Date.now() - started > 4.5 * 60 * 1000) { left++; return; }
      const note = changelog_(k ? vs[k - 1].body : '', v.body);
      sh.getRange(v.i + 2, 4).setValue(note);
      SpreadsheetApp.flush();
      n++;
      console.log(id + ' Rev ' + v.rev + ':\n' + note);
    });
  });
  if (n) bumpLibrary_();
  console.log(n + ' catatan perubahan ditulis.' + (left ? ' Masih ada ' + left + ', jalankan compileChangelogs() sekali lagi.' : ' Semua sudah terisi.'));
}

/** Short changelog lines (one per line) for a new revision; a plain fallback when Gemini is unavailable. */
function changelog_(prevBody, body) {
  const fallback = prevBody ? 'Pembaruan isi dokumen.' : 'Rilis pertama.';
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) return fallback;
  let text;
  if (prevBody) {
    const diff = sectionDiff_(prevBody, body);
    if (!diff) return 'Tidak ada perubahan isi.';
    text = 'PERUBAHAN PER BAGIAN (baris "-" dihapus, baris "+" ditambah atau diubah):\n\n' + diff;
  } else {
    text = 'DOKUMEN BARU (rilis pertama):\n' + body.slice(0, CONTEXT_CHARS);
  }
  try {
    const out = gemini_(key, CHANGELOG_PROMPT, text, CHANGELOG_SCHEMA, 0.1);
    const items = (out.items || []).map(x => String(x).replace(/^[-•\s]+/, '').trim()).filter(Boolean).slice(0, 6);
    return items.length ? items.join('\n') : fallback;
  } catch (e) {
    console.error('Changelog: ' + e.message);
    return fallback;
  }
}

/**
 * Only what changed between two bodies, grouped by "# " section: added and removed lines per section.
 * Keeps the Gemini request small (a 27-page SOP sends a few hundred lines instead of two full copies).
 */
function sectionDiff_(prevBody, body) {
  const split = text => {
    const out = {}, order = [];
    let cur = '(Pembuka)';
    String(text).split('\n').forEach(raw => {
      const line = raw.trim();
      if (!line) return;
      if (/^# /.test(line)) { cur = line.slice(2).trim(); return; }
      if (!out[cur]) { out[cur] = []; order.push(cur); }
      out[cur].push(line);
    });
    return { out: out, order: order };
  };
  const a = split(prevBody), b = split(body);
  const names = b.order.concat(a.order.filter(x => b.order.indexOf(x) < 0));
  const parts = [];
  names.forEach(name => {
    const before = a.out[name] || [], after = b.out[name] || [];
    if (!before.length) { parts.push('## ' + name + ' (bagian baru)\n' + after.map(l => '+ ' + l).join('\n')); return; }
    if (!after.length) { parts.push('## ' + name + ' (bagian dihapus)'); return; }
    const gone = before.filter(l => after.indexOf(l) < 0), added = after.filter(l => before.indexOf(l) < 0);
    if (gone.length || added.length) parts.push('## ' + name + '\n' + gone.map(l => '- ' + l).concat(added.map(l => '+ ' + l)).join('\n'));
  });
  return parts.join('\n\n').slice(0, CONTEXT_CHARS);
}

const THINKING = [{ thinkingLevel: 'minimal' }, { thinkingBudget: 0 }, { thinkingLevel: 'low' }, null];

/** One generateContent call with a JSON response schema; returns the parsed JSON. Moves to the next model when one is busy. */
function gemini_(key, system, text, schema, temperature, lane) {
  const own = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL');
  const all = (own ? [own] : []).concat(MODELS.filter(m => m !== own));
  const start = (lane || 0) % all.length;
  const models = all.slice(start).concat(all.slice(0, start));
  let res;
  for (let m = 0; m < models.length; m++) {
    res = callModel_(key, models[m], system, text, schema, temperature);
    if ([404, 429, 500, 503].indexOf(res.getResponseCode()) < 0) break;
    console.warn(models[m] + ' HTTP ' + res.getResponseCode() + ', trying the next model');
  }
  if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 500));
  const body = JSON.parse(res.getContentText());
  return JSON.parse(body.candidates[0].content.parts.map(p => p.text || '').join(''));
}

/**
 * Least thinking is the biggest speed win, but models accept different settings: try them fastest first
 * and remember (per model, 6 hours) the first one the model accepts.
 */
function callModel_(key, model, system, text, schema, temperature) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent';
  const cache = CacheService.getScriptCache(), memo = 'think:' + model;
  const known = cache.get(memo);
  let i = known === null ? 0 : Number(known), res;
  for (; i < THINKING.length; i++) {
    const generation = { temperature: temperature, maxOutputTokens: 1024, responseMimeType: 'application/json', responseSchema: schema };
    if (THINKING[i]) generation.thinkingConfig = THINKING[i];
    res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': key },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        systemInstruction: { parts: [{ text: system }] },
        contents: [{ role: 'user', parts: [{ text: text }] }],
        generationConfig: generation
      })
    });
    if (res.getResponseCode() !== 400 || i === THINKING.length - 1) break;
  }
  if (res.getResponseCode() === 200 && String(i) !== known) cache.put(memo, String(i), 21600);
  return res;
}

/**
 * @param {string} question
 * @param {string=} docId  the open document, or empty for every document
 * @param {number=} lane   0 normally; 1 for the backup request the page sends when the first is slow (starts at the next model)
 */
function askMancu(question, docId, lane) {
  question = String(question || '').trim().slice(0, 500);
  docId = String(docId || '');
  if (!question) throw new Error('Pertanyaan kosong');
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('GEMINI_API_KEY');
  if (!key) return null;

  // Same question on the same page within the hour: answer instantly, without Gemini or the daily quota.
  const cache = CacheService.getScriptCache();
  const gen = props.getProperty('LIB_GEN') || '0';
  const ansKey = 'ans:' + gen + ':' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, docId + '|' + question.toLowerCase().replace(/\s+/g, ' ')));
  const hit = cache.get(ansKey);
  if (hit) return JSON.parse(hit);
  countQuestion_(props);

  // On a document page Mancu reads only that document; elsewhere it reads every published one.
  const all = cachedChunks_(cache, gen);
  const scope = docId ? all.filter(c => c.docId === docId) : all;
  if (!scope.length) throw new Error('Dokumen ini belum bisa dibaca Mancu.');
  const chunks = pickChunks_(question, scope, docId ? DOC_CONTEXT_CHARS : CONTEXT_CHARS);
  const context = chunks.map(c => '[' + c.id + '] ' + c.label + '\n' + c.text).join('\n\n---\n\n');
  const focus = docId ? 'Staf sedang membuka ' + scope[0].head + '. Jawab hanya dari dokumen ini; jika tidak ada, katakan jawabannya tidak ada di dokumen ini.\n\n' : '';
  let out;
  try {
    out = gemini_(key, SYSTEM_PROMPT, focus + 'DOKUMEN:\n\n' + context + '\n\nPERTANYAAN STAF:\n' + question, ANSWER_SCHEMA, 0.2, Number(lane) || 0);
  } catch (e) {
    // A missing "connect to an external service" authorization lands here too: run authorizeMancu() once in the editor.
    console.error('Gemini: ' + e.message);
    throw new Error(/permission|izin/i.test(e.message) ? 'Mancu belum bisa terhubung ke Gemini. Hubungi admin Knowledge Base.' : 'Mancu sedang tidak tersedia. Coba lagi sebentar.');
  }
  const byId = {};
  chunks.forEach(c => { byId[c.id] = c; });
  const seen = {};
  const result = {
    intro: String(out.intro || ''),
    items: (out.items || []).map(String).filter(Boolean).slice(0, 8),
    note: String(out.note || ''),
    sources: (out.sourceIds || []).map(id => byId[id]).filter(c => c && !seen[c.id] && (seen[c.id] = true))
      .map(c => ({ docId: c.docId, sec: c.sec, label: c.label }))
  };
  try { cache.put(ansKey, JSON.stringify(result), ANSWER_TTL); } catch (e) {}
  return result;
}

/** Three questions the open document can answer, for the empty chat. Empty list when Gemini is unavailable. */
function suggestQuestions(docId) {
  docId = String(docId || '');
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('GEMINI_API_KEY');
  if (!key) return [];
  const cache = CacheService.getScriptCache();
  const gen = props.getProperty('LIB_GEN') || '0';
  const chunks = cachedChunks_(cache, gen).filter(c => c.docId === docId);
  if (!chunks.length) return [];
  const cacheKey = 'sugg:' + gen + ':' + docId;
  const hit = cache.get(cacheKey);
  if (hit) return JSON.parse(hit);
  try {
    const text = chunks.map(c => c.label + '\n' + c.text).join('\n\n').slice(0, SUGGEST_CHARS);
    const out = gemini_(key, SUGGEST_PROMPT, text, SUGGEST_SCHEMA, 0.3);
    const list = (out.items || []).map(q => String(q).trim()).filter(Boolean).slice(0, 3);
    if (list.length) cache.put(cacheKey, JSON.stringify(list), SUGGEST_TTL);
    return list;
  } catch (e) {
    console.warn('Saran pertanyaan: ' + e.message);
    return [];
  }
}

/** Saves a 👍/👎 on a Mancu answer to the "Masukan Mancu" sheet. */
function rateAnswer(question, docId, answer, vote) {
  const label = vote === 'up' ? 'Membantu' : vote === 'down' ? 'Kurang membantu' : '';
  if (!label || !String(question || '').trim()) return;
  const cache = CacheService.getScriptCache();
  const countKey = 'votes:' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd');
  const n = Number(cache.get(countKey)) || 0;
  if (n >= FEEDBACK_DAILY_LIMIT) return;
  cache.put(countKey, String(n + 1), 90000);
  // Typed text must stay text: a leading "=" would otherwise run as a sheet formula.
  const safe = (t, max) => { t = String(t || '').slice(0, max); return /^[=+\-@\t\r]/.test(t) ? "'" + t : t; };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(FEEDBACK_SHEET) || setupSheet_(ss, FEEDBACK_SHEET, FEEDBACK_COLS, []);
  sheet.appendRow([new Date(), safe(docId, 100), safe(question, 500), label, safe(answer, 2000)]);
}

/** Called when the chat opens so the first question skips the spreadsheet read. */
function warmMancu() {
  const props = PropertiesService.getScriptProperties();
  cachedChunks_(CacheService.getScriptCache(), props.getProperty('LIB_GEN') || '0');
  return true;
}

/** Section chunks of every published document, cached so most questions skip the spreadsheet read. */
function cachedChunks_(cache, gen) {
  const key = 'chunks:' + gen;
  const hit = cache.get(key);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const chunks = chunks_(publishedDocs_()).map(c => ({ id: c.id, docId: c.docId, sec: c.sec, label: c.label, head: c.head, keywords: c.keywords, text: c.text }));
  try { cache.put(key, JSON.stringify(chunks), CHUNK_TTL); } catch (e) {} // over 100 KB: just read the sheet next time
  return chunks;
}

/** Makes the next question re-read the library (after a save or a changelog run). */
function bumpLibrary_() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('LIB_GEN', String((Number(props.getProperty('LIB_GEN')) || 0) + 1));
}

/** Run once from the Apps Script editor to grant the "connect to an external service" permission Mancu needs. */
function authorizeMancu() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models', { headers: { 'x-goog-api-key': key || '' }, muteHttpExceptions: true });
  console.log(key ? 'Gemini terhubung (HTTP ' + res.getResponseCode() + ')' : 'Izin diberikan. Tambahkan GEMINI_API_KEY di Script Properties.');
}

const SYSTEM_PROMPT = [
  'Kamu adalah Mancu, asisten SOP gudang Segari. Staf gudang (inbound, QC, TL/Captain) bertanya lewat aplikasi.',
  'Jawab HANYA dari DOKUMEN yang diberikan. Jangan mengarang aturan, angka, jam, atau nama grup.',
  'Baca semua DOKUMEN dengan teliti sebelum menyimpulkan tidak ada jawaban; istilah bisa berbeda (mis. "PO" = Purchase Order).',
  'Hanya jika benar-benar tidak ada, intro: "Saya belum menemukan jawabannya di SOP." dan satu item: "Tanyakan ke TL/Captain shift Anda."',
  'Bahasa Indonesia, kalimat pendek dan jelas, sentence case. Boleh memakai "Saya".',
  'intro: satu kalimat jawaban langsung. items: langkah atau poin penting secara berurutan, maksimal 6, tanpa nomor di depan.',
  'note: peringatan penting jika ada (mis. kondisi inbound ditolak), selain itu string kosong.',
  'sourceIds: id dokumen dalam kurung siku (mis. c3) yang benar-benar dipakai untuk menjawab.'
].join('\n');

const CHANGELOG_PROMPT = [
  'Kamu menyusun catatan perubahan (changelog) untuk dokumen SOP gudang Segari.',
  'Kamu menerima PERUBAHAN PER BAGIAN antara versi sebelumnya dan versi baru. Tulis perubahan yang benar-benar terjadi: aturan, angka, langkah, tanggung jawab, atau bagian yang ditambah, diubah, atau dihapus.',
  'Pasangan baris "-" dan "+" yang mirip berarti baris itu diubah: jelaskan nilai lama dan barunya bila penting (mis. batas tumpukan 120 cm menjadi 150 cm).',
  'Sebut nama bagian (mis. "C. Inbound Supplier FC") bila jelas. Abaikan perubahan format, ejaan, atau spasi.',
  'Untuk DOKUMEN BARU, tulis satu poin "Rilis pertama" lalu ringkas isi utamanya per bagian.',
  'Bahasa Indonesia, sentence case, setiap poin satu kalimat pendek tanpa nomor. Maksimal 6 poin, urut dari yang paling penting.'
].join('\n');

const SUGGEST_CHARS = 9000;
const SUGGEST_PROMPT = [
  'Anda membantu staf gudang Segari. Dari isi dokumen berikut, tulis tepat 3 pertanyaan singkat dalam Bahasa Indonesia yang sering ditanyakan staf saat bekerja.',
  'Setiap pertanyaan harus bisa dijawab langsung dari dokumen, spesifik pada satu hal, dan paling banyak 12 kata. Jangan menyebut nama dokumen.'
].join('\n');
const SUGGEST_SCHEMA = { type: 'OBJECT', properties: { items: { type: 'ARRAY', items: { type: 'STRING' } } }, required: ['items'] };
const CHANGELOG_SCHEMA = { type: 'OBJECT', properties: { items: { type: 'ARRAY', items: { type: 'STRING' } } }, required: ['items'] };

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
      out.push({ docId: d.id, sec: 's' + idx, label: d.code + ' · ' + c.title, head: head, keywords: d.keywords.join(' '),
        get text() { return head + ' (Rev ' + v.rev + ')\n' + c.title + '\n' + c.lines.join('\n').replace(/<\/?(?:i|u|mark)>/g, '').trim(); } });
    }
  });
  out.forEach((c, n) => { c.id = 'c' + (n + 1); });
  return out;
}

const STOP = 'apa yang harus jika dan untuk bagaimana berapa dari ada atau dengan saat tidak sudah pada ini itu bisa kalau mana kapan siapa'.split(' ');

/** Most relevant chunks first, within the context budget (smaller input, faster answer). */
function pickChunks_(question, chunks, budget) {
  const words = question.toLowerCase().split(/[^a-z0-9]+/).filter(w => w.length >= 2 && STOP.indexOf(w) < 0);
  const scored = chunks.map(c => {
    const hay = (c.label + ' ' + c.keywords + ' ' + c.text).toLowerCase();
    return { c: c, n: words.reduce((s, w) => s + (hay.indexOf(w) >= 0 ? 1 : 0), 0) };
  }).sort((a, b) => b.n - a.n);
  const picked = [];
  let size = 0;
  scored.forEach(o => {
    const len = o.c.text.length;
    if (size + len <= budget) { picked.push(o.c); size += len; }
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

/** Opens the library spreadsheet, adding missing sheets and applying Seed.js once per SEED_VERSION. */
function library_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const docs = ss.getSheetByName(DOC_SHEET) || setupSheet_(ss, DOC_SHEET, DOC_COLS, ['A', 'G', 'K']);
  const vers = ss.getSheetByName(VER_SHEET) || setupSheet_(ss, VER_SHEET, VER_COLS, ['A', 'B', 'C']);
  const prompts = ss.getSheetByName(PROMPT_SHEET) || setupSheet_(ss, PROMPT_SHEET, ['Pertanyaan'], []);
  const props = PropertiesService.getScriptProperties();
  if (Number(props.getProperty('SEED_VERSION')) < SEED_VERSION) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (Number(props.getProperty('SEED_VERSION')) < SEED_VERSION) {
        syncSeed_(docs, vers, prompts);
        props.setProperty('SEED_VERSION', String(SEED_VERSION));
      }
    } finally {
      lock.releaseLock();
    }
  }
  return ss;
}

function setupSheet_(ss, name, cols, textCols) {
  const sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
  sh.setFrozenRows(1);
  textCols.forEach(c => sh.getRange(c + ':' + c).setNumberFormat('@'));
  return sh;
}

/** Drops SEED_REMOVED_IDS, adds seed documents the sheet doesn't have yet, and fills an empty question list. Rows people added are left alone. */
function syncSeed_(docs, vers, prompts) {
  [docs, vers].forEach(sh => {
    const ids = values_(sh, 1).map(r => String(r[0]).trim());
    for (let i = ids.length - 1; i >= 0; i--) if (SEED_REMOVED_IDS.indexOf(ids[i]) >= 0) sh.deleteRow(i + 2);
  });
  const have = values_(docs, 1).map(r => String(r[0]).trim());
  const add = SEED_DOCS.filter(d => have.indexOf(d.id) < 0);
  const stamp = now_();
  if (add.length) {
    docs.getRange(docs.getLastRow() + 1, 1, add.length, DOC_COLS.length).setValues(add.map(d =>
      [d.id, d.code, d.title, d.short, d.type, d.areas.join(', '), String(d.pages || ''), d.keywords.join(', '), d.signoff, 'terbit', stamp]));
    const rows = [];
    // Gemini writes the changelog of a seeded revision that came without one.
    add.forEach(d => d.versions.slice().sort((x, y) => x.rev - y.rev).forEach((v, k, all) =>
      rows.push([d.id, String(v.rev), v.effectiveDate, v.note || changelog_(k ? all[k - 1].body : '', v.body), v.body])));
    if (rows.length) vers.getRange(vers.getLastRow() + 1, 1, rows.length, VER_COLS.length).setValues(rows);
    bumpLibrary_();
  }
  if (prompts.getLastRow() < 2) prompts.getRange(2, 1, SEED_PROMPTS.length, 1).setValues(SEED_PROMPTS.map(q => [q]));
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
  const date = str(a.effectiveDate);
  return {
    id: str(a.id),
    title: str(a.title),
    type: TYPES.indexOf(a.type) >= 0 ? a.type : 'sop',
    areas: (a.areas || []).filter(x => AREAS.indexOf(x) >= 0),
    body: String(a.body || ''),
    effectiveDate: /^\d{4}-\d{2}-\d{2}$/.test(date) ? date : Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'),
    keywords: (a.keywords || []).map(str).filter(Boolean)
  };
}
