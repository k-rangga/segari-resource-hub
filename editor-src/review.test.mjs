// Run: node review.test.mjs
// Runs Code.js and Review.js against a simulated spreadsheet and a fake Gemini, so no real sheet or quota is used.
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

class Range {
  constructor(sheet, r, c, h = 1, w = 1) { Object.assign(this, { sheet, r, c, h, w }); }
  getValues() { return Array.from({ length: this.h }, (_, i) => Array.from({ length: this.w }, (_, j) => (this.sheet.cells[this.r - 1 + i] || [])[this.c - 1 + j] ?? '')); }
  getValue() { return this.getValues()[0][0]; }
  setValues(v) { v.forEach((row, i) => row.forEach((x, j) => { const y = this.r - 1 + i; (this.sheet.cells[y] = this.sheet.cells[y] || [])[this.c - 1 + j] = x; })); return this; }
  setValue(x) { return this.setValues([[x]]); }
  setFontWeight() { return this; } setNote() { return this; } setDataValidation() { return this; } setNumberFormat() { return this; }
}
class Sheet {
  constructor(name) { this.name = name; this.cells = []; }
  getLastRow() { return this.cells.length; } getMaxRows() { return 1000; }
  getRange(r, c, h, w) { if (typeof r === 'string') return new Range(this, 1, 1); return new Range(this, r, c, h, w); }
  appendRow(row) { this.cells.push([...row]); } setFrozenRows() {}
}
const book = { sheets: {}, getSheetByName(n) { return this.sheets[n] || null; }, insertSheet(n) { return (this.sheets[n] = new Sheet(n)); } };
const store = new Map();
const ctx = vm.createContext({
  console: { log() {}, warn() {}, error(m) { console.log('server error:', m); } },
  SpreadsheetApp: { openById: () => book, flush() {}, newDataValidation: () => { const b = { requireValueInList: () => b, setAllowInvalid: () => b, build: () => ({}) }; return b; } },
  CacheService: { getScriptCache: () => ({ get: k => store.get(k) ?? null, put: (k, v) => store.set(k, v), remove: k => store.delete(k) }) },
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => ({ GEMINI_API_KEY: 'test', LIB_GEN: '1' })[k] ?? null, setProperty() {} }) },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  Utilities: { formatDate: () => '2026-11-04 10:00', getUuid: () => 'x', computeDigest: (a, s) => s, base64EncodeWebSafe: s => s, DigestAlgorithm: { MD5: 'md5' } },
  HtmlService: {}, ScriptApp: {}, UrlFetchApp: {}
});
const root = new URL('../', import.meta.url);
for (const f of ['Code.js', 'Review.js']) vm.runInContext(readFileSync(new URL(f, root), 'utf8'), ctx, { filename: f });

const chunks = [
  { id: 'c1', docId: 'd1', sec: 's1', label: 'SOP-1 · A. DMT', head: 'SOP-1', keywords: '', text: 'SOP-1\nDMT memakai scan barcode' },
  { id: 'c2', docId: 'd1', sec: 's2', label: 'SOP-1 · B. DMG', head: 'SOP-1', keywords: '', text: 'SOP-1\nDMG memakai keranjang' }
];
ctx.cachedChunks_ = () => chunks;
let geminiCalls = [], reply = null;
ctx.gemini_ = (key, system, text, schema) => { geminiCalls.push({ system, text, schema }); return reply(text); };

let failed = 0;
const check = (name, ok, detail) => { if (!ok) { failed++; console.log('FAIL', name, detail === undefined ? '' : JSON.stringify(detail)); } else console.log('ok  ', name); };

// ---- review
reply = () => { throw new Error('should not call Gemini'); };
check('no feedback sheet is fine', ctx.reviewFeedback() === 'Belum ada masukan.');

ctx.rateAnswer('=SUM(A1) apa itu DMT?', 'd1', 'Saya belum menemukan', 'down');
ctx.rateAnswer('cara scan PO', '', 'Scan barcode', 'up');
const fb = book.sheets['Masukan Mancu'];
check('feedback formula is stored as text', String(fb.cells[1][2]).startsWith("'="), fb.cells[1][2]);

reply = () => ({
  summary: 'Staf bertanya istilah.',
  reviews: [
    { no: 1, kategori: 'Istilah belum dikenal', analisis: '=bahaya', saran: 'Tambahkan penjelasan DMT' },
    { no: 2, kategori: 'Jawaban sudah tepat', analisis: 'Sudah benar', saran: '' }
  ],
  topics: [{ no: 1, topik: 'DMT', istilah: 'DMT, dry goods', maksud: 'Area Processing Center', dokumen: '' }]
});
const msg = ctx.reviewFeedback();
check('review reports its work', msg === '2 masukan ditinjau. Semua sudah ditinjau.', msg);
const an = book.sheets['Analisis Masukan'];
check('analysis has a row per feedback', an.cells.length === 3, an.cells.length);
check('analysis keeps category and advice', an.cells[1][5] === 'Istilah belum dikenal' && an.cells[1][7] === 'Tambahkan penjelasan DMT');
check('Gemini text cannot become a formula', String(an.cells[1][6]).startsWith("'="), an.cells[1][6]);
check('input includes the SOP sections and stripped question', geminiCalls.at(-1).text.includes('DMT memakai scan barcode') && geminiCalls.at(-1).text.includes('Pertanyaan: =SUM'));
const rs = book.sheets['Ringkasan Masukan'];
check('summary counts votes', rs.cells[1][2] === 2 && rs.cells[1][3] === 1 && rs.cells[1][4] === 1, rs.cells[1]);
check('feedback marked reviewed', fb.cells[1][5] && fb.cells[2][5]);

const before = geminiCalls.length;
check('second run has nothing to do', ctx.reviewFeedback() === 'Tidak ada masukan baru.' && geminiCalls.length === before);
check('busy flag released', !store.has('review:busy'));

// ---- repository
const kb = book.sheets['Konteks Mancu'];
check('new topic starts as Draf', kb.cells[1][0] === 'DMT' && kb.cells[1][5] === 'Draf', kb.cells[1]);
store.delete('ctxnotes');
check('Draf is not read by Mancu', ctx.contextNotes_(ctx.CacheService.getScriptCache()).length === 0);

kb.cells[1][3] = 'DMT adalah alur inbound di Processing Center.';
kb.cells[1][5] = 'Aktif';
ctx.rateAnswer('DMT itu apa', '', 'x', 'down');
reply = () => ({ summary: 's', reviews: [{ no: 1, kategori: 'Lainnya', analisis: 'a', saran: 'b' }], topics: [{ no: 1, topik: 'dmt', maksud: 'tebakan baru Gemini' }] });
ctx.reviewFeedback();
check('review never overwrites the admin', kb.cells[1][3].startsWith('DMT adalah') && kb.cells[1][2] === 'Area Processing Center' && kb.cells[1][5] === 'Aktif', kb.cells[1]);
check('known topic only raises its count', kb.cells.length === 2 && kb.cells[1][6] === 2, kb.cells[1]);

// failed Gemini leaves rows for next time
ctx.rateAnswer('pertanyaan gagal', '', 'x', 'down');
reply = () => { throw new Error('503'); };
check('Gemini failure reports nothing done', ctx.reviewFeedback() === 'Tidak ada baris yang berhasil ditinjau.');
check('failed row stays unreviewed', !fb.cells.at(-1)[5]);

// ---- Mancu reads Aktif notes
store.delete('ctxnotes');
const cache = ctx.CacheService.getScriptCache();
const notes = ctx.contextNotes_(cache);
check('Aktif note is read, admin text preferred', notes.length === 1 && notes[0].text.startsWith('DMT adalah') && notes[0].id === 'k2', notes);
check('note matches a whole short word', ctx.pickNotes_(notes, 'apa itu DMT?', '').length === 1);
check('unrelated question gets no note', ctx.pickNotes_(notes, 'cara scan barcode', '').length === 0);
kb.cells[1][4] = 'd2';
store.delete('ctxnotes');
const scoped = ctx.contextNotes_(cache);
check('note for another document stays out', ctx.pickNotes_(scoped, 'apa itu DMT', 'd1').length === 0 && ctx.pickNotes_(scoped, 'apa itu DMT', 'd2').length === 1);
kb.cells[1][4] = '';
store.delete('ctxnotes');

reply = text => ({ intro: 'DMT adalah alur inbound.', items: ['satu'], note: '', sourceIds: ['k2', 'c1'] });
let a = ctx.askMancu('apa itu DMT', '', 0, []);
const prompt = geminiCalls.at(-1).text;
check('note reaches Gemini', prompt.includes('CATATAN ADMIN') && prompt.includes('[k2] DMT'));
check('note appears as a source', a.sources.some(s => s.label === 'Catatan admin · DMT' && !s.docId) && a.sources.some(s => s.docId === 'd1'), a.sources);
check('no table on a normal answer', !a.table);

// an edited note must not serve the old saved answer
const calls = geminiCalls.length;
ctx.askMancu('apa itu DMT', '', 0, []);
check('same question reuses the saved answer', geminiCalls.length === calls);
kb.cells[1][3] = 'Penjelasan baru dari admin.';
store.delete('ctxnotes');
ctx.askMancu('apa itu DMT', '', 0, []);
check('edited note gets a fresh answer', geminiCalls.length === calls + 1);

// ---- comparison table
reply = () => ({ intro: 'Ada perbedaan.', items: ['dibuang'], note: '', sourceIds: ['c1', 'c2'], table: { headers: ['Aspek', 'DMT', 'DMG'], rows: [['Kedatangan', 'Scan barcode', 'Keranjang'], ['Kurang kolom', 'x'], ['Terlalu', 'banyak', 'kolom', 'buang']] } });
a = ctx.askMancu('perbedaan DMG dan DMT', '', 0, []);
check('table is kept', a.table && a.table.headers.length === 3 && a.table.rows[0][1] === 'Scan barcode', a.table);
check('short rows filled, long rows cut', a.table.rows[1].length === 3 && a.table.rows[1][2] === 'Tidak disebutkan' && a.table.rows[2].length === 3, a.table.rows);
check('table replaces the list', a.items.length === 0);
reply = () => ({ intro: 'Jawaban.', items: ['a'], note: '', sourceIds: ['c1'], table: { headers: [], rows: [] } });
a = ctx.askMancu('pertanyaan biasa saja', '', 0, []);
check('empty table is dropped', !a.table && a.items.length === 1);
check('history is limited', ctx.historyTurns_([1, 2, 3, 4, 5].map(i => ({ q: 'q' + i, a: 'x'.repeat(5000) }))).length === 3);

console.log(failed ? failed + ' FAILED' : 'all passed');
process.exit(failed ? 1 : 0);
