/**
 * Segari Knowledge Base — server side.
 * Serves the web app and stores articles written in "Mode penulis" in a spreadsheet.
 * The spreadsheet is created on the first save; its ID is kept in Script Properties.
 */

const TZ = 'Asia/Jakarta';
const SS_PROP = 'ARTICLE_SPREADSHEET_ID';
const SHEET = 'Artikel';
const COLS = ['ID', 'Status', 'Judul', 'Jenis', 'Area', 'Isi', 'Versi', 'Kata kunci', 'Rev terbaru', 'Berlaku', 'Diperbarui', 'Oleh'];
const TYPES = ['sop', 'ik', 'spec', 'safety'];
const AREAS = ['pc', 'dmg', 'fc', 'storage', 'outbound'];

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

// ---------------------------------------------------------------- articles

/** Published articles, newest first. */
function getArticles() {
  const sh = sheet_(false);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, COLS.length).getValues()
    .filter(r => r[0] && r[1] === 'published')
    .map(rowToArticle_)
    .reverse();
}

/**
 * Saves a draft or publishes an article.
 * @param {{id?:string,title:string,type:string,areas:string[],body:string,versions:{rev:number,effectiveDate:string,note:string}[],keywords:string[]}} a
 * @param {'draft'|'published'} status
 */
function saveArticle(a, status) {
  a = clean_(a);
  if (status === 'published' && !a.title) throw new Error('Isi judul artikel dulu');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_(true);
    const id = a.id || Utilities.getUuid().slice(0, 8);
    const latest = a.versions[a.versions.length - 1];
    const row = [id, status === 'published' ? 'published' : 'draft', a.title, a.type, a.areas.join(', '), a.body,
      JSON.stringify(a.versions), a.keywords.join(', '), latest.rev, latest.effectiveDate,
      Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), '–'];
    const n = sh.getLastRow();
    const ids = n > 1 ? sh.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0])) : [];
    const i = ids.indexOf(id);
    if (i >= 0) sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return rowToArticle_(row);
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- helpers

function sheet_(create) {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(SS_PROP);
  if (!id) {
    if (!create) return null;
    const ss = SpreadsheetApp.create('Segari Knowledge Base – Artikel');
    ss.getSheets()[0].setName(SHEET).appendRow(COLS);
    ss.getSheets()[0].setFrozenRows(1);
    id = ss.getId();
    props.setProperty(SS_PROP, id);
  }
  return SpreadsheetApp.openById(id).getSheetByName(SHEET);
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

function rowToArticle_(r) {
  let versions = [];
  try { versions = JSON.parse(r[6] || '[]'); } catch (e) {}
  return {
    id: String(r[0]), status: r[1], title: r[2], type: r[3],
    areas: String(r[4]).split(',').map(s => s.trim()).filter(Boolean),
    body: r[5], versions: versions,
    keywords: String(r[7]).split(',').map(s => s.trim()).filter(Boolean)
  };
}
