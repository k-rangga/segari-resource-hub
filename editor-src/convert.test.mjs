// Run: node convert.test.mjs [seed.json]
// Every document must survive body -> editor -> body with the same meaning and the same non-blank lines,
// and editing must never add noise to the Gemini changelog.
import { readFileSync } from 'node:fs';
import { bodyToDoc, docToBody } from './convert.js';

// The reader's own parser, so the test compares what readers would actually see.
const app = readFileSync(new URL('../App.html', import.meta.url), 'utf8');
const start = app.indexOf('function parseBody('), end = app.indexOf('/** "**bold**"', start);
const parseBody = new Function(app.slice(start, end) + '; return parseBody;')();

const round = b => docToBody(bodyToDoc(b));
const shape = b => JSON.stringify(parseBody(b, 'T'));
const lines = b => b.split('\n').map(l => l.trim()).filter(Boolean);

let failed = 0;
const check = (name, ok, detail) => { if (!ok) { failed++; console.log('FAIL', name, detail || ''); } else console.log('ok  ', name); };

const cases = {
  'sections and text': '# A\nTeks\n\nParagraf dua\nbaris dua\n\n# B\n',
  'sub-heading with bold': '# A\n\n## **Sub** judul\n',
  'steps keep their numbers': '1. a\n2. b\n\n5. c\n1. d\n',
  'bullets': '- a\n- b **tebal**\n',
  'link inside bold': '- **[teks](https://a.id/x?y=1#z)** lalu\n',
  'all marks': 'a **b** <i>c</i> <u>d</u> <mark>e</mark> [f](https://g.id)\n',
  'nested marks': '<mark><u><i>**[a](https://b.id)**</i></u></mark>\n',
  'table with pipe': '| A | B |\n| x \\| y | **z** |\n',
  'empty table cell': '| A | B |\n|  | y |\n',
  'stars stay text': 'Lantai* dan plafon*\n\n*) Berlaku semua\n',
  'text before first section': 'Pembuka\n\n# A\n',
};
for (const [name, b] of Object.entries(cases)) {
  const r = round(b);
  check(name, shape(r) === shape(b) && round(r) === r, '\n  in : ' + JSON.stringify(b) + '\n  out: ' + JSON.stringify(r));
}
// Text that isn't in stored form still means the same thing afterwards.
for (const [name, b] of Object.entries({ 'separator row dropped': '| A | B |\n| --- | --- |\n| x | y |\n', 'bullet dot': '• a\n• b\n', 'uneven table': '| A | B |\n| x |\n' })) {
  const r = round(b);
  check(name, round(r) === r && !/---|•/.test(r), JSON.stringify(r));
}

// Pass an exported library as JSON to test live data; by default this checks the documents in Seed.js.
const seedSrc = readFileSync(new URL('../Seed.js', import.meta.url), 'utf8');
const seed = process.argv[2] ? JSON.parse(readFileSync(process.argv[2], 'utf8')) : { docs: new Function(seedSrc + '; return SEED_DOCS;')() };
for (const d of seed.docs) for (const v of d.versions) {
  const r = round(v.body), tag = d.id + ' Rev ' + v.rev;
  check(tag + ' same meaning', shape(r) === shape(v.body));
  check(tag + ' stable', round(r) === r);
  const a = lines(v.body), b = lines(r), gone = a.filter(l => !b.includes(l)), added = b.filter(l => !a.includes(l));
  check(tag + ' no changelog noise (' + a.length + ' lines)', !gone.length && !added.length, '\n  - ' + gone.slice(0, 4).join('\n  - ') + '\n  + ' + added.slice(0, 4).join('\n  + '));
}
console.log(failed ? failed + ' FAILED' : 'all passed');
process.exit(failed ? 1 : 0);
