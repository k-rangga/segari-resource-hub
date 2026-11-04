// Run after every deploy: npm run check-live
// Apps Script strips comments from served code with a simple scanner that can misread slashes and quotes,
// which can break or silently change the app. This fetches the live page and proves its code still matches App.html.
import { readFileSync } from 'node:fs';
import { transform } from 'esbuild';
const URL_LIVE = 'https://script.google.com/macros/s/AKfycbzjmWxkCzkhmCLPFoEYvah-zCFz_W6ypLsZ2N69LdQ_B2pdyY0A6KSZ5etHO-avTH1CdA/exec';
const page = await (await fetch(URL_LIVE, { redirect: 'follow' })).text();
const start = page.indexOf('goog.script.init(') + 17, end = page.indexOf('", "", undefined', start);
const cfg = JSON.parse((0, eval)(page.slice(start, end + 1)));
const scripts = [...cfg.userHtml.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m => m[1]);
const served = scripts.reduce((a, b) => a.length > b.length ? a : b);
const local = readFileSync(new URL('../App.html', import.meta.url), 'utf8').match(/<script>([\s\S]*)<\/script>/)[1];
const min = code => transform(code, { minify: true }).then(r => r.code);
const same = (await min(served)) === (await min(local));
console.log(same ? 'OK: live code matches App.html' : 'MISMATCH: live code differs from App.html. Check for a slash or quote that confuses the Apps Script comment stripper, or wait a minute if you just deployed.');
process.exit(same ? 0 : 1);
