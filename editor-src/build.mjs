// Builds ../EditorBundle.html. Apps Script's HTML parser alters raw JavaScript, so the file holds it Base64-encoded.
import { execFileSync } from 'node:child_process';
import { readFileSync, writeFileSync, rmSync } from 'node:fs';
const tmp = new URL('./.bundle.js', import.meta.url).pathname;
execFileSync('npx', ['esbuild', 'entry.js', '--bundle', '--minify', '--format=iife', '--target=es2019', '--outfile=' + tmp, '--log-level=warning'], { stdio: 'inherit', cwd: new URL('.', import.meta.url).pathname });
const b64 = readFileSync(tmp).toString('base64');
rmSync(tmp);
writeFileSync(new URL('../EditorBundle.html', import.meta.url), '<script type="text/plain">' + b64 + '</script>\n');
console.log('EditorBundle.html: ' + Math.round(b64.length / 1024) + ' KB');
