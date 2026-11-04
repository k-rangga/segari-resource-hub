/**
 * Article body text <-> editor document.
 * The body text stays the stored format, so the reader, Mancu, search and the changelog keep working unchanged:
 *   "# " section, "## " sub-heading, "1. " steps, "- " bullets, "| a | b |" tables (first row is the header),
 *   **bold**, <i>italic</i>, <u>underline</u>, <mark>highlight</mark>, [text](https://url).
 * Keep the block rules in step with parseBody() in App.html.
 */

const TOKEN = /(\*\*[^*]+\*\*|<i>.*?<\/i>|<u>.*?<\/u>|<mark>.*?<\/mark>|\[[^\]]+\]\(https?:\/\/[^\s)]+\))/g;
const WRAP = [['<mark>', '</mark>', 'highlight'], ['<u>', '</u>', 'underline'], ['<i>', '</i>', 'italic']];

function parseInline(text, marks = []) {
  const out = [];
  String(text).split(TOKEN).forEach(part => {
    if (!part) return;
    if (/^\*\*[^*]+\*\*$/.test(part)) return void out.push(...parseInline(part.slice(2, -2), [...marks, { type: 'bold' }]));
    for (const [open, close, type] of WRAP) {
      if (part.startsWith(open) && part.endsWith(close)) return void out.push(...parseInline(part.slice(open.length, -close.length), [...marks, { type }]));
    }
    const link = /^\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)$/.exec(part);
    if (link) return void out.push(...parseInline(link[1], [...marks, { type: 'link', attrs: { href: link[2] } }]));
    out.push(marks.length ? { type: 'text', text: part, marks } : { type: 'text', text: part });
  });
  return out;
}

/** A paragraph may span lines; each line break in the text is a hard break in the editor. */
function inlineWithBreaks(text) {
  const out = [];
  String(text).split('\n').forEach((line, i) => {
    if (i) out.push({ type: 'hardBreak' });
    out.push(...parseInline(line));
  });
  return out;
}

const para = text => { const content = parseInline(text); return content.length ? { type: 'paragraph', content } : { type: 'paragraph' }; };
const splitCells = line => line.slice(1, -1).split(/(?<!\\)\|/).map(c => c.trim().replace(/\\\|/g, '|'));

export function bodyToDoc(body) {
  const blocks = [];
  let lines = [];
  const flush = () => {
    if (lines.length) { blocks.push({ type: 'paragraph', content: inlineWithBreaks(lines.join('\n')) }); lines = []; }
  };
  const last = () => blocks[blocks.length - 1];
  String(body || '').split('\n').forEach(raw => {
    const line = raw.replace(/\s+$/, '');
    let m;
    if (!line.trim()) return flush();
    if (/^# /.test(line)) {
      flush();
      const title = line.slice(2).trim();
      return void blocks.push(title ? { type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: title }] } : { type: 'heading', attrs: { level: 1 } });
    }
    if (/^## /.test(line)) {
      flush();
      return void blocks.push({ type: 'heading', attrs: { level: 2 }, content: parseInline(line.slice(3).trim()) });
    }
    if ((m = /^(\d+)[.)]\s+(.*)$/.exec(line))) {
      flush();
      const n = Number(m[1]), b = last(), item = { type: 'listItem', content: [para(m[2])] };
      // A number that doesn't follow on starts a new list, so every step keeps the number it was written with.
      if (b && b.type === 'orderedList' && b.attrs.start + b.content.length === n) b.content.push(item);
      else blocks.push({ type: 'orderedList', attrs: { start: n }, content: [item] });
      return;
    }
    if ((m = /^[-•]\s+(.*)$/.exec(line))) {
      flush();
      const b = last(), item = { type: 'listItem', content: [para(m[1])] };
      if (b && b.type === 'bulletList') b.content.push(item);
      else blocks.push({ type: 'bulletList', content: [item] });
      return;
    }
    if (/^\|.*\|$/.test(line)) {
      flush();
      const cells = splitCells(line);
      if (cells.every(c => /^:?-{2,}:?$/.test(c))) return;
      const b = last();
      const row = kind => ({ type: 'tableRow', content: cells.map(c => ({ type: kind, content: [para(c)] })) });
      if (b && b.type === 'table') b.content.push(row('tableCell'));
      else blocks.push({ type: 'table', content: [row('tableHeader')] });
      return;
    }
    lines.push(line);
  });
  flush();
  // The editor needs every table row to have the same number of cells.
  blocks.filter(b => b.type === 'table').forEach(t => {
    const width = Math.max(...t.content.map(r => r.content.length));
    t.content.forEach((r, i) => { while (r.content.length < width) r.content.push({ type: i ? 'tableCell' : 'tableHeader', content: [para('')] }); });
  });
  return { type: 'doc', content: blocks.length ? blocks : [{ type: 'paragraph' }] };
}

function inlineText(nodes, breakAs) {
  return (nodes || []).map(n => {
    if (n.type === 'hardBreak') return breakAs;
    if (n.type !== 'text') return '';
    const marks = n.marks || [], has = t => marks.find(m => m.type === t);
    let s = n.text;
    const link = has('link');
    // Only web links render in the reader; anything else stays plain text.
    if (link && /^https?:\/\/\S+$/.test(link.attrs.href)) s = '[' + s.replace(/\[/g, '(').replace(/\]/g, ')') + '](' + link.attrs.href.replace(/\)/g, '%29') + ')';
    if (has('bold')) s = '**' + s + '**';
    if (has('italic')) s = '<i>' + s + '</i>';
    if (has('underline')) s = '<u>' + s + '</u>';
    if (has('highlight')) s = '<mark>' + s + '</mark>';
    return s;
  }).join('');
}

const plainText = nodes => (nodes || []).map(n => n.type === 'text' ? n.text : ' ').join('').trim();
const oneLine = node => inlineText((node.content || []).flatMap(c => c.content || []), ' ').replace(/\s+$/, '');

export function docToBody(doc) {
  const out = [];
  (doc.content || []).forEach(b => {
    if (b.type === 'heading') {
      // Section titles are plain text in the reader; sub-headings keep their formatting.
      const text = b.attrs.level === 1 ? plainText(b.content) : inlineText(b.content, ' ').trim();
      if (text) out.push((b.attrs.level === 1 ? '# ' : '## ') + text);
    } else if (b.type === 'paragraph') {
      const text = inlineText(b.content, '\n').split('\n').map(l => l.replace(/\s+$/, '')).join('\n').replace(/^\n+|\n+$/g, '');
      if (text) out.push(text.replace(/\n{2,}/g, '\n'));
    } else if (b.type === 'bulletList' || b.type === 'orderedList') {
      const start = (b.attrs && b.attrs.start) || 1;
      const items = (b.content || []).map(oneLine).filter(t => t.trim());
      if (items.length) out.push(items.map((t, i) => (b.type === 'bulletList' ? '- ' : (start + i) + '. ') + t).join('\n'));
    } else if (b.type === 'table') {
      const rows = (b.content || []).map(r => (r.content || []).flatMap(c => {
        const text = oneLine(c).trim().replace(/\|/g, '\\|');
        // Merged cells from pasted tables: keep the columns lined up.
        return [text, ...Array(Math.max(0, ((c.attrs && c.attrs.colspan) || 1) - 1)).fill('')];
      }));
      const width = Math.max(0, ...rows.map(r => r.length));
      if (width) out.push(rows.map(r => '| ' + [...r, ...Array(width - r.length).fill('')].join(' | ') + ' |').join('\n'));
    }
  });
  return out.join('\n\n') + (out.length ? '\n' : '');
}
