// The Mode penulis editor. Built into ../EditorBundle.html and fetched only when the writer opens, so readers never load it.
import { Editor, Extension } from '@tiptap/core';
import { Node as PMNode, Slice } from '@tiptap/pm/model';
import { NodeSelection } from '@tiptap/pm/state';
import Document from '@tiptap/extension-document';
import Paragraph from '@tiptap/extension-paragraph';
import Text from '@tiptap/extension-text';
import Heading from '@tiptap/extension-heading';
import Bold from '@tiptap/extension-bold';
import Italic from '@tiptap/extension-italic';
import Underline from '@tiptap/extension-underline';
import Highlight from '@tiptap/extension-highlight';
import Link from '@tiptap/extension-link';
import HardBreak from '@tiptap/extension-hard-break';
import { BulletList, OrderedList, ListItem } from '@tiptap/extension-list';
import { Table, TableRow, TableCell, TableHeader } from '@tiptap/extension-table';
import { UndoRedo, Placeholder, Dropcursor, Gapcursor } from '@tiptap/extensions';
import { bodyToDoc, docToBody } from './convert.js';

// The stored format has no nested lists and one line per item or cell, so the editor only allows what can be saved.
const FlatItem = ListItem.extend({
  content: 'paragraph',
  addKeyboardShortcuts() { return { Enter: () => this.editor.commands.splitListItem(this.name) }; }
});
const OneLineCell = TableCell.extend({ content: 'paragraph' });
const OneLineHeader = TableHeader.extend({ content: 'paragraph' });
// Pasted h3-h6 become sub-headings rather than losing their heading.
const Headings = Heading.extend({
  parseHTML() {
    return [{ tag: 'h1', attrs: { level: 1 } }, ...['h2', 'h3', 'h4', 'h5', 'h6'].map(tag => ({ tag, attrs: { level: 2 } }))];
  }
});

/** Word marks lists as styled paragraphs: rebuild real lists and drop its typed bullets and numbers. */
export function cleanWordHtml(html) {
  if (!/class="?Mso|mso-list|urn:schemas-microsoft/i.test(html)) return html;
  const doc = new DOMParser().parseFromString(html.replace(/<!\[if !supportLists\]>|<!\[endif\]>|<!--\[if !supportLists\]-->|<!--\[endif\]-->/g, ''), 'text/html');
  let list = null;
  [...doc.body.querySelectorAll('p, h1, h2, h3, h4, h5, h6, table, div')].forEach(el => {
    const isItem = el.tagName === 'P' && /mso-list:/i.test(el.getAttribute('style') || '');
    if (!isItem) { list = null; return; }
    const marker = [...el.querySelectorAll('span')].find(s => /mso-list:\s*ignore/i.test(s.getAttribute('style') || ''));
    const label = marker ? marker.textContent.trim() : '';
    if (marker) marker.remove();
    const kind = /^\(?\d+[.)]?$|^[a-zA-Z][.)]$/.test(label) ? 'OL' : 'UL';
    if (!list || list.tagName !== kind || list.nextElementSibling !== el) {
      list = doc.createElement(kind);
      el.before(list);
    }
    const li = doc.createElement('li');
    li.innerHTML = el.innerHTML;
    list.append(li);
    el.remove();
  });
  return doc.body.innerHTML;
}

/** Top-level block actions behind the hover handle. `pos` is the position just before the block. */
const blocks = {
  /** Position of the top-level block rendered by this element. */
  posOf(editor, el) {
    const $p = editor.state.doc.resolve(editor.view.posAtDOM(el, 0));
    return $p.depth ? $p.before(1) : $p.pos;
  },
  /** Hands a block to the editor's own drag and drop, which shows the drop line and moves it as one undo step. */
  startDrag(editor, pos, event, el) {
    const { view } = editor;
    const sel = NodeSelection.create(view.state.doc, pos);
    view.dispatch(view.state.tr.setSelection(sel));
    view.dragging = { slice: sel.content(), move: true };
    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/plain', sel.node.textContent);
    if (el) event.dataTransfer.setDragImage(el, 0, 0);
  },
  /** A new line under the block with the slash menu already open. */
  addBelow(editor, pos) {
    const node = editor.state.doc.nodeAt(pos);
    if (node.type.name === 'paragraph' && !node.content.size) {
      editor.chain().focus().insertContentAt(pos + 1, '/').run();
      return;
    }
    const end = pos + node.nodeSize;
    editor.chain().focus().insertContentAt(end, { type: 'paragraph', content: [{ type: 'text', text: '/' }] }).setTextSelection(end + 2).run();
  },
  duplicate(editor, pos) {
    const node = editor.state.doc.nodeAt(pos);
    editor.view.dispatch(editor.state.tr.insert(pos + node.nodeSize, node));
    editor.commands.focus();
  },
  remove(editor, pos) {
    const node = editor.state.doc.nodeAt(pos);
    editor.view.dispatch(editor.state.tr.delete(pos, pos + node.nodeSize));
    editor.commands.focus();
  },
  /** Selects the block's text so a style can be applied to all of it. */
  select(editor, pos) {
    const node = editor.state.doc.nodeAt(pos);
    editor.commands.setTextSelection({ from: pos + 1, to: pos + node.nodeSize - 1 });
  }
};

function create({ element, body, onChange, onState, onNotice, onLinkKey, onKeyDown }) {
  const Keys = Extension.create({
    name: 'kbKeys',
    addKeyboardShortcuts() { return { 'Mod-k': () => { onLinkKey && onLinkKey(); return true; } }; }
  });
  const editor = new Editor({
    element,
    content: bodyToDoc(body),
    extensions: [
      Document, Paragraph, Text, HardBreak, UndoRedo.configure({ depth: 200 }), Dropcursor, Gapcursor,
      Headings.configure({ levels: [1, 2] }), Bold, Italic, Underline, Highlight,
      Link.configure({ openOnClick: false, autolink: true, linkOnPaste: true, defaultProtocol: 'https', protocols: ['http', 'https'], HTMLAttributes: { class: 'doc-link', rel: 'noopener noreferrer', target: '_blank' } }),
      BulletList, OrderedList, FlatItem,
      Table.configure({ resizable: false }), TableRow, OneLineCell, OneLineHeader,
      Placeholder.configure({
        placeholder: ({ editor: ed, node }) => node.type.name === 'heading' ? (node.attrs.level === 1 ? 'Judul bagian' : 'Subjudul')
          : ed.isEmpty ? 'Tulis prosedur, atau ketik / untuk menu' : 'Ketik / untuk menu'
      }),
      Keys
    ],
    editorProps: {
      handleKeyDown: (view, event) => !!(onKeyDown && onKeyDown(event)),
      attributes: { class: 'pm', role: 'textbox', 'aria-multiline': 'true', 'aria-label': 'Isi artikel', spellcheck: 'false' },
      transformPastedHTML(html) {
        if (/<img[\s>]/i.test(html)) onNotice && onNotice('Gambar belum didukung, jadi tidak ikut ditempel.');
        return cleanWordHtml(html);
      },
      // Plain text that uses the article markup (old drafts, notes) pastes as real headings, lists and tables.
      clipboardTextParser(text, $ctx, plain, view) {
        if (!/\n|^(#{1,2} |\d+[.)] |[-•] |\|)/.test(text)) return null;
        const doc = PMNode.fromJSON(view.state.schema, bodyToDoc(text.replace(/\r\n?/g, '\n')));
        return Slice.maxOpen(doc.content);
      },
      handleDrop(view, event) {
        if (event.dataTransfer && [...event.dataTransfer.files].some(f => f.type.startsWith('image/'))) { onNotice && onNotice('Gambar belum didukung.'); return true; }
        return false;
      },
      handlePaste(view, event) {
        const items = event.clipboardData ? [...event.clipboardData.items] : [];
        if (items.length && items.every(i => i.type.startsWith('image/'))) { onNotice && onNotice('Gambar belum didukung.'); return true; }
        return false;
      }
    }
  });

  // Opening a document and undoing back to the start must not count as a change, or a clean revision would look edited.
  const getBody = () => docToBody(editor.getJSON());
  const clean = getBody();
  editor.on('update', () => { const b = getBody(); onChange(b === clean ? body : b); });
  ['transaction', 'focus', 'blur'].forEach(name => editor.on(name, () => onState && onState()));
  return editor;
}

window.KBEditor = { create, blocks, bodyToDoc, docToBody, cleanWordHtml };
