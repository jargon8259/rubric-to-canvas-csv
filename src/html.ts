import type { Block, ParsedDoc } from './types.ts';
import { decodeEntities, cleanCell, squash } from './text.ts';

const BLOCK_TAGS = new Set(['p', 'div', 'li', 'ul', 'ol', 'tr', 'section', 'article', 'blockquote', 'pre']);

interface TableState { rows: string[][]; row: string[] | null; cell: string | null; span: number; merged: boolean; nested: boolean }

/** Reads pasted HTML (a table from Word, Excel or a web page, or a whole document) into headings, text and tables. */
export function htmlToBlocks(html: string): Pick<ParsedDoc, 'blocks'> {
  let src = html;
  const body = /<body[^>]*>([\s\S]*)/i.exec(src);
  if (body) src = body[1];
  src = src
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<!\[[^\]]*\]>/g, '')
    .replace(/<(style|script|head|title)\b[\s\S]*?<\/\1\s*>/gi, '')
    .replace(/<\?[\s\S]*?\?>/g, '');

  const blocks: Block[] = [];
  const tables: TableState[] = []; // only the outermost table is read
  let depth = 0;
  let line = '';
  let headingLevel = 0;

  const flushLine = () => {
    const text = squash(line);
    line = '';
    if (!text) return;
    blocks.push(headingLevel ? { kind: 'heading', level: headingLevel, text } : { kind: 'text', text });
  };
  const cur = (): TableState | undefined => tables[0];
  const addText = (t: string) => {
    if (depth === 0) { line += t; return; }
    if (depth > 1) return;
    const tb = cur()!;
    if (tb.cell !== null) tb.cell += t;
  };
  const newline = () => {
    if (depth === 0) { flushLine(); return; }
    if (depth > 1) return;
    const tb = cur()!;
    if (tb.cell !== null && !tb.cell.endsWith('\n')) tb.cell += '\n';
  };
  const closeCell = (tb: TableState) => {
    if (tb.cell === null) return;
    if (!tb.row) tb.row = [];
    tb.row.push(cleanCell(tb.cell));
    for (let k = 1; k < tb.span; k++) tb.row.push('');
    tb.cell = null; tb.span = 1;
  };
  const closeRow = (tb: TableState) => {
    closeCell(tb);
    if (tb.row && tb.row.length) tb.rows.push(tb.row);
    tb.row = null;
  };

  const re = /<(\/?)([a-zA-Z][\w:-]*)([^>]*)>|([^<]+)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(src)) !== null) {
    if (m[4] !== undefined) {
      addText(decodeEntities(m[4]).replace(/[\s\u00a0]+/g, ' '));
      continue;
    }
    const closing = m[1] === '/';
    const tag = m[2].toLowerCase();
    const attrs = m[3] || '';
    if (tag === 'table') {
      if (!closing) {
        if (depth === 0) { flushLine(); tables.unshift({ rows: [], row: null, cell: null, span: 1, merged: false, nested: false }); }
        else cur()!.nested = true;
        depth++;
      } else if (depth > 0) {
        depth--;
        if (depth === 0) {
          const tb = tables.shift()!;
          closeRow(tb);
          blocks.push({ kind: 'table', rows: tb.rows, merged: tb.merged, nested: tb.nested });
        }
      }
      continue;
    }
    if (depth === 1) {
      const tb = cur()!;
      if (tag === 'tr') { closeRow(tb); if (!closing) tb.row = []; continue; }
      if (tag === 'td' || tag === 'th') {
        if (!closing) {
          closeCell(tb);
          tb.cell = '';
          const cs = /colspan\s*=\s*["']?(\d+)/i.exec(attrs);
          const rs = /rowspan\s*=\s*["']?(\d+)/i.exec(attrs);
          tb.span = cs ? Math.max(1, parseInt(cs[1], 10)) : 1;
          if (tb.span > 1 || (rs && parseInt(rs[1], 10) > 1)) tb.merged = true;
        } else closeCell(tb);
        continue;
      }
    }
    if (tag === 'br') { newline(); continue; }
    if (/^h[1-6]$/.test(tag)) {
      if (depth === 0) {
        flushLine();
        headingLevel = closing ? 0 : parseInt(tag[1], 10);
      } else newline();
      continue;
    }
    if (BLOCK_TAGS.has(tag)) newline();
  }
  flushLine();
  const tb = cur();
  if (tb) { closeRow(tb); blocks.push({ kind: 'table', rows: tb.rows, merged: tb.merged, nested: tb.nested }); }
  return { blocks };
}
