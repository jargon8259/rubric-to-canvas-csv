import type { Block, ParsedDoc } from './types.ts';
import { listZip, readZipText } from './zip.ts';
import { parseXml, child } from './xml.ts';
import type { XNode } from './xml.ts';
import { cleanCell } from './text.ts';

const SKIP_IN_TEXT = new Set(['w:pPr', 'w:rPr', 'w:del', 'w:moveFrom', 'w:drawing', 'w:pict', 'mc:AlternateContent', 'w:object']);

/** Text of a paragraph (or any run container). Deleted text is skipped; inserted text is kept. */
function inlineText(node: XNode): string {
  let out = '';
  for (const c of node.children) {
    if (c.tag === '#text') continue;
    if (SKIP_IN_TEXT.has(c.tag)) continue;
    switch (c.tag) {
      case 'w:t':
        for (const t of c.children) if (t.tag === '#text') out += t.text;
        break;
      case 'w:tab': out += '\t'; break;
      case 'w:cr': out += '\n'; break;
      case 'w:br': if (!c.attrs['w:type'] || c.attrs['w:type'] === 'textWrapping') out += '\n'; break;
      case 'w:noBreakHyphen': out += '-'; break;
      default: out += inlineText(c);
    }
  }
  return out;
}

function paragraphStyle(p: XNode): string {
  const ppr = child(p, 'w:pPr');
  const st = ppr ? child(ppr, 'w:pStyle') : undefined;
  return st ? st.attrs['w:val'] || '' : '';
}

/** Content controls wrap paragraphs; flatten them so callers see plain w:p and w:tbl nodes. */
function flatten(nodes: XNode[]): XNode[] {
  const out: XNode[] = [];
  for (const n of nodes) {
    if (n.tag === 'w:sdt') {
      const content = child(n, 'w:sdtContent');
      if (content) out.push(...flatten(content.children));
    } else if (n.tag === 'w:ins' || n.tag === 'w:moveTo' || n.tag === 'w:customXml') {
      out.push(...flatten(n.children));
    } else {
      out.push(n);
    }
  }
  return out;
}

interface TableFlags { merged: boolean; nested: boolean }

function tableRows(tbl: XNode, flags: TableFlags): string[][] {
  const rows: string[][] = [];
  for (const tr of flatten(tbl.children).filter((n) => n.tag === 'w:tr')) {
    const trPr = child(tr, 'w:trPr');
    if (trPr && child(trPr, 'w:del')) continue; // row deleted by a tracked change
    const cells: string[] = [];
    for (const tc of flatten(tr.children).filter((n) => n.tag === 'w:tc')) {
      const tcPr = child(tc, 'w:tcPr');
      const span = tcPr ? child(tcPr, 'w:gridSpan') : undefined;
      const vm = tcPr ? child(tcPr, 'w:vMerge') : undefined;
      const items = flatten(tc.children);
      if (items.some((n) => n.tag === 'w:tbl')) flags.nested = true;
      const paras = items.filter((n) => n.tag === 'w:p').map((p) => inlineText(p));
      const isContinuation = vm && vm.attrs['w:val'] !== 'restart';
      if (vm) flags.merged = true;
      cells.push(isContinuation ? '' : cleanCell(paras.join('\n')));
      const n = span ? parseInt(span.attrs['w:val'] || '1', 10) : 1;
      if (n > 1) {
        flags.merged = true;
        for (let k = 1; k < n; k++) cells.push('');
      }
    }
    rows.push(cells);
  }
  return rows;
}

export function parseDocumentXml(xml: string): Pick<ParsedDoc, 'blocks' | 'trackedInserts' | 'trackedDeletes'> {
  const root = parseXml(xml);
  const doc = child(root, 'w:document');
  const body = doc ? child(doc, 'w:body') : undefined;
  if (!body) throw new Error('NO_BODY');
  const blocks: Block[] = [];
  for (const n of flatten(body.children)) {
    if (n.tag === 'w:p') {
      const text = cleanCell(inlineText(n));
      if (!text) continue;
      const m = /^heading\s*([1-9])$/i.exec(paragraphStyle(n));
      blocks.push(m ? { kind: 'heading', level: parseInt(m[1], 10), text } : { kind: 'text', text });
    } else if (n.tag === 'w:tbl') {
      const flags: TableFlags = { merged: false, nested: false };
      const rows = tableRows(n, flags);
      blocks.push({ kind: 'table', rows, merged: flags.merged, nested: flags.nested });
    }
  }
  const count = (re: RegExp) => (xml.match(re) || []).length;
  return {
    blocks,
    trackedInserts: count(/<w:(?:ins|moveTo)\s/g),
    trackedDeletes: count(/<w:(?:del|moveFrom)\s/g),
  };
}

/** Comments in word/comments.xml, and how many are not marked resolved in commentsExtended.xml. */
export function countComments(commentsXml: string | null, extendedXml: string | null): { total: number; open: number } {
  if (!commentsXml) return { total: 0, open: 0 };
  const parts = commentsXml.split(/<w:comment\s/).slice(1);
  const done = new Set<string>();
  if (extendedXml) {
    const re = /<w15:commentEx\b[^>]*>/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(extendedXml)) !== null) {
      const id = /w15:paraId="([0-9A-Fa-f]+)"/.exec(m[0]);
      const dn = /w15:done="(1|true)"/.exec(m[0]);
      if (id && dn) done.add(id[1].toUpperCase());
    }
  }
  let open = 0;
  for (const part of parts) {
    const ids = part.match(/w14:paraId="([0-9A-Fa-f]+)"/g) || [];
    const last = ids.length ? /"([0-9A-Fa-f]+)"/.exec(ids[ids.length - 1])![1].toUpperCase() : '';
    if (!last || !done.has(last)) open++;
  }
  return { total: parts.length, open };
}

export async function parseDocx(bytes: Uint8Array): Promise<ParsedDoc> {
  const entries = listZip(bytes);
  const xml = await readZipText(bytes, entries, 'word/document.xml');
  if (xml === null) throw new Error('NO_DOCUMENT');
  const parsed = parseDocumentXml(xml);
  const comments = countComments(
    await readZipText(bytes, entries, 'word/comments.xml'),
    await readZipText(bytes, entries, 'word/commentsExtended.xml'),
  );
  return { ...parsed, kind: 'docx', comments: comments.total, openComments: comments.open };
}
