import { decodeEntities } from './text.ts';

export interface XNode {
  tag: string;
  attrs: Record<string, string>;
  children: XNode[];
  /** Only set on '#text' nodes. */
  text?: string;
}

/** A small, forgiving XML reader. Enough for WordprocessingML; not a general parser. */
export function parseXml(xml: string): XNode {
  const root: XNode = { tag: '#root', attrs: {}, children: [] };
  const stack: XNode[] = [root];
  let i = 0;
  const n = xml.length;
  const top = () => stack[stack.length - 1];
  while (i < n) {
    const lt = xml.indexOf('<', i);
    if (lt < 0) break;
    if (lt > i) {
      const t = xml.slice(i, lt);
      if (t !== '') top().children.push({ tag: '#text', attrs: {}, children: [], text: decodeEntities(t) });
    }
    if (xml.startsWith('<!--', lt)) { const e = xml.indexOf('-->', lt + 4); i = e < 0 ? n : e + 3; continue; }
    if (xml.startsWith('<![CDATA[', lt)) {
      const e = xml.indexOf(']]>', lt + 9);
      const end = e < 0 ? n : e;
      top().children.push({ tag: '#text', attrs: {}, children: [], text: xml.slice(lt + 9, end) });
      i = e < 0 ? n : e + 3; continue;
    }
    if (xml.startsWith('<?', lt) || xml.startsWith('<!', lt)) { const e = xml.indexOf('>', lt + 2); i = e < 0 ? n : e + 1; continue; }
    let j = lt + 1;
    let quote = '';
    while (j < n) {
      const c = xml[j];
      if (quote) { if (c === quote) quote = ''; } else if (c === '"' || c === "'") quote = c; else if (c === '>') break;
      j++;
    }
    const inner = xml.slice(lt + 1, j);
    i = j + 1;
    if (inner[0] === '/') {
      const name = inner.slice(1).trim();
      for (let k = stack.length - 1; k > 0; k--) {
        if (stack[k].tag === name) { stack.length = k; break; }
      }
      continue;
    }
    const selfClose = inner.endsWith('/');
    const body = selfClose ? inner.slice(0, -1) : inner;
    const m = /^([^\s/]+)([\s\S]*)$/.exec(body);
    if (!m) continue;
    const attrs: Record<string, string> = {};
    const re = /([^\s=]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
    let a: RegExpExecArray | null;
    while ((a = re.exec(m[2])) !== null) attrs[a[1]] = decodeEntities(a[3] !== undefined ? a[3] : a[4]);
    const node: XNode = { tag: m[1], attrs, children: [] };
    top().children.push(node);
    if (!selfClose) stack.push(node);
  }
  return root;
}

export function child(node: XNode, tag: string): XNode | undefined {
  return node.children.find((c) => c.tag === tag);
}
