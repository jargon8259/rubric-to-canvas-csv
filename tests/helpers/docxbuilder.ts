import { deflateRawSync } from 'node:zlib';

function crc32(buf: Buffer): number {
  let c, crc = 0xffffffff;
  for (let n = 0; n < buf.length; n++) {
    c = (crc ^ buf[n]) & 0xff;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crc = (crc >>> 8) ^ c;
  }
  return (crc ^ 0xffffffff) >>> 0;
}

export interface ZipOpts { dataDescriptor?: boolean; store?: boolean }

/** Minimal zip writer. With dataDescriptor, local headers carry zero sizes, like Word's output often does. */
export function makeZip(files: Record<string, string>, opts: ZipOpts = {}): Uint8Array {
  const parts: Buffer[] = [];
  const central: Buffer[] = [];
  let offset = 0;
  for (const [name, content] of Object.entries(files)) {
    const raw = Buffer.from(content, 'utf8');
    const method = opts.store ? 0 : 8;
    const data = opts.store ? raw : deflateRawSync(raw);
    const crc = crc32(raw);
    const nameB = Buffer.from(name, 'utf8');
    const flags = opts.dataDescriptor ? 0x08 : 0;
    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(flags, 6); lh.writeUInt16LE(method, 8);
    lh.writeUInt32LE(opts.dataDescriptor ? 0 : crc, 14);
    lh.writeUInt32LE(opts.dataDescriptor ? 0 : data.length, 18);
    lh.writeUInt32LE(opts.dataDescriptor ? 0 : raw.length, 22);
    lh.writeUInt16LE(nameB.length, 26);
    const local = [lh, nameB, data];
    if (opts.dataDescriptor) {
      const dd = Buffer.alloc(16);
      dd.writeUInt32LE(0x08074b50, 0); dd.writeUInt32LE(crc, 4); dd.writeUInt32LE(data.length, 8); dd.writeUInt32LE(raw.length, 12);
      local.push(dd);
    }
    const chunk = Buffer.concat(local);
    const ch = Buffer.alloc(46);
    ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6); ch.writeUInt16LE(flags, 8); ch.writeUInt16LE(method, 10);
    ch.writeUInt32LE(crc, 16); ch.writeUInt32LE(data.length, 20); ch.writeUInt32LE(raw.length, 24);
    ch.writeUInt16LE(nameB.length, 28); ch.writeUInt32LE(offset, 42);
    central.push(Buffer.concat([ch, nameB]));
    parts.push(chunk);
    offset += chunk.length;
  }
  const cd = Buffer.concat(central);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0); eocd.writeUInt16LE(central.length, 8); eocd.writeUInt16LE(central.length, 10);
  eocd.writeUInt32LE(cd.length, 12); eocd.writeUInt32LE(offset, 16);
  return new Uint8Array(Buffer.concat([...parts, cd, eocd]));
}

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

export const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';

export function p(text: string, style?: string, raw?: string): string {
  const ppr = style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : '';
  return `<w:p>${ppr}${raw ?? `<w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r>`}</w:p>`;
}
export const h = (level: number, text: string) => p(text, `Heading${level}`);

export interface CellSpec { text?: string; xml?: string; span?: number; vMerge?: 'restart' | 'continue' }
export function cell(c: string | CellSpec): string {
  const s: CellSpec = typeof c === 'string' ? { text: c } : c;
  const props = (s.span ? `<w:gridSpan w:val="${s.span}"/>` : '') + (s.vMerge === 'restart' ? '<w:vMerge w:val="restart"/>' : s.vMerge === 'continue' ? '<w:vMerge/>' : '');
  const tcPr = props ? `<w:tcPr>${props}</w:tcPr>` : '';
  const body = s.xml ?? (s.text ?? '').split('\n').map((line) => p(line)).join('');
  return `<w:tc>${tcPr}${body || p('')}</w:tc>`;
}
export function table(rows: (string | CellSpec)[][], trPrXml: string[] = []): string {
  return `<w:tbl><w:tblPr/><w:tblGrid/>${rows.map((r, i) => `<w:tr>${trPrXml[i] ? `<w:trPr>${trPrXml[i]}</w:trPr>` : ''}${r.map(cell).join('')}</w:tr>`).join('')}</w:tbl>`;
}

export function documentXml(bodyXml: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W_NS}><w:body>${bodyXml}</w:body></w:document>`;
}

export function makeDocx(bodyXml: string, extra: Record<string, string> = {}, opts: ZipOpts = {}): Uint8Array {
  return makeZip({
    '[Content_Types].xml': '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
    'word/document.xml': documentXml(bodyXml),
    ...extra,
  }, opts);
}
