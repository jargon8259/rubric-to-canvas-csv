export interface ZipEntry {
  name: string;
  method: number;
  flags: number;
  compSize: number;
  size: number;
  localOffset: number;
}

const MAX_UNCOMPRESSED = 100 * 1024 * 1024;

function u16(d: DataView, o: number): number { return d.getUint16(o, true); }
function u32(d: DataView, o: number): number { return d.getUint32(o, true); }

/** Reads the central directory. Sizes come from here because local headers may hold zeros (data descriptors). */
export function listZip(bytes: Uint8Array): ZipEntry[] {
  const d = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1;
  const lowest = Math.max(0, bytes.length - 22 - 0xffff);
  for (let i = bytes.length - 22; i >= lowest; i--) {
    if (u32(d, i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('NOT_ZIP');
  const count = u16(d, eocd + 10);
  const cdOffset = u32(d, eocd + 16);
  if (count === 0xffff || cdOffset === 0xffffffff) throw new Error('ZIP64');
  const entries: ZipEntry[] = [];
  let p = cdOffset;
  const dec = new TextDecoder('utf-8');
  for (let n = 0; n < count; n++) {
    if (p + 46 > bytes.length || u32(d, p) !== 0x02014b50) throw new Error('NOT_ZIP');
    const flags = u16(d, p + 8);
    const method = u16(d, p + 10);
    const compSize = u32(d, p + 20);
    const size = u32(d, p + 24);
    const nameLen = u16(d, p + 28);
    const extraLen = u16(d, p + 30);
    const commentLen = u16(d, p + 32);
    const localOffset = u32(d, p + 42);
    const name = dec.decode(bytes.subarray(p + 46, p + 46 + nameLen));
    entries.push({ name, method, flags, compSize, size, localOffset });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

async function inflateRaw(data: Uint8Array, limit: number): Promise<Uint8Array> {
  const stream = new Blob([data as BlobPart]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  const reader = stream.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.length;
    if (total > limit) { await reader.cancel(); throw new Error('TOO_BIG'); }
    chunks.push(value);
  }
  const out = new Uint8Array(total);
  let o = 0;
  for (const c of chunks) { out.set(c, o); o += c.length; }
  return out;
}

export async function readZipEntry(bytes: Uint8Array, entry: ZipEntry): Promise<Uint8Array> {
  if (entry.flags & 1) throw new Error('ENCRYPTED');
  if (entry.size > MAX_UNCOMPRESSED) throw new Error('TOO_BIG');
  const d = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const p = entry.localOffset;
  if (p + 30 > bytes.length || u32(d, p) !== 0x04034b50) throw new Error('NOT_ZIP');
  const start = p + 30 + u16(d, p + 26) + u16(d, p + 28);
  const raw = bytes.subarray(start, start + entry.compSize);
  if (entry.method === 0) return raw;
  if (entry.method === 8) return inflateRaw(raw, MAX_UNCOMPRESSED);
  throw new Error('METHOD');
}

export async function readZipText(bytes: Uint8Array, entries: ZipEntry[], name: string): Promise<string | null> {
  const e = entries.find((x) => x.name === name);
  if (!e) return null;
  return new TextDecoder('utf-8').decode(await readZipEntry(bytes, e));
}
