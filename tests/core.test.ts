import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseDelimited, toCsv, detectDelimiter, escapeField } from '../src/csv.ts';
import { parseXml } from '../src/xml.ts';
import { decodeEntities, cleanCell } from '../src/text.ts';
import { listZip, readZipText } from '../src/zip.ts';
import { makeZip } from './helpers/docxbuilder.ts';

test('csv: quotes, doubled quotes, embedded newlines, CRLF', () => {
  const r = parseDelimited('a,"b,1","say ""hi""","line1\nline2"\r\nc,d,e,f\r\n', ',');
  assert.deepEqual(r.rows, [['a', 'b,1', 'say "hi"', 'line1\nline2'], ['c', 'd', 'e', 'f']]);
  assert.equal(r.unterminatedQuote, false);
});
test('csv: BOM stripped, blank lines dropped, trailing empty fields kept', () => {
  assert.deepEqual(parseDelimited('\ufeffa,b,\n\n,,\nx,y,z', ',').rows, [['a', 'b', ''], ['x', 'y', 'z']]);
});
test('csv: tab delimited', () => {
  assert.deepEqual(parseDelimited('a\tb\n"x\ny"\tz', '\t').rows, [['a', 'b'], ['x\ny', 'z']]);
  assert.equal(detectDelimiter('a\tb'), '\t');
  assert.equal(detectDelimiter('a,b'), ',');
});
test('csv: unterminated quote is reported', () => {
  assert.equal(parseDelimited('a,"b\nc', ',').unterminatedQuote, true);
});
test('csv: output uses \\n, trailing newline, no BOM, standard quoting', () => {
  assert.equal(toCsv([['a', 'b,c', 'd"e', 'f\ng', 'h\r\ni']]), 'a,"b,c","d""e","f\ng","h\ni"\n');
  assert.equal(escapeField('plain'), 'plain');
});
test('text helpers', () => {
  assert.equal(decodeEntities('A&amp;B &#8220;x&#x201D; &nbsp;&unknown;&constructor;'), 'A&B \u201cx\u201d  &unknown;&constructor;');
  assert.equal(cleanCell('  a  b \n\n  c\u00a0d  \n'), 'a b\nc d');
});
test('xml: attributes, entities, self-closing, comments, quotes containing >', () => {
  const root = parseXml('<?xml version="1.0"?><a x="1&amp;2" y=\'q>r\'><!-- c --><b/><c>t&lt;</c></a>');
  const a = root.children[0];
  assert.equal(a.attrs.x, '1&2');
  assert.equal(a.attrs.y, 'q>r');
  assert.equal(a.children[0].tag, 'b');
  assert.equal(a.children[1].children[0].text, 't<');
});
test('zip: reads deflated and stored entries, with and without data descriptors', async () => {
  for (const opts of [{}, { dataDescriptor: true }, { store: true }, { store: true, dataDescriptor: true }]) {
    const z = makeZip({ 'a.txt': 'hello wörld '.repeat(50), 'b/c.xml': '<x/>' }, opts);
    const entries = listZip(z);
    assert.deepEqual(entries.map((e) => e.name), ['a.txt', 'b/c.xml']);
    assert.equal(await readZipText(z, entries, 'a.txt'), 'hello wörld '.repeat(50));
    assert.equal(await readZipText(z, entries, 'missing'), null);
  }
});
test('zip: rejects non-zip data', () => {
  assert.throws(() => listZip(new TextEncoder().encode('this is not a zip file at all, just text')), /NOT_ZIP/);
});
