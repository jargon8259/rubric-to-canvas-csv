import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseDocx, parseDocumentXml, countComments } from '../src/docx.ts';
import { htmlToBlocks } from '../src/html.ts';
import { textToBlocks } from '../src/plaintext.ts';
import { documentXml, makeDocx, makeZip, table, h, p } from './helpers/docxbuilder.ts';

const tableOf = (blocks: any[], i = 0) => blocks.filter((b) => b.kind === 'table')[i];

test('docx: headings, tables, multi-line cells, curly quotes and non-ASCII', async () => {
  const body = h(1, 'Rubric') + table([['Criterion Name', 'Competent'], ['A. Prompt', 'Line one\nLine \u201ctwo\u201d \u2013 caf\u00e9 \u4f60\u597d']]);
  for (const opts of [{}, { dataDescriptor: true }]) {
    const d = await parseDocx(makeDocx(body, {}, opts));
    assert.deepEqual(d.blocks[0], { kind: 'heading', level: 1, text: 'Rubric' });
    assert.deepEqual(tableOf(d.blocks).rows[1], ['A. Prompt', 'Line one\nLine \u201ctwo\u201d \u2013 caf\u00e9 \u4f60\u597d']);
  }
});
test('docx: tabs become spaces, line breaks are kept', () => {
  const xml = documentXml(table([['H'], [{ xml: '<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t></w:r></w:p>' }]]));
  assert.equal(tableOf(parseDocumentXml(xml).blocks).rows[1][0], 'a b\nc');
});
test('docx: tracked changes keep insertions, ignore deletions, and are counted', () => {
  const para = '<w:p><w:r><w:t>Keep </w:t></w:r><w:ins w:id="1" w:author="x"><w:r><w:t>added</w:t></w:r></w:ins><w:del w:id="2" w:author="x"><w:r><w:delText> removed</w:delText></w:r></w:del></w:p>';
  const r = parseDocumentXml(documentXml(table([['H'], [{ xml: para }]])));
  assert.equal(tableOf(r.blocks).rows[1][0], 'Keep added');
  assert.equal(r.trackedInserts, 1);
  assert.equal(r.trackedDeletes, 1);
});
test('docx: a row deleted by a tracked change is skipped', () => {
  const xml = documentXml(table([['H'], ['gone'], ['stays']], ['', '<w:del w:id="3" w:author="x"/>', '']));
  assert.deepEqual(tableOf(parseDocumentXml(xml).blocks).rows, [['H'], ['stays']]);
});
test('docx: merged cells are padded and flagged', () => {
  const xml = documentXml(table([['H1', 'H2', 'H3'], [{ text: 'wide', span: 2 }, 'x'], [{ text: 'top', vMerge: 'restart' }, 'a', 'b'], [{ vMerge: 'continue' }, 'c', 'd']]));
  const t = tableOf(parseDocumentXml(xml).blocks);
  assert.equal(t.merged, true);
  assert.deepEqual(t.rows[1], ['wide', '', 'x']);
  assert.deepEqual(t.rows[3], ['', 'c', 'd']);
});
test('docx: content controls (sdt) are read through', () => {
  const sdt = '<w:sdt><w:sdtContent><w:p><w:r><w:t>inside</w:t></w:r></w:p></w:sdtContent></w:sdt>';
  assert.deepEqual(parseDocumentXml(documentXml(sdt)).blocks, [{ kind: 'text', text: 'inside' }]);
});
test('docx: nested tables are flagged and ignored', () => {
  const xml = documentXml(table([['H'], [{ xml: p('outer') + table([['inner']]) }]]));
  const t = tableOf(parseDocumentXml(xml).blocks);
  assert.equal(t.nested, true);
  assert.equal(t.rows[1][0], 'outer');
});
test('docx: comments are counted, resolved ones are not "open"', async () => {
  const comments = '<w:comments xmlns:w="x"><w:comment w:id="0"><w:p w14:paraId="AA11"><w:r><w:t>a</w:t></w:r></w:p></w:comment><w:comment w:id="1"><w:p w14:paraId="BB22"><w:r><w:t>b</w:t></w:r></w:p></w:comment></w:comments>';
  const ext = '<w15:commentsEx><w15:commentEx w15:paraId="AA11" w15:done="1"/><w15:commentEx w15:paraId="BB22" w15:done="0"/></w15:commentsEx>';
  assert.deepEqual(countComments(comments, ext), { total: 2, open: 1 });
  assert.deepEqual(countComments(comments, null), { total: 2, open: 2 });
  assert.deepEqual(countComments(null, null), { total: 0, open: 0 });
  const d = await parseDocx(makeDocx(h(1, 'x'), { 'word/comments.xml': comments, 'word/commentsExtended.xml': ext }));
  assert.equal(d.comments, 2);
  assert.equal(d.openComments, 1);
});
test('docx: not a docx gives clear error codes', async () => {
  await assert.rejects(parseDocx(new TextEncoder().encode('plain text, long enough to scan for a zip end record....')), /NOT_ZIP/);
  await assert.rejects(parseDocx(makeZip({ 'x.txt': 'hi' })), /NO_DOCUMENT/);
});

test('html: Word-style table with multi-paragraph cells, entities and conditional comments', () => {
  const html = '<html><body><!--StartFragment--><h1>Rubric</h1><![if !supportLists]>x<![endif]><table><tr><td><p>Criterion&nbsp;Name</p></td><td><p>Competent</p></td></tr>'
    + '<tr><td><p>A. Prompt</p></td><td><p class=MsoNormal>Line one<o:p></o:p></p><p>Line &ldquo;two&rdquo;</p></td></tr></table><!--EndFragment--></body></html>';
  const { blocks } = htmlToBlocks(html);
  assert.deepEqual(blocks[0], { kind: 'heading', level: 1, text: 'Rubric' });
  assert.deepEqual(tableOf(blocks).rows, [['Criterion Name', 'Competent'], ['A. Prompt', 'Line one\nLine \u201ctwo\u201d']]);
});
test('html: br inside cell, th cells, colspan, source line wrapping collapses to spaces', () => {
  const html = '<table><tr><th>H1</th><th colspan="2">H2</th></tr><tr><td>one<br>two</td><td>wrapped\n   text</td><td>z</td></tr></table>';
  const t = tableOf(htmlToBlocks(html).blocks);
  assert.equal(t.merged, true);
  assert.deepEqual(t.rows, [['H1', 'H2', ''], ['one\ntwo', 'wrapped text', 'z']]);
});
test('html: nested table flagged; text outside tables kept', () => {
  const { blocks } = htmlToBlocks('<p>Before</p><table><tr><td>outer<table><tr><td>inner</td></tr></table></td></tr></table><p>After</p>');
  assert.equal(tableOf(blocks).nested, true);
  assert.deepEqual(tableOf(blocks).rows, [['outer']]);
  assert.deepEqual(blocks.filter((b) => b.kind === 'text').map((b: any) => b.text), ['Before', 'After']);
});
test('html: Excel-style paste', () => {
  const html = '<table><tr><td>Criterion Name</td><td>Competent</td></tr><tr><td>A</td><td style="x:y">desc &amp; more</td></tr></table>';
  assert.deepEqual(tableOf(htmlToBlocks(html).blocks).rows[1], ['A', 'desc & more']);
});

test('plain text: tab separated, blank cells kept, quoted multi-line cells (Excel); known headers split tables', () => {
  const { blocks } = textToBlocks('Field\tValue\nAssessment Name (learner facing)\t\nCode\tX1\n\nCriterion Name\tPoints\nA\t"line1\nline2"\n');
  assert.deepEqual(tableOf(blocks, 0).rows, [['Field', 'Value'], ['Assessment Name (learner facing)', ''], ['Code', 'X1']]);
  assert.deepEqual(tableOf(blocks, 1).rows, [['Criterion Name', 'Points'], ['A', 'line1\nline2']]);
});
test('plain text: single-field lines become text and split tables', () => {
  const { blocks } = textToBlocks('a\tb\nHeading line\nc\td');
  assert.equal(blocks.length, 3);
  assert.equal(blocks[1].kind, 'text');
});
test('plain text: comma separated', () => {
  assert.deepEqual(tableOf(textToBlocks('Criterion Name,Competent\nA,"x, y"').blocks).rows[1], ['A', 'x, y']);
});
test('plain text: a stray quote does not swallow the rest', () => {
  const r = textToBlocks('Criterion Name\tCompetent\nA\t"unclosed\nB\tfine');
  assert.equal(r.usedFallback, true);
  assert.equal(tableOf(r.blocks).rows.length, 3);
});
