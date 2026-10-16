import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { parseDocx, parseDocumentXml } from '../src/docx.ts';
import { htmlToBlocks } from '../src/html.ts';
import { textToBlocks } from '../src/plaintext.ts';
import { parseDelimited, detectDelimiter } from '../src/csv.ts';
import { interpret, parseRatingHeader, looksLikeCanvasCsv, normalizePoints } from '../src/rba.ts';
import { validate, blockers, totalPoints, expectedTotal } from '../src/validate.ts';
import { scoreFor, parsePoints, criterionScores, ruleFor } from '../src/scoring.ts';
import { buildCsv, buildCsvRows, csvFileName } from '../src/export.ts';
import { Store, addCriterion, removeCriterion, moveCriterion } from '../src/state.ts';
import type { ParsedDoc, RubricModel } from '../src/types.ts';
import { documentXml, makeDocx, table, h } from './helpers/docxbuilder.ts';
import { rbaBody, RUBRIC_HEAD, metaRows } from './helpers/rba.ts';

const doc = (blocks: ParsedDoc['blocks'], kind: ParsedDoc['kind'] = 'docx', extra: Partial<ParsedDoc> = {}): ParsedDoc =>
  ({ blocks, kind, comments: 0, openComments: 0, trackedInserts: 0, trackedDeletes: 0, ...extra });
const fromXml = (body: string) => doc(parseDocumentXml(documentXml(body)).blocks);
const load = (body: string) => interpret(fromXml(body), 'x').model!;
const ids = (m: RubricModel) => validate(m).map((i) => i.id);
const noId = (cs: RubricModel['criteria']) => cs.map(({ id, ...rest }) => rest);
const NAMES = ['Competent', 'Approaching Competency', 'Not Competent', 'Not Submitted'];

test('scoring rule: full, half rounded up, zero, zero', () => {
  const s = (pts: string) => criterionScores(NAMES.map((name) => ({ name })), pts);
  assert.deepEqual(s('20'), [20, 10, 0, 0]);
  assert.deepEqual(s('15'), [15, 8, 0, 0]);
  assert.deepEqual(s('5'), [5, 3, 0, 0]);
  assert.deepEqual(s('1'), [1, 1, 0, 0]);
  assert.deepEqual(s('100'), [100, 50, 0, 0]);
  assert.deepEqual(s('0'), [0, 0, 0, 0]);
  assert.deepEqual(s('abc'), [null, null, null, null]);
  assert.equal(scoreFor('half', 7), 4);
  assert.equal(ruleFor('  approaching   COMPETENCY '), 'half');
  assert.equal(ruleFor('constructor'), null);
  assert.equal(parsePoints('-3'), null);
  assert.equal(parsePoints('2.5'), 2.5);
  assert.equal(parsePoints('1e3'), null);
  assert.equal(parsePoints(''), null);
});
test('points cell variants normalize', () => {
  assert.equal(normalizePoints(' 20 '), '20');
  assert.equal(normalizePoints('20 points'), '20');
  assert.equal(normalizePoints('15 pts'), '15');
  assert.equal(normalizePoints('2.5'), '2.5');
  assert.equal(normalizePoints('twenty'), 'twenty');
});
test('rating headers: plain names; old "(10 points)" suffix is stripped and reported', () => {
  assert.deepEqual(parseRatingHeader('Competent'), { name: 'Competent', legacyPoints: null });
  assert.deepEqual(parseRatingHeader('Competent  (10 points)'), { name: 'Competent', legacyPoints: '10' });
  assert.deepEqual(parseRatingHeader('Approaching Competency (5 pts)'), { name: 'Approaching Competency', legacyPoints: '5' });
});

test('full RBA: rubric name is the Assessment Name; points per criterion; total matches Assessment Points', () => {
  const { model, error } = interpret(fromXml(rbaBody()), 'x.docx');
  assert.equal(error, undefined);
  const m = model!;
  assert.equal(m.rubricName, 'Sample Assessment');
  assert.equal(m.expectedPoints, '100');
  assert.deepEqual(m.ratings.map((r) => r.name), NAMES);
  assert.deepEqual(m.criteria.map((c) => c.points), ['60', '40']);
  assert.ok(m.criteria.every((c) => typeof c.id === 'string' && c.id.length > 0));
  assert.equal(new Set(m.criteria.map((c) => c.id)).size, 2);
  assert.equal(m.criteria[0].cells[3], 'No submission.');
  assert.deepEqual(blockers(validate(m)), []);
  assert.equal(totalPoints(m), 100);
  assert.ok(!ids(m).includes('ap-mismatch'));
});
test('metadata table is not mistaken for the rubric (it comes first in the file)', () => {
  assert.equal(load(rbaBody()).criteria[0].name, 'A. First Prompt');
});

test('Assessment Points is read from Metadata, not forced to 100', () => {
  const m = load(rbaBody({ meta: { 'Assessment Points': '80' }, criteria: [['A', '50', 'x', '', '', ''], ['B', '30', 'y', '', '', '']] }));
  assert.equal(m.expectedPoints, '80');
  assert.equal(expectedTotal(m), 80);
  assert.ok(!ids(m).includes('ap-mismatch'));
  assert.ok(m.loadIssues.some((i) => i.id === 'ap-nonstandard' && i.severity === 'info'));
  m.criteria[1].points = '40';
  const mismatch = validate(m).find((i) => i.id === 'ap-mismatch')!;
  assert.equal(mismatch.severity, 'warn');
  assert.match(mismatch.message, /add up to 90.*is 80/);
  assert.deepEqual(blockers(validate(m)), []);
});
test('Assessment Points: "100 points" text is understood', () => {
  assert.equal(load(rbaBody({ meta: { 'Assessment Points': '100 points' } })).expectedPoints, '100');
});
test('Assessment Points defaults to 100 when the value is blank', () => {
  const m = load(rbaBody({ meta: { 'Assessment Points': '' } }));
  assert.equal(m.expectedPoints, '100');
  assert.ok(m.loadIssues.some((i) => i.id === 'ap-default' && i.severity === 'info'));
  assert.ok(!ids(m).includes('ap-mismatch'));
});
test('Assessment Points defaults to 100 when the row is missing from Metadata', () => {
  const rows = metaRows().filter((r) => r[0] !== 'Assessment Points');
  const m = load(rbaBody({ metaRows: rows }));
  assert.equal(m.expectedPoints, '100');
  assert.ok(m.loadIssues.some((i) => i.id === 'ap-default'));
});
test('Assessment Points that is not a number warns and falls back to 100', () => {
  const m = load(rbaBody({ meta: { 'Assessment Points': 'TBD' } }));
  assert.equal(m.expectedPoints, '100');
  assert.ok(m.loadIssues.some((i) => i.id === 'ap-text' && i.severity === 'warn'));
});
test('no Metadata table (rubric-only paste): defaults to 100 and says the name was not loaded', () => {
  const m = load(table([RUBRIC_HEAD, ['A', '70', '', '', '', ''], ['B', '30', '', '', '', '']]));
  assert.equal(m.expectedPoints, '100');
  assert.ok(m.loadIssues.some((i) => i.id === 'no-meta'));
  assert.ok(!ids(m).includes('ap-mismatch'));
});
test('expected total can be edited; a blank or invalid value is checked against 100', () => {
  const m = load(rbaBody());
  m.expectedPoints = '90';
  assert.ok(ids(m).includes('ap-mismatch'));
  m.expectedPoints = '';
  assert.equal(expectedTotal(m), 100);
  assert.ok(!ids(m).includes('ap-mismatch') && !ids(m).includes('ap-invalid'));
  m.expectedPoints = 'lots';
  assert.equal(expectedTotal(m), 100);
  assert.ok(ids(m).includes('ap-invalid'));
});

test('old "Rubric Name (Canvas)" row is not used; blank assessment name blocks export until typed', () => {
  const m = load(rbaBody({ meta: { 'Assessment Name (learner facing)': '', 'Rubric Name (Canvas)': 'Old Name' } }));
  assert.equal(m.rubricName, '');
  assert.ok(validate(m).some((i) => i.id === 'name-blank' && i.severity === 'block' && i.field === 'rubricName'));
  m.rubricName = 'Typed Name';
  assert.deepEqual(blockers(validate(m)), []);
});
test('template placeholder as the rubric name is flagged', () => {
  assert.ok(ids(load(rbaBody({ meta: { 'Assessment Name (learner facing)': '[Assessment Name (learner facing)]' } }))).includes('name-placeholder'));
});

test('points total not matching: warns, does not block', () => {
  const m = load(rbaBody({ criteria: [['A', '30', 'x', '', '', ''], ['B', '40', 'y', '', '', '']] }));
  const v = validate(m);
  assert.ok(v.some((i) => i.id === 'ap-mismatch' && i.severity === 'warn' && /add up to 70/.test(i.message)));
  assert.deepEqual(blockers(v), []);
});
test('criterion points: blank or text blocks; zero and decimals warn; 1 is noted', () => {
  const m = load(rbaBody({ criteria: [['A', '', '', '', '', ''], ['B', 'lots', '', '', '', ''], ['C', '0', '', '', '', ''], ['D', '7.5', '', '', '', ''], ['E', '1', '', '', '', ''], ['F', '-5', '', '', '', '']] }));
  const v = validate(m);
  const by = (id: string) => v.find((i) => i.id === id);
  assert.equal(by('cpts-0')!.severity, 'block');
  assert.equal(by('cpts-1')!.severity, 'block');
  assert.equal(by('czero-2')!.severity, 'warn');
  assert.equal(by('cdec-3')!.severity, 'warn');
  assert.equal(by('chalf-4')!.severity, 'info');
  assert.equal(by('cpts-5')!.severity, 'block');
  assert.equal(totalPoints(m), null);
  assert.ok(!ids(m).includes('ap-mismatch'));
});
test('points typed as "15 points" in Word still load as 15', () => {
  const m = load(rbaBody({ criteria: [['A', '15 points', '', '', '', ''], ['B', '85 pts', '', '', '', '']] }));
  assert.deepEqual(m.criteria.map((c) => c.points), ['15', '85']);
  assert.deepEqual(blockers(validate(m)), []);
});
test('no Points column (old layout): warns and blocks until points are entered in the grid', () => {
  const m = load(table([['Criterion Name', 'Competent', 'Approaching Competency', 'Not Competent', 'Not Submitted'], ['A', 'x', 'y', 'z', 'No submission.']]));
  assert.deepEqual(m.ratings.map((r) => r.name), NAMES);
  assert.equal(m.criteria[0].cells[0], 'x');
  assert.ok(m.loadIssues.some((i) => i.id === 'no-points-col'));
  m.rubricName = 'N';
  assert.ok(ids(m).includes('cpts-0'));
  m.criteria[0].points = '100';
  assert.deepEqual(blockers(validate(m)), []);
});
test('old layout with "(10 points)" headers and a Rating Scale table: names kept, old info ignored', () => {
  const m = load(table([['Rating Name', 'Points'], ['Competent', '10']]) + table([['Criteria', 'Competent  (10 points)', 'Not Competent (0 points)'], ['A', 'x', 'y']]));
  assert.deepEqual(m.ratings.map((r) => r.name), ['Competent', 'Not Competent']);
  const l = m.loadIssues.map((i) => i.id);
  assert.ok(l.includes('legacy-0') && l.includes('old-scale') && l.includes('no-points-col'));
});
test('unknown rating name blocks export; renaming it fixes it', () => {
  const m = load(table([['Criterion Name', 'Points', 'Competent', 'Mystery'], ['A', '100', 'x', 'y']]));
  m.rubricName = 'N';
  assert.deepEqual(blockers(validate(m)).map((i) => i.id), ['rrule-1']);
  m.ratings[1].name = 'not competent';
  assert.deepEqual(blockers(validate(m)), []);
});
test('no header row: six columns assume Points; five do not', () => {
  const six = interpret(doc([{ kind: 'table', rows: [['A. Prompt', '30', 'a', 'b', 'c', 'No submission.']], merged: false, nested: false }], 'html'), 'x').model!;
  assert.deepEqual(noId(six.criteria)[0], { name: 'A. Prompt', points: '30', cells: ['a', 'b', 'c', 'No submission.'] });
  assert.ok(six.loadIssues.some((i) => i.id === 'no-header'));
  const five = interpret(doc([{ kind: 'table', rows: [['A. Prompt', 'a', 'b', 'c', 'No submission.']], merged: false, nested: false }], 'html'), 'x').model!;
  assert.equal(five.criteria[0].points, '');
  assert.ok(five.loadIssues.some((i) => i.id === 'no-points-col'));
});
test('no tables at all, and tables without a rubric, give friendly errors', () => {
  assert.match(interpret(doc([{ kind: 'text', text: 'hi' }]), 'x').error!, /couldn't find any table/);
  assert.match(interpret(doc([{ kind: 'table', rows: [['Field', 'Value'], ['a', 'b']], merged: false, nested: false }]), 'x').error!, /Criterion Name/);
});
test('two rubric tables: uses the first and warns', () => {
  const t = table([RUBRIC_HEAD, ['A', '100', '', '', '', '']]);
  const m = load(t + t);
  assert.ok(m.loadIssues.some((i) => i.id === 'many-rubrics'));
  assert.equal(m.criteria.length, 1);
});
test('comments, tracked changes, merged cells and plain-text paste produce warnings, not errors', () => {
  const merged = table([RUBRIC_HEAD, [{ text: 'A', span: 2 }, '', '', '', '']]);
  const d = doc(parseDocumentXml(documentXml(merged)).blocks, 'docx', { comments: 3, openComments: 2, trackedInserts: 1, trackedDeletes: 4 });
  const l = interpret(d, 'x').model!.loadIssues.map((i) => i.id);
  assert.ok(l.includes('comments') && l.includes('tracked') && l.includes('merged'));
  assert.ok(interpret(doc(textToBlocks('Criterion Name\tPoints\tCompetent\nA\t100\tx').blocks, 'text'), 'x').model!.loadIssues.some((i) => i.id === 'plain-paste'));
});
test('a body row wider than the header becomes an unnamed, flagged column; nothing is lost', () => {
  const m = load(table([['Criterion Name', 'Points', 'Competent'], ['A', '100', 'x', 'extra text']]));
  assert.equal(m.ratings.length, 2);
  assert.equal(m.criteria[0].cells[1], 'extra text');
  assert.ok(ids(m).includes('rname-1'));
});
test('placeholders and duplicates are flagged', () => {
  const m = load(rbaBody({ criteria: [['#. Mechanics', '50', 'goes here', '', '', ''], ['Title of Sub-Prompt', '25', '', '', '', ''], ['Title of Sub-Prompt', '25', '', '', '', '']] }));
  const i = ids(m);
  assert.ok(i.includes('ph-0') && i.includes('ph-1') && i.includes('phc-0') && i.includes('dup-2'));
});
test('empty descriptions are informational; an empty rubric blocks', () => {
  const m = load(rbaBody({ criteria: [['A', '100', '', '', '', '']] }));
  assert.equal(validate(m).find((x) => x.id === 'empty')!.severity, 'info');
  m.criteria = [];
  assert.ok(blockers(validate(m)).some((x) => x.id === 'no-criteria'));
});
test('blank criterion name blocks export', () => {
  const m = load(rbaBody());
  m.criteria[1].name = '  ';
  assert.deepEqual(blockers(validate(m)).map((x) => x.id), ['cname-1']);
});

test('export: header shape, one row per criterion, name repeated, per-criterion points, enable range lowercase', () => {
  const m = load(rbaBody({ criteria: [['A. First', '85', 'Meets all.', 'Some.', 'None.', 'No submission.'], ['B. Second', '15', 'x', 'y', 'z', 'No submission.']] }));
  const rows = buildCsvRows(m);
  assert.equal(rows[0].length, 4 + 3 * 4);
  assert.deepEqual(rows[0].slice(0, 7), ['Rubric Name', 'Criteria Name', 'Criteria Description', 'Criteria Enable Range', 'Rating Name', 'Rating Description', 'Rating Points']);
  assert.deepEqual(rows[1].slice(0, 7), ['Sample Assessment', 'A. First', '', 'false', 'Competent', 'Meets all.', '85']);
  assert.deepEqual([rows[1][6], rows[1][9], rows[1][12], rows[1][15]], ['85', '43', '0', '0']);
  assert.deepEqual([rows[2][6], rows[2][9], rows[2][12], rows[2][15]], ['15', '8', '0', '0']);
  m.enableRange = true;
  assert.equal(buildCsvRows(m)[1][3], 'true');
});
test('export: the internal ids and expected total never appear in the CSV', () => {
  const m = load(rbaBody({ meta: { 'Assessment Points': '80' }, criteria: [['A', '80', 'x', '', '', '']] }));
  const csv = buildCsv(m);
  assert.ok(!/\bc\d+\b/.test(csv.split('\n')[1].split(',')[1]));
  assert.ok(!csv.includes('Assessment Points') && !csv.includes('expected'));
});
test('export: quoting, no BOM, trailing newline, \\n only', () => {
  const csv = buildCsv(load(rbaBody({ criteria: [['A, "quoted"', '100', 'two\nlines', '', '', '']] })));
  assert.ok(csv.endsWith('\n') && !csv.includes('\r') && csv.charCodeAt(0) === 82);
  assert.ok(csv.includes('"A, ""quoted"""'));
  assert.ok(csv.includes('"two\nlines"'));
});
test('export: file name from rubric name', () => {
  assert.equal(csvFileName('NUR101: Case/Study? "A"'), 'NUR101 CaseStudy A.csv');
  assert.equal(csvFileName('   '), 'rubric.csv');
  assert.equal(csvFileName('x'.repeat(300)).length, 104);
});
test('export: matches the structure of the Canvas template file', () => {
  const tpl = parseDelimited(readFileSync(new URL('./fixtures/canvas_template.csv', import.meta.url), 'utf8'), ',').rows;
  const mine = buildCsvRows(load(rbaBody()));
  assert.deepEqual(mine[0].slice(0, 4), tpl[0].slice(0, 4));
  for (let g = 4; g < tpl[0].length; g += 3) assert.deepEqual(tpl[0].slice(g, g + 3), mine[0].slice(4, 7));
  assert.ok(['true', 'false'].includes(tpl[1][3]));
});
test('round trip: exported CSV re-parses to the same data', () => {
  const m = load(rbaBody({ criteria: [['A, "q"', '70', 'a\nb', 'c,d', '', 'No submission.'], ['B', '30', '\u201cx\u201d \u2013 caf\u00e9', '', '', '']] }));
  const rows = parseDelimited(buildCsv(m), ',').rows;
  assert.deepEqual(rows, buildCsvRows(m));
  const back = rows.slice(1).map((r) => ({ name: r[1], cells: [r[5], r[8], r[11], r[14]], max: r[6] }));
  assert.deepEqual(back.map((b) => ({ name: b.name, cells: b.cells })), m.criteria.map((c) => ({ name: c.name, cells: c.cells })));
  assert.deepEqual(back.map((b) => b.max), ['70', '30']);
  assert.deepEqual([rows[1][4], rows[1][7], rows[1][10], rows[1][13]], m.ratings.map((r) => r.name));
});
test('canvas csv detection', () => {
  const tpl = parseDelimited(readFileSync(new URL('./fixtures/canvas_template.csv', import.meta.url), 'utf8'), ',').rows;
  assert.equal(looksLikeCanvasCsv(tpl), true);
  assert.equal(looksLikeCanvasCsv([RUBRIC_HEAD]), false);
});

test('same rubric from .docx, HTML paste, tab-text paste and CSV gives the same preview', async () => {
  const rows = [['A. First Prompt', '65', 'Line one\nLine two', 'Some, commas', 'Quote "x"', 'No submission.'], ['B. Second', '35', '', 'b', 'c', 'No submission.']];
  const fromDocx = interpret(await parseDocx(makeDocx(rbaBody({ criteria: rows }))), 'docx').model!;
  const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/\n/g, '<br>');
  const tbl = (r: string[][]) => `<table>${r.map((x) => `<tr>${x.map((c) => `<td><p>${esc(c)}</p></td>`).join('')}</tr>`).join('')}</table>`;
  const fromHtml = interpret(doc(htmlToBlocks(`<h1>Metadata</h1>${tbl(metaRows())}<h1>Rubric</h1>${tbl([RUBRIC_HEAD, ...rows])}`).blocks, 'html'), 'html').model!;
  const q = (s: string) => (/[\t\n"]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s);
  const tsv = (r: string[][]) => r.map((x) => x.map(q).join('\t')).join('\n');
  const fromText = interpret(doc(textToBlocks([tsv(metaRows()), tsv([RUBRIC_HEAD, ...rows])].join('\n')).blocks, 'text'), 'text').model!;
  const csvText = [RUBRIC_HEAD, ...rows].map((r) => r.map((c) => (/[,\n"]/.test(c) ? '"' + c.replace(/"/g, '""') + '"' : c)).join(',')).join('\n');
  const fromCsv = interpret(doc([{ kind: 'table', rows: parseDelimited(csvText, detectDelimiter(csvText)).rows, merged: false, nested: false }], 'delimited'), 'csv').model!;
  const core = (m: RubricModel) => ({ n: m.rubricName, e: m.expectedPoints, r: m.ratings, c: noId(m.criteria) });
  assert.deepEqual(core(fromHtml), core(fromDocx));
  assert.deepEqual(core(fromText), core(fromDocx));
  assert.deepEqual(noId(fromCsv.criteria), noId(fromDocx.criteria));
  assert.deepEqual(fromCsv.ratings, fromDocx.ratings);
  assert.equal(fromDocx.criteria[0].points, '65');
});

test('the real draft RBA template loads: 10 criteria, points blank, name blank, Assessment Points read, placeholders and comments flagged', async () => {
  const d = await parseDocx(new Uint8Array(readFileSync(new URL('./fixtures/RBA_template_draft.docx', import.meta.url))));
  const m = interpret(d, 'template').model!;
  assert.equal(m.criteria.length, 10);
  assert.deepEqual(m.ratings.map((r) => r.name), NAMES);
  assert.equal(m.criteria[0].name, 'A. Main Prompt');
  assert.equal(m.rubricName, '');
  assert.equal(m.expectedPoints, '100');
  assert.ok(m.criteria.every((c) => c.points === ''));
  const v = validate(m);
  assert.ok(v.some((i) => i.id === 'name-blank'));
  assert.equal(v.filter((i) => /^cpts-/.test(i.id)).length, 10);
  assert.ok(v.some((i) => i.id.startsWith('ph-')));
  assert.ok(m.loadIssues.some((i) => i.id === 'comments'));
  assert.ok(!m.loadIssues.some((i) => i.id === 'many-rubrics'));
});
test('the sample RBA is clean: named, 100 points, no blockers or warnings', async () => {
  const d = await parseDocx(new Uint8Array(readFileSync(new URL('../samples/sample-rba.docx', import.meta.url))));
  const m = interpret(d, 'sample').model!;
  assert.equal(m.rubricName, 'Sample Reflective Journal');
  assert.equal(totalPoints(m), 100);
  assert.deepEqual(blockers(validate(m)), []);
  assert.deepEqual(validate(m).filter((i) => i.severity === 'warn'), []);
  const rows = buildCsvRows(m);
  assert.deepEqual([rows[1][6], rows[1][9]], ['15', '8']);
  assert.deepEqual([rows[9][6], rows[9][9]], ['5', '3']);
});

test('store: undo and redo restore earlier versions; structural helpers; ids are unique', () => {
  const s = new Store(load(rbaBody()));
  s.checkpoint(); s.model.rubricName = 'Changed';
  s.checkpoint(); addCriterion(s.model, 1);
  assert.equal(s.model.criteria.length, 3);
  assert.deepEqual(noId([s.model.criteria[1]])[0], { name: '', points: '', cells: ['', '', '', ''] });
  assert.equal(new Set(s.model.criteria.map((c) => c.id)).size, 3);
  assert.ok(s.undo()); assert.equal(s.model.criteria.length, 2); assert.equal(s.model.rubricName, 'Changed');
  assert.ok(s.undo()); assert.equal(s.model.rubricName, 'Sample Assessment');
  assert.equal(s.undo(), false);
  assert.ok(s.redo() && s.redo());
  assert.equal(s.model.criteria.length, 3);
  s.checkpoint(); assert.equal(s.canRedo, false);
  assert.equal(moveCriterion(s.model, 0, -1), 0);
  assert.equal(moveCriterion(s.model, 0, 1), 1);
  assert.equal(s.model.criteria[1].name, 'A. First Prompt');
  removeCriterion(s.model, 1); assert.equal(s.model.criteria.length, 2);
});
