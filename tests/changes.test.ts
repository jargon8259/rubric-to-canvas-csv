import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseDocumentXml } from '../src/docx.ts';
import { interpret } from '../src/rba.ts';
import { diffModels, formatNotes } from '../src/changes.ts';
import { validate } from '../src/validate.ts';
import { Store, addCriterion, removeCriterion, moveCriterion } from '../src/state.ts';
import type { ParsedDoc } from '../src/types.ts';
import { documentXml } from './helpers/docxbuilder.ts';
import { rbaBody } from './helpers/rba.ts';

const doc = (xml: string): ParsedDoc => ({ blocks: parseDocumentXml(documentXml(xml)).blocks, kind: 'docx', comments: 0, openComments: 0, trackedInserts: 0, trackedDeletes: 0 });
const fresh = (body = rbaBody()) => new Store(interpret(doc(body), 'RBA.docx').model!);
const keys = (s: Store) => diffModels(s.baseline, s.model).map((c) => c.key);
const msgs = (s: Store) => diffModels(s.baseline, s.model).map((c) => c.message);

test('no edits: nothing to report', () => {
  assert.deepEqual(diffModels(fresh().baseline, fresh().model).length >= 0, true);
  const s = fresh();
  assert.deepEqual(diffModels(s.baseline, s.model), []);
});
test('changing only the export setting (Criteria Enable Range) is not reported: it is not in the RBA', () => {
  const s = fresh();
  s.model.enableRange = true;
  assert.deepEqual(diffModels(s.baseline, s.model), []);
});
test('the baseline is a snapshot: editing the model never changes it', () => {
  const s = fresh();
  s.model.criteria[0].name = 'Changed';
  s.model.criteria[0].cells[0] = 'Changed';
  assert.equal(s.baseline.criteria[0].name, 'A. First Prompt');
  assert.equal(s.baseline.criteria[0].cells[0], 'Meets all of A.');
});

test('rubric name typed because the RBA name was blank', () => {
  const s = fresh(rbaBody({ meta: { 'Assessment Name (learner facing)': '' } }));
  s.model.rubricName = 'My Assessment';
  assert.deepEqual(keys(s), ['name']);
  assert.match(msgs(s)[0], /typed in the app.*My Assessment.*Assessment Name \(learner facing\)/);
});
test('rubric name changed from the RBA value', () => {
  const s = fresh();
  s.model.rubricName = 'Different Name';
  assert.match(msgs(s)[0], /changed from "Sample Assessment" to "Different Name".*update it there/);
});
test('expected total and rating renames are reported', () => {
  const s = fresh();
  s.model.expectedPoints = '90';
  s.model.ratings[1].name = 'Approaching';
  assert.deepEqual(keys(s), ['total', 'rating:1']);
  assert.match(msgs(s)[0], /from "100" to "90".*Assessment Points/);
  assert.match(msgs(s)[1], /Rating column 2 renamed.*"Approaching Competency".*"Approaching"/);
});
test('criterion edits: name, points, description (edited, filled in, cleared)', () => {
  const s = fresh(rbaBody({ criteria: [['A. First', '60', 'Full', '', 'None', 'No submission.'], ['B. Second', '40', 'Full', 'Some', 'None', 'No submission.']] }));
  const [a, b] = s.model.criteria;
  a.name = 'A. First (fixed)';
  a.points = '55';
  a.cells[0] = 'Full credit';      // edited
  a.cells[1] = 'Some credit';      // filled in (was blank)
  b.cells[1] = '';                 // cleared
  const out = msgs(s);
  assert.equal(out.length, 5);
  assert.ok(out.some((m) => /name changed from "A. First" to "A. First \(fixed\)"/.test(m)));
  assert.ok(out.some((m) => /"A. First \(fixed\)": points changed from "60" to "55"/.test(m)));
  assert.ok(out.some((m) => /Competent description edited. Was: "Full" Now: "Full credit"/.test(m)));
  assert.ok(out.some((m) => /Approaching Competency description filled in. Was: \(blank\)/.test(m)));
  assert.ok(out.some((m) => /"B. Second": Approaching Competency description cleared/.test(m)));
});
test('long text is shortened in the messages', () => {
  const s = fresh();
  s.model.criteria[0].cells[0] = 'word '.repeat(80);
  assert.ok(msgs(s)[0].length < 260);
  assert.ok(msgs(s)[0].includes('\u2026'));
});
test('added and removed criteria are reported with their position and points', () => {
  const s = fresh();
  addCriterion(s.model, 2);
  s.model.criteria[2].name = 'C. New';
  s.model.criteria[2].points = '10';
  removeCriterion(s.model, 0);
  const out = msgs(s);
  assert.equal(out.length, 2);
  assert.ok(out.some((m) => /Criterion removed: "A. First Prompt"/.test(m)));
  assert.ok(out.some((m) => /Criterion added: "C. New" \(row 2, 10 points\)/.test(m)));
});
test('a blank added row is described by position', () => {
  const s = fresh();
  addCriterion(s.model, 1);
  assert.match(msgs(s)[0], /Criterion added: Row 2 \(row 2, no points\)/);
});
test('reordering is one change, and does not look like renames or point edits', () => {
  const s = fresh();
  moveCriterion(s.model, 1, -1);
  assert.deepEqual(keys(s), ['order']);
  assert.match(msgs(s)[0], /order was changed/);
});
test('a rename plus a move reports both, because rows are tracked by id, not position', () => {
  const s = fresh();
  s.model.criteria[1].name = 'B. Renamed';
  moveCriterion(s.model, 1, -1);
  const k = keys(s);
  assert.ok(k.includes('order') && k.some((x) => x.startsWith('ren:')));
  assert.equal(k.length, 2);
});

test('it compares end states: typing then restoring the original value reports nothing', () => {
  const s = fresh();
  s.checkpoint(); s.model.criteria[0].points = '99';
  assert.equal(keys(s).length, 1);
  s.checkpoint(); s.model.criteria[0].points = '60';
  assert.deepEqual(keys(s), []);
});
test('undo removes the reported edit; redo brings it back', () => {
  const s = fresh();
  s.checkpoint(); s.model.rubricName = 'Changed';
  assert.deepEqual(keys(s), ['name']);
  s.undo();
  assert.deepEqual(keys(s), []);
  s.redo();
  assert.deepEqual(keys(s), ['name']);
});
test('undoing a removal restores the same row id, so it is no longer reported', () => {
  const s = fresh();
  s.checkpoint(); removeCriterion(s.model, 0);
  assert.equal(keys(s).length, 1);
  s.undo();
  assert.deepEqual(keys(s), []);
});
test('whitespace-only differences are ignored', () => {
  const s = fresh();
  s.model.criteria[0].name = '  A. First Prompt ';
  s.model.criteria[0].cells[0] = 'Meets all of A.\r\n';
  assert.deepEqual(diffModels(s.baseline, s.model), []);
});

test('notes text: numbered edits, parsing-only section, and warnings still open', () => {
  const s = fresh(rbaBody({ criteria: [['A', '30', 'x', '', '', ''], ['B', '40', 'y', '', '', '']] })); // totals 70, so one warning
  s.model.criteria[0].points = '35';
  s.model.criteria[1].name = 'B. Fixed';
  const edits = diffModels(s.baseline, s.model);
  const parsing = new Set([edits.find((e) => e.key.startsWith('ren:'))!.key]);
  const text = formatNotes({ rubricName: s.model.rubricName, source: 'RBA.docx', edits, parsingOnly: parsing, warnings: validate(s.model).filter((i) => i.severity === 'warn') });
  assert.match(text, /^RBA update notes\nRubric: Sample Assessment\nLoaded from: RBA.docx/);
  assert.match(text, /Make the same fix in the RBA[^\n]*\n1\. "A": points changed from "30" to "35"/);
  assert.match(text, /Marked as parsing only[^\n]*\n- Criterion name changed from "B" to "B. Fixed"/);
  assert.match(text, /Warnings still showing[^\n]*\n- The criterion points add up to 75/);
  assert.ok(text.endsWith('\n'));
});
test('notes text with nothing to report', () => {
  const s = fresh();
  const text = formatNotes({ rubricName: 'X', source: 's', edits: [], parsingOnly: new Set(), warnings: [] });
  assert.match(text, /Make the same fix in the RBA[^\n]*\nNone\./);
  assert.match(text, /not fixed in the app\):\nNone\./);
  assert.ok(!text.includes('parsing only'));
  void s;
});
