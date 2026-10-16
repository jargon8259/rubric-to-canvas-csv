// Runs the BUILT single-file app (dist/rubric-to-canvas.html) in a Node vm against a fake DOM.
// It proves the shipped script has no scope errors and the main flows work. It is not a visual test.
import { test, before } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import vm from 'node:vm';
import { FakeDocument, FakeNode } from './helpers/fakedom.ts';
import { makeDocx } from './helpers/docxbuilder.ts';
import { rbaBody } from './helpers/rba.ts';
import { parseDelimited } from '../src/csv.ts';

const html = () => readFileSync(new URL('../dist/rubric-to-canvas.html', import.meta.url), 'utf8');
const sleep = (ms = 30) => new Promise((r) => setTimeout(r, ms));

before(() => { execFileSync(process.execPath, ['build.mjs'], { cwd: new URL('..', import.meta.url).pathname, stdio: 'ignore' }); });

function boot() {
  const scripts = [...html().matchAll(/<script>\n?([\s\S]*?)\n?<\/script>/g)].map((m) => m[1]);
  const doc = new FakeDocument();
  const root = new FakeNode('div', doc);
  const downloads: { name: string; blob: Blob }[] = [];
  const copied: string[] = [];
  const realCreate = URL.createObjectURL.bind(URL);
  const sandbox: any = {
    document: Object.assign(doc, { getElementById: (id: string) => (id === 'app' ? root : null) }),
    console, setTimeout, clearTimeout, atob, Blob,
    URL: Object.assign(function () {}, { createObjectURL: (b: Blob) => { sandbox._last = b; return realCreate(b); }, revokeObjectURL() {} }),
    TextDecoder, TextEncoder, DecompressionStream, Response, Uint8Array, Promise,
    navigator: { clipboard: { writeText: async (t: string) => { copied.push(t); } } },
  };
  sandbox.globalThis = sandbox;
  sandbox.window = sandbox;
  vm.createContext(sandbox);
  const origCreate = doc.createElement.bind(doc);
  doc.createElement = (tag: string) => {
    const n = origCreate(tag);
    if (tag === 'a') n.click = () => downloads.push({ name: n.attrs.download, blob: sandbox._last });
    return n;
  };
  for (const s of scripts) vm.runInContext(s, sandbox);
  return { doc, root, downloads, copied };
}

const pasteHtml = (app: any, htmlText: string, plain = '') =>
  app.doc.fire('paste', { target: app.root, clipboardData: { files: [], getData: (t: string) => (t === 'text/html' ? htmlText : plain) } });

const RBA_HTML = (name: string, pts1 = '60', pts2 = '40', assessmentPoints = '100') => `<h1>Metadata</h1><table><tr><td>Field</td><td>Value</td></tr><tr><td>Assessment Name (learner facing)</td><td>${name}</td></tr><tr><td>Assessment Points</td><td>${assessmentPoints}</td></tr></table>
  <h1>Rubric</h1><table><tr><td>Criterion Name</td><td>Points</td><td>Competent</td><td>Approaching Competency</td><td>Not Competent</td><td>Not Submitted</td></tr>
  <tr><td>A. One</td><td>${pts1}</td><td>Good</td><td>Some</td><td>None</td><td>No submission.</td></tr>
  <tr><td>B. Two</td><td>${pts2}</td><td>Good, "quoted"</td><td>Some</td><td>None</td><td>No submission.</td></tr></table>`;

const type = (el: any, v: string) => { el.value = v; el.fire('input'); };

test('built file: starts with the empty state and hides review and export', () => {
  const app = boot();
  assert.ok(app.root.textContent.includes('Rubric to Canvas CSV'));
  assert.equal(app.root.byTest('review').hidden, true);
  assert.equal(app.root.byTest('export').hidden, true);
});

test('built file: blank assessment name loads, warns, and blocks export until named; CSV has per-criterion points', async () => {
  const app = boot();
  const e = pasteHtml(app, RBA_HTML(''));
  assert.equal(e.defaultPrevented, true);
  const R = app.root;
  assert.equal(R.byTest('review').hidden, false);
  assert.ok(R.byTest('status').textContent.includes('2 criteria'));
  const dl = R.byTest('download');
  assert.equal(dl.disabled, true);
  assert.ok(R.byTest('reason').textContent.includes('Add a rubric name'));
  assert.ok(R.byTest('summary').textContent.includes('100 (expected 100)'));
  assert.equal(R.byTest('calc-0').textContent, 'Scores: 60 / 30 / 0 / 0');
  type(R.byTest('rubric-name'), 'My Rubric: Draft?');
  assert.equal(dl.disabled, false);
  assert.equal(R.byTest('copy').disabled, false);
  assert.ok(R.byTest('filename').textContent.includes('My Rubric Draft.csv'));
  const csv = R.byTest('csv-preview').textContent;
  assert.ok(csv.startsWith('Rubric Name,Criteria Name,Criteria Description,Criteria Enable Range,Rating Name,Rating Description,Rating Points,'));
  assert.ok(csv.includes('My Rubric: Draft?,A. One,,false,Competent,Good,60,Approaching Competency,Some,30,Not Competent,None,0,Not Submitted,No submission.,0\n'));
  assert.ok(csv.includes('"Good, ""quoted"""'));
  dl.click();
  assert.equal(app.downloads.length, 1);
  assert.equal(app.downloads[0].name, 'My Rubric Draft.csv');
  assert.equal(await app.downloads[0].blob.text(), csv);
  assert.equal(csv.charCodeAt(0), 82);
});

test('built file: Assessment Points is read from the RBA; a different total is checked against it, and an editable field is shown', () => {
  const app = boot();
  pasteHtml(app, RBA_HTML('Legacy', '50', '30', '80'));
  const R = app.root;
  assert.equal(R.byTest('expected-points').value, '80');
  assert.ok(R.byTest('summary').textContent.includes('80 (expected 80)'));
  assert.ok(!R.byTest('issues').textContent.includes('add up to'));
  assert.ok(R.byTest('issues').textContent.includes('not the usual 100'));
  type(R.byTest('cpts-1'), '40');
  assert.ok(R.byTest('issues').textContent.includes('add up to 90'));
  assert.equal(R.byTest('download').disabled, false); // warning only
  type(R.byTest('expected-points'), '90');
  assert.ok(!R.byTest('issues').textContent.includes('add up to'));
});

test('built file: missing Assessment Points defaults to 100', () => {
  const app = boot();
  pasteHtml(app, '<table><tr><td>Criterion Name</td><td>Points</td><td>Competent</td></tr><tr><td>A</td><td>100</td><td>x</td></tr></table>');
  assert.equal(app.root.byTest('expected-points').value, '100');
  assert.ok(app.root.byTest('summary').textContent.includes('100 (expected 100)'));
});

test('built file: no edits means the export says so and the CSV matches the RBA', async () => {
  const app = boot();
  pasteHtml(app, RBA_HTML('Clean'));
  const R = app.root;
  assert.ok(R.byTest('edit-summary').textContent.includes('No edits were made'));
  R.byTest('download').click();
  assert.ok(R.byTest('export-status').textContent.includes('No edits were made in the app'));
  assert.ok(!R.byTest('changes').className.includes('exported'));
});

test('built file: edits are listed live, and the export reminds the ADS to fix the RBA', async () => {
  const app = boot();
  pasteHtml(app, RBA_HTML('Named'));
  const R = app.root;
  type(R.byTest('cpts-0'), '55');
  type(R.byTest('cell-1-0'), 'Better text');
  type(R.byTest('cpts-1'), '45');
  const list = R.byTest('edit-list').textContent;
  assert.ok(list.includes('"A. One": points changed from "60" to "55"'));
  assert.ok(list.includes('"B. Two": Competent description edited'));
  assert.ok(R.byTest('edit-summary').textContent.includes('3 edits made in the app are not in the RBA yet'));
  R.byTest('download').click();
  const st = R.byTest('export-status').textContent;
  assert.ok(st.includes('Downloaded Named.csv'));
  assert.ok(st.includes('3 edits made in the app are not in the RBA yet'));
  assert.ok(R.byTest('changes').className.includes('exported'));
  assert.equal(app.downloads.length, 1);
  // The CSV itself carries the edited values and nothing about the edits.
  const csv = await app.downloads[0].blob.text();
  assert.ok(csv.includes('A. One,,false,Competent,Good,55,'));
  assert.ok(!/Better text.*Was:|Parsing|RBA update/.test(csv));
});

test('built file: "parsing fix only" moves an edit out of the to-do count, and keeps keyboard focus', async () => {
  const app = boot();
  pasteHtml(app, RBA_HTML('Named'));
  const R = app.root;
  type(R.byTest('cpts-0'), '55');
  type(R.byTest('cpts-1'), '45');
  const first = R.find((n) => (n.attrs['data-test'] || '').startsWith('po-'))[0];
  first.checked = true; first.fire('change');
  assert.ok(R.byTest('edit-summary').textContent.includes('1 edit made in the app is not in the RBA yet'));
  assert.equal(app.doc.activeElement.attrs['data-test'], first.attrs['data-test']); // focus returns to the same checkbox
  const second = R.find((n) => (n.attrs['data-test'] || '').startsWith('po-') && !n.checked)[0];
  second.checked = true; second.fire('change');
  assert.ok(R.byTest('edit-summary').textContent.includes('All 2 edits are marked as parsing fixes'));
  R.byTest('download').click();
  assert.ok(R.byTest('export-status').textContent.includes('All edits are marked as parsing fixes'));
});

test('built file: still-flagged warnings are listed, and "Copy these notes" copies the full text', async () => {
  const app = boot();
  pasteHtml(app, RBA_HTML('Named', '30', '40')); // totals 70 vs 100
  const R = app.root;
  assert.ok(R.byTest('warn-list').textContent.includes('add up to 70'));
  type(R.byTest('cpts-0'), '35');
  R.byTest('copy-notes').click();
  await sleep();
  assert.equal(app.copied.length, 1);
  assert.match(app.copied[0], /^RBA update notes\nRubric: Named\nLoaded from: pasted content/);
  assert.match(app.copied[0], /1\. "A. One": points changed from "30" to "35"/);
  assert.match(app.copied[0], /Warnings still showing[^\n]*\n- The criterion points add up to 75/);
  assert.ok(R.byTest('export-status').textContent.includes('Copied the RBA update notes'));
});

test('built file: undo clears the reported edit; loading again or starting over clears the list and the ticks', async () => {
  const app = boot();
  pasteHtml(app, RBA_HTML('Named'));
  const R = app.root;
  type(R.byTest('cpts-0'), '55');
  assert.ok(R.byTest('edit-list').textContent.includes('points changed'));
  R.byTest('undo').click();
  assert.ok(R.byTest('edit-summary').textContent.includes('No edits were made'));
  type(R.byTest('cpts-0'), '50');
  const cb = R.find((n) => (n.attrs['data-test'] || '').startsWith('po-'))[0];
  cb.checked = true; cb.fire('change');
  pasteHtml(app, RBA_HTML('Second'));
  assert.ok(R.byTest('edit-summary').textContent.includes('No edits were made'));
  type(R.byTest('cpts-0'), '50');
  assert.ok(R.byTest('edit-summary').textContent.includes('1 edit made'), 'ticks from the previous load must not carry over');
});

test('built file: removing and reordering rows are reported', () => {
  const app = boot();
  pasteHtml(app, RBA_HTML('Named'));
  const R = app.root;
  R.byTest('down-0').click();
  assert.ok(R.byTest('edit-list').textContent.includes('Criteria order was changed'));
  R.byTest('remove-0').click();
  assert.ok(R.byTest('edit-list').textContent.includes('Criterion removed: "B. Two"'));
});

test('built file: the page never prints "null", "undefined" or "NaN"', () => {
  const app = boot();
  const bad = /\bnull\b|\bundefined\b|\bNaN\b|\[object/;
  assert.ok(!bad.test(app.root.textContent));
  pasteHtml(app, RBA_HTML(''));
  assert.ok(!bad.test(app.root.textContent), app.root.textContent.match(/.{20}(null|undefined|NaN|\[object).{20}/)?.[0]);
  type(app.root.byTest('rubric-name'), 'x');
  type(app.root.byTest('cpts-0'), '');
  assert.ok(!bad.test(app.root.textContent), app.root.textContent.match(/.{20}(null|undefined|NaN|\[object).{20}/)?.[0]);
});

test('built file: edit, add, move and remove criteria, with undo and redo', async () => {
  const app = boot();
  pasteHtml(app, RBA_HTML('Named'));
  const R = app.root;
  const rowsOf = () => R.find((n) => /^cname-\d+$/.test(n.attrs['data-test'] || '')).map((n) => n.value);
  assert.deepEqual(rowsOf(), ['A. One', 'B. Two']);
  type(R.byTest('cname-0'), 'A edited');
  assert.ok(R.byTest('csv-preview').textContent.includes('A edited'));
  R.byTest('down-0').click();
  assert.deepEqual(rowsOf(), ['B. Two', 'A edited']);
  assert.equal(R.byTest('cpts-1').value, '60');
  R.byTest('add-criterion').click();
  assert.equal(rowsOf().length, 3);
  assert.ok(R.byTest('reason').textContent.includes('no name'));
  R.byTest('remove-2').click();
  assert.deepEqual(rowsOf(), ['B. Two', 'A edited']);
  R.byTest('undo').click(); assert.equal(rowsOf().length, 3);
  R.byTest('undo').click(); R.byTest('undo').click();
  assert.deepEqual(rowsOf(), ['A edited', 'B. Two']);
  R.byTest('undo').click();
  assert.deepEqual(rowsOf(), ['A. One', 'B. Two']);
  R.byTest('redo').click();
  assert.deepEqual(rowsOf(), ['A edited', 'B. Two']);
});

test('built file: a non-standard rating name blocks export until renamed in the unlocked header', () => {
  const app = boot();
  pasteHtml(app, '<table><tr><td>Criterion Name</td><td>Points</td><td>Competent</td><td>Mystery</td></tr><tr><td>A</td><td>100</td><td>x</td><td>y</td></tr></table>');
  const R = app.root;
  type(R.byTest('rubric-name'), 'N');
  assert.equal(R.byTest('download').disabled, true);
  const u = R.byTest('unlock'); u.checked = true; u.fire('change');
  type(R.byTest('rname-1'), 'Not Competent');
  assert.equal(R.byTest('download').disabled, false);
  assert.ok(R.byTest('csv-preview').textContent.includes(',Not Competent,y,0\n'));
  assert.ok(R.byTest('edit-list').textContent.includes('Rating column 2 renamed from "Mystery" to "Not Competent"'));
});

test('built file: loading a .docx through the file input, then a Canvas CSV gives the plain message', async () => {
  const app = boot();
  const bytes = makeDocx(rbaBody());
  const input = app.root.byTest('file-input');
  input.files = [{ name: 'rba.docx', arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) }];
  input.fire('change');
  await sleep(100);
  assert.equal(app.root.byTest('review').hidden, false);
  assert.ok(app.root.byTest('summary').textContent.includes('Sample Assessment'));
  assert.equal(app.root.byTest('download').disabled, false);
  const canvas = 'Rubric Name,Criteria Name,Criteria Description,Criteria Enable Range,Rating Name,Rating Description,Rating Points\nR,C,,false,A,B,1\n';
  input.files = [{ name: 'canvas.csv', text: async () => canvas }];
  input.fire('change');
  await sleep();
  assert.ok(app.root.byTest('status').textContent.includes('already in Canvas'));
  assert.equal(app.root.byTest('review').hidden, false);
});

test('built file: friendly errors for .doc and unknown files; start over needs two clicks', async () => {
  const app = boot();
  const input = app.root.byTest('file-input');
  input.files = [{ name: 'old.doc' }]; input.fire('change'); await sleep();
  assert.ok(app.root.byTest('status').textContent.includes('Save As'));
  input.files = [{ name: 'x.pdf' }]; input.fire('change'); await sleep();
  assert.ok(app.root.byTest('status').textContent.includes('.docx, .csv, .tsv and .txt'));
  input.files = [{ name: 'bad.docx', arrayBuffer: async () => new TextEncoder().encode('not a zip at all, but long enough to scan').buffer }];
  input.fire('change'); await sleep();
  assert.ok(app.root.byTest('status').textContent.includes('does not look like a Word'));
  pasteHtml(app, RBA_HTML('N'));
  app.root.byTest('reset').click();
  assert.equal(app.root.byTest('review').hidden, false);
  app.root.byTest('reset').click();
  assert.equal(app.root.byTest('review').hidden, true);
});

test('built file: the bundled sample loads, is clean, and exports a CSV with no edits to report', async () => {
  const app = boot();
  app.root.byTest('load-sample').click();
  await sleep(150);
  const R = app.root;
  assert.equal(R.byTest('download').disabled, false);
  assert.ok(R.byTest('summary').textContent.includes('Sample Reflective Journal'));
  assert.ok(R.byTest('summary').textContent.includes('100 (expected 100)'));
  assert.equal(R.byTest('calc-0').textContent, 'Scores: 15 / 8 / 0 / 0');
  const rows = parseDelimited(R.byTest('csv-preview').textContent, ',').rows;
  assert.equal(rows.length, 11);
  assert.ok(rows.every((r) => r.length === 16));
  assert.ok(R.byTest('edit-summary').textContent.includes('No edits were made'));
  assert.ok(R.byTest('changes').textContent.includes('No open warnings'));
});

test('built file: privacy checks on the shipped HTML', () => {
  const src = html();
  assert.ok(/Content-Security-Policy[^>]*connect-src 'none'/.test(src));
  assert.ok(/default-src 'none'/.test(src));
  for (const bad of ['fetch(', 'XMLHttpRequest', 'WebSocket', 'sendBeacon', 'EventSource', 'importScripts', 'eval(', 'new Function', 'localStorage', 'sessionStorage', 'indexedDB', 'document.cookie']) {
    assert.ok(!src.includes(bad), 'found ' + bad);
  }
  assert.equal([...src.matchAll(/(?:src|href|action)\s*=\s*["']?(https?:)?\/\//gi)].length, 0, 'external resource reference found');
  assert.ok(!/@import|url\(\s*["']?https?:/i.test(src));
  assert.ok(!/<script[^>]+src=/i.test(src) && !/<link[^>]+href=/i.test(src));
});
