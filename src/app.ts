import type { Issue, ParsedDoc, RubricModel } from './types.ts';
import { Store, addCriterion, removeCriterion, moveCriterion } from './state.ts';
import { parseDocx } from './docx.ts';
import { htmlToBlocks } from './html.ts';
import { textToBlocks } from './plaintext.ts';
import { parseDelimited, detectDelimiter } from './csv.ts';
import { interpret, looksLikeCanvasCsv } from './rba.ts';
import { validate, blockers, totalPoints, expectedTotal } from './validate.ts';
import { criterionScores, ruleFor, formatNumber } from './scoring.ts';
import { buildCsv, csvFileName } from './export.ts';
import { diffModels, formatNotes } from './changes.ts';

export interface AppConfig {
  repoUrl?: string;
  version?: string;
  sampleDocxBase64?: string;
  sampleCsv?: string;
}

type Kid = Node | string | null | undefined | false;
type StatusKind = 'info' | 'warn' | 'error' | 'success';

let doc: Document;

/** Tiny element builder. Properties for form state, attributes for everything else. */
function mk(tag: string, props: Record<string, any> = {}, ...kids: Kid[]): any {
  const e: any = doc.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (v === undefined || v === null || v === false) continue;
    if (k === 'class') e.className = v;
    else if (k === 'text') e.textContent = v;
    else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else if (k === 'value' || k === 'checked' || k === 'disabled' || k === 'hidden' || k === 'open') e[k] = v;
    else e.setAttribute(k, v === true ? '' : String(v));
  }
  for (const kid of kids) {
    if (kid === null || kid === undefined || kid === false) continue;
    e.append(typeof kid === 'string' ? doc.createTextNode(kid) : kid);
  }
  return e;
}

/** Drop empty slots so replaceChildren never receives null (browsers would print the word "null"). */
function compact(items: Kid[]): Node[] {
  return items.filter((x): x is Node => x !== null && x !== undefined && x !== false && typeof x !== 'string');
}

const SEV_LABEL = { block: 'Blocks export', warn: 'Warning', info: 'Note' } as const;
const EMPTY_DOC: Omit<ParsedDoc, 'blocks' | 'kind'> = { comments: 0, openComments: 0, trackedInserts: 0, trackedDeletes: 0 };

function base64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

function ruleText(name: string): string {
  const r = ruleFor(name);
  return r === 'full' ? 'Full points' : r === 'half' ? 'Half, rounded up' : r === 'zero' ? '0 points' : 'Not a standard level';
}

function plural(n: number, one: string, many: string): string {
  return n === 1 ? one : many;
}

export function initApp(d: Document, root: HTMLElement, cfg: AppConfig): void {
  doc = d;
  let store: Store | null = null;
  let unlocked = false;
  let resetArmed = false;
  let editStarted = false;
  let dragDepth = 0;
  let cellRefs: any[][] = [];
  let noteCells: any[] = [];
  let calcRefs: any[] = [];
  let ruleRefs: any[] = [];
  const parsingOnly = new Set<string>(); // edit keys the ADS marked "parsing fix only"
  const parsingRefs = new Map<string, any>();

  // ---------- static shell ----------
  const status = mk('div', { class: 'status', 'aria-live': 'polite', 'data-test': 'status' });
  const fileInput = mk('input', { type: 'file', accept: '.docx,.csv,.tsv,.txt', class: 'sr-only', id: 'file-input', 'data-test': 'file-input', tabindex: '-1' });
  const pasteBox = mk('textarea', { id: 'paste-box', rows: 3, class: 'paste', 'data-test': 'paste', placeholder: 'Click here, then press Ctrl+V (or Cmd+V)' });
  const overlay = mk('div', { class: 'overlay', hidden: true, 'aria-hidden': 'true' }, mk('p', { text: 'Drop the file anywhere to load it' }));

  const steps = mk('ol', { class: 'steps' },
    mk('li', {}, mk('strong', { text: 'Load. ' }), 'Open the finished RBA (.docx), paste the whole document or just the rubric table, or load a CSV.'),
    mk('li', {}, mk('strong', { text: 'Check. ' }), 'Review the summary and fix anything flagged. You can edit text and points right in the grid.'),
    mk('li', {}, mk('strong', { text: 'Export. ' }), 'Download the Canvas rubric CSV for AIT. The app lists any edits you made so you can fix them in the RBA too.'),
  );

  const sampleDownloads = mk('p', { class: 'small' },
    'New here? ',
    cfg.sampleDocxBase64 ? mk('button', { type: 'button', class: 'link', 'data-test': 'load-sample', onclick: () => loadSample() }, 'Load the sample') : null,
    cfg.sampleDocxBase64 ? ' or ' : null,
    cfg.sampleDocxBase64 ? mk('button', { type: 'button', class: 'link', onclick: () => downloadBlob(new Blob([base64ToBytes(cfg.sampleDocxBase64!) as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }), 'sample-rba.docx') }, 'download the sample .docx') : null,
    cfg.sampleCsv ? (cfg.sampleDocxBase64 ? ' or ' : '') : null,
    cfg.sampleCsv ? mk('button', { type: 'button', class: 'link', onclick: () => downloadBlob(new Blob([cfg.sampleCsv!], { type: 'text/csv;charset=utf-8' }), 'sample-rubric-table.csv') }, 'download the sample CSV') : null,
    '.',
  );

  const loadSection = mk('section', { 'aria-labelledby': 'h-load' },
    mk('h2', { id: 'h-load', text: '1. Load' }),
    steps,
    mk('div', { class: 'loadbox' },
      mk('div', { class: 'row' },
        mk('button', { type: 'button', class: 'btn primary', 'data-test': 'choose', onclick: () => fileInput.click() }, 'Choose a file'),
        mk('span', { class: 'hint', text: '.docx, .csv or .tsv. You can also drop a file anywhere on this page.' }),
      ),
      mk('div', { class: 'pastewrap' }, mk('label', { for: 'paste-box', text: 'Or paste a table, or the whole RBA' }), pasteBox),
      sampleDownloads,
    ),
    status,
  );

  const reviewHeading = mk('h2', { id: 'h-review', tabindex: '-1', text: '2. Check and edit' });
  const summary = mk('dl', { class: 'summary', 'data-test': 'summary' });
  const issuesBox = mk('div', { class: 'issues', 'data-test': 'issues' });
  const nameInput = mk('input', { type: 'text', id: 'rubric-name', class: 'keep-paste', 'data-test': 'rubric-name', autocomplete: 'off', 'aria-describedby': 'name-help' });
  const nameHelp = mk('p', { id: 'name-help', class: 'small' });
  const expectedInput = mk('input', { type: 'text', id: 'expected-points', class: 'keep-paste', inputmode: 'decimal', 'data-test': 'expected-points', autocomplete: 'off', 'aria-describedby': 'expected-help' });
  const expectedHelp = mk('p', { id: 'expected-help', class: 'small', text: 'From Assessment Points in the Metadata table. Every RBA is normally 100, which is used when it is not found.' });
  const undoBtn = mk('button', { type: 'button', class: 'btn', 'data-test': 'undo', onclick: () => step('undo') }, 'Undo');
  const redoBtn = mk('button', { type: 'button', class: 'btn', 'data-test': 'redo', onclick: () => step('redo') }, 'Redo');
  const addBtn = mk('button', { type: 'button', class: 'btn', 'data-test': 'add-criterion', onclick: () => structural((m) => { addCriterion(m, m.criteria.length); return m.criteria.length - 1; }) }, 'Add criterion');
  const unlockBox = mk('input', { type: 'checkbox', id: 'unlock', 'data-test': 'unlock', onchange: () => { unlocked = unlockBox.checked; renderGrid(); refresh(); } });
  const rangeBox = mk('input', { type: 'checkbox', id: 'range', 'data-test': 'range', onchange: () => { if (!store) return; store.checkpoint(); store.model.enableRange = rangeBox.checked; refresh(); } });
  const resetBtn = mk('button', { type: 'button', class: 'btn quiet', 'data-test': 'reset', onclick: () => startOver() }, 'Start over');
  const gridHost = mk('div', { class: 'gridwrap', role: 'region', 'aria-label': 'Rubric grid', tabindex: '0' });

  const reviewSection = mk('section', { 'aria-labelledby': 'h-review', hidden: true, 'data-test': 'review' },
    reviewHeading,
    summary,
    mk('div', { class: 'fieldrow' },
      mk('div', { class: 'field' },
        mk('label', { for: 'rubric-name' }, 'Rubric name ', mk('span', { class: 'req', text: '(required)' })),
        nameInput, nameHelp,
      ),
      mk('div', { class: 'field short' },
        mk('label', { for: 'expected-points', text: 'Assessment points (expected total)' }),
        expectedInput, expectedHelp,
      ),
    ),
    issuesBox,
    mk('div', { class: 'toolbar' },
      undoBtn, redoBtn, addBtn,
      mk('label', { class: 'check', for: 'range' }, rangeBox, ' Criteria Enable Range'),
      mk('label', { class: 'check', for: 'unlock' }, unlockBox, ' Edit rating names'),
      resetBtn,
    ),
    mk('p', { class: 'small', text: 'Scoring: Competent = the criterion\'s points. Approaching Competency = half, rounded up. Not Competent and Not Submitted = 0. Tip: Ctrl+Up and Ctrl+Down move between rows in the same column. Enter in the name or points box moves to the next row.' }),
    gridHost,
  );

  const downloadBtn = mk('button', { type: 'button', class: 'btn primary', 'data-test': 'download', onclick: () => doDownload() }, 'Download CSV');
  const copyBtn = mk('button', { type: 'button', class: 'btn', 'data-test': 'copy', onclick: () => doCopy() }, 'Copy CSV to clipboard');
  const exportReason = mk('div', { class: 'reason', 'aria-live': 'polite', 'data-test': 'reason' });
  const exportStatus = mk('div', { class: 'exportstatus', 'aria-live': 'polite', 'data-test': 'export-status' });
  const fileNote = mk('p', { class: 'small', 'data-test': 'filename' });
  const preview = mk('pre', { class: 'csvpreview', tabindex: '0', 'data-test': 'csv-preview' });
  const changesBox = mk('div', { class: 'changes', 'data-test': 'changes' });
  const exportSection = mk('section', { 'aria-labelledby': 'h-export', hidden: true, 'data-test': 'export' },
    mk('h2', { id: 'h-export', text: '3. Export' }),
    mk('div', { class: 'row' }, downloadBtn, copyBtn),
    fileNote, exportReason, exportStatus,
    changesBox,
    mk('details', { class: 'previewbox' }, mk('summary', { text: 'Preview the exact CSV text' }), preview),
  );

  const footer = mk('footer', {},
    mk('h2', { text: 'How it works, and why it is private' }),
    mk('p', {}, 'This page reads your file with code that runs inside your browser tab. Nothing is uploaded or sent anywhere. There are no accounts, no analytics and no outside scripts, fonts or services, and the page blocks network requests with a content security policy. You can save this one HTML file and run it offline.'),
    mk('p', {}, 'The Canvas CSV is built from the rubric table in the RBA. The rubric name is the Assessment Name (learner facing) from the Metadata table. Each criterion\'s Points column is its maximum. Competent gets all of it, Approaching Competency gets half rounded up, and Not Competent and Not Submitted get 0. The criterion points should add up to Assessment Points (100 if the RBA does not say).'),
    mk('p', {}, 'The app remembers what the RBA said when you loaded it. Anything you change here is listed at export so you can make the same fix in the RBA. The app never changes your Word file.'),
    mk('p', { class: 'small' },
      cfg.version ? 'Version ' + cfg.version + '. ' : '',
      cfg.repoUrl ? mk('a', { href: cfg.repoUrl, rel: 'noopener noreferrer' }, 'Source code and README') : 'Source code and README are in the project repository.'),
  );

  root.replaceChildren(
    mk('a', { class: 'skip', href: '#main', text: 'Skip to main content' }),
    mk('header', {},
      mk('h1', { text: 'Rubric to Canvas CSV' }),
      mk('p', { class: 'lede', text: 'Turn the rubric in a finished RBA into a CSV that Canvas can import.' }),
      mk('p', { class: 'badge', text: 'Private by design: everything happens in this browser tab. Nothing is uploaded.' }),
    ),
    mk('main', { id: 'main' }, loadSection, reviewSection, exportSection),
    footer, fileInput, overlay,
  );

  // ---------- messages ----------
  function showStatus(text: string, kind: StatusKind = 'info'): void {
    status.replaceChildren(mk('p', { class: 'msg ' + kind, role: kind === 'error' ? 'alert' : 'status' }, text));
  }
  function setExportStatus(text: string, kind: StatusKind = 'info'): void {
    exportStatus.replaceChildren(...(text ? [mk('p', { class: 'msg ' + kind }, text)] : []));
  }

  // ---------- loading ----------
  async function loadFile(file: File): Promise<void> {
    const name = file.name || 'file';
    const ext = (/\.([^.]+)$/.exec(name)?.[1] || '').toLowerCase();
    try {
      if (ext === 'docx') {
        let parsed: ParsedDoc;
        try {
          parsed = await parseDocx(new Uint8Array(await file.arrayBuffer()));
        } catch (err) {
          showStatus(docxError(err), 'error');
          return;
        }
        accept(parsed, name);
      } else if (ext === 'csv' || ext === 'tsv' || ext === 'txt') {
        const text = await file.text();
        if (ext === 'txt') acceptText(text, name);
        else acceptDelimited(text, name);
      } else if (ext === 'doc') {
        showStatus('That is an older .doc file. Open it in Word, choose Save As, and pick Word Document (.docx). Then load the .docx.', 'error');
      } else if (ext === 'xlsx' || ext === 'xls') {
        showStatus('Excel files cannot be read directly. Select the rubric table in Excel and paste it here, or save it as CSV and load that.', 'error');
      } else {
        showStatus('I can read .docx, .csv, .tsv and .txt files. "' + name + '" is not one of those.', 'error');
      }
    } catch {
      showStatus('Something went wrong reading "' + name + '". Try the file again, or paste the table instead.', 'error');
    }
  }

  function docxError(err: unknown): string {
    const code = err instanceof Error ? err.message : '';
    if (code === 'NOT_ZIP' || code === 'NO_DOCUMENT') return 'That file does not look like a Word .docx document. If it is an older .doc, open it in Word and save it as .docx.';
    if (code === 'ENCRYPTED') return 'That document is password protected. Remove the password in Word, save it, and load it again.';
    if (code === 'TOO_BIG') return 'That document is too large to read here (over 100 MB once unpacked).';
    return 'I could not read that document. Try saving it again from Word as a .docx, or paste the rubric table instead.';
  }

  function canvasCsvMessage(): void {
    showStatus('This file is already in Canvas\'s import format (its first column is "Rubric Name"), so there is nothing to convert. You can import it into Canvas as it is. To start from the Word RBA instead, load the .docx.', 'info');
  }

  function acceptDelimited(text: string, label: string): void {
    const rows = parseDelimited(text, detectDelimiter(text)).rows;
    if (rows.length === 0) { showStatus('That file is empty.', 'error'); return; }
    if (looksLikeCanvasCsv(rows)) { canvasCsvMessage(); return; }
    accept({ blocks: [{ kind: 'table', rows, merged: false, nested: false }], kind: 'delimited', ...EMPTY_DOC }, label);
  }

  function acceptText(text: string, label: string): void {
    if (text.trim() === '') { showStatus('Nothing to load. The pasted text was empty.', 'error'); return; }
    const r = textToBlocks(text);
    const first = r.blocks.find((b) => b.kind === 'table');
    if (first && first.kind === 'table' && looksLikeCanvasCsv(first.rows)) { canvasCsvMessage(); return; }
    accept({ blocks: r.blocks, kind: 'text', ...EMPTY_DOC }, label);
  }

  function accept(parsed: ParsedDoc, label: string): void {
    const r = interpret(parsed, label);
    if (r.error || !r.model) { showStatus(r.error || 'I could not read a rubric from that.', 'error'); return; }
    store = new Store(r.model);
    resetArmed = false; resetBtn.textContent = 'Start over';
    unlocked = false; unlockBox.checked = false;
    parsingOnly.clear();
    setExportStatus('');
    changesBox.classList.remove('exported');
    reviewSection.hidden = false; exportSection.hidden = false;
    steps.hidden = true;
    renderGrid();
    refresh();
    const m = r.model;
    showStatus('Loaded ' + label + ': ' + m.criteria.length + ' criteria and ' + m.ratings.length + ' rating levels. Review the checks below.', 'success');
    reviewHeading.focus();
  }

  function loadSample(): void {
    if (!cfg.sampleDocxBase64) return;
    parseDocx(base64ToBytes(cfg.sampleDocxBase64)).then((p) => accept(p, 'the sample RBA'), (e) => showStatus(docxError(e), 'error'));
  }

  function onPaste(e: any): void {
    const t = e.target;
    if (t && t.classList && t.classList.contains('keep-paste')) return; // normal paste inside grid cells
    const cd = e.clipboardData;
    if (!cd) return;
    e.preventDefault();
    const file = cd.files && cd.files[0];
    if (file) { void loadFile(file); return; }
    const html = cd.getData('text/html') || '';
    const text = cd.getData('text/plain') || '';
    pasteBox.value = '';
    if (/<table[\s>]/i.test(html)) {
      const blocks = htmlToBlocks(html).blocks;
      const firstTable = blocks.find((b) => b.kind === 'table');
      if (firstTable && firstTable.kind === 'table' && looksLikeCanvasCsv(firstTable.rows)) { canvasCsvMessage(); return; }
      accept({ blocks, kind: 'html', ...EMPTY_DOC }, 'pasted content');
    } else {
      acceptText(text, 'pasted text');
    }
  }
  d.addEventListener('paste', onPaste);

  fileInput.addEventListener('change', () => {
    const f = fileInput.files && fileInput.files[0];
    if (f) void loadFile(f);
    fileInput.value = '';
  });

  // Drag and drop anywhere on the page
  const hasFiles = (e: any) => !!(e.dataTransfer && e.dataTransfer.types && Array.from(e.dataTransfer.types).includes('Files'));
  d.addEventListener('dragenter', (e: any) => { if (!hasFiles(e)) return; e.preventDefault(); dragDepth++; overlay.hidden = false; });
  d.addEventListener('dragover', (e: any) => { if (hasFiles(e)) e.preventDefault(); });
  d.addEventListener('dragleave', (e: any) => { if (!hasFiles(e)) return; dragDepth = Math.max(0, dragDepth - 1); if (dragDepth === 0) overlay.hidden = true; });
  d.addEventListener('drop', (e: any) => {
    if (!hasFiles(e)) return;
    e.preventDefault(); dragDepth = 0; overlay.hidden = true;
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) void loadFile(f);
  });

  // ---------- editing ----------
  function structural(op: (m: RubricModel) => number | void): void {
    if (!store) return;
    store.checkpoint();
    const row = op(store.model);
    renderGrid();
    refresh();
    if (typeof row === 'number') focusCell(row, 0);
  }

  function step(dir: 'undo' | 'redo'): void {
    if (!store) return;
    if (dir === 'undo' ? store.undo() : store.redo()) { renderGrid(); refresh(); }
  }

  function startOver(): void {
    if (!resetArmed) {
      resetArmed = true;
      resetBtn.textContent = 'Click again to discard this rubric';
      return;
    }
    store = null; resetArmed = false; resetBtn.textContent = 'Start over';
    reviewSection.hidden = true; exportSection.hidden = true; steps.hidden = false;
    gridHost.replaceChildren(); status.replaceChildren(); setExportStatus('');
    changesBox.replaceChildren(); changesBox.classList.remove('exported'); parsingOnly.clear();
    cellRefs = []; noteCells = []; calcRefs = []; ruleRefs = [];
  }

  /** Column 0 is the criterion name, 1 is Points, 2 and up are the rating descriptions. */
  function focusCell(r: number, c: number): void {
    const row = cellRefs[r];
    const target = row && row[c];
    if (target) target.focus();
  }

  /** Wires a text control to the model. Saves one undo step per editing session (focus to blur). */
  function bind(control: any, read: (m: RubricModel) => string, write: (m: RubricModel, v: string) => void, r: number, c: number): void {
    control.addEventListener('focus', () => { editStarted = false; });
    control.addEventListener('input', () => {
      if (!store) return;
      if (!editStarted) { store.checkpoint(); editStarted = true; }
      write(store.model, control.value);
      refresh();
    });
    control.addEventListener('keydown', (e: any) => {
      if (e.ctrlKey && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
        e.preventDefault();
        focusCell(r + (e.key === 'ArrowDown' ? 1 : -1), c);
      } else if (e.key === 'Enter' && c <= 1 && !e.ctrlKey) {
        e.preventDefault();
        focusCell(r + (e.shiftKey ? -1 : 1), c);
      }
    });
    control.value = read(store!.model);
  }

  function bindTop(input: any, write: (m: RubricModel, v: string) => void): void {
    input.addEventListener('focus', () => { editStarted = false; });
    input.addEventListener('input', () => {
      if (!store) return;
      if (!editStarted) { store.checkpoint(); editStarted = true; }
      write(store.model, input.value);
      refresh();
    });
  }
  bindTop(nameInput, (m, v) => { m.rubricName = v; });
  bindTop(expectedInput, (m, v) => { m.expectedPoints = v; });

  // ---------- rendering ----------
  function renderGrid(): void {
    gridHost.replaceChildren();
    cellRefs = []; noteCells = []; calcRefs = []; ruleRefs = [];
    if (!store) return;
    const m = store.model;
    const headRow = mk('tr', {},
      mk('th', { scope: 'col', class: 'namecol', text: 'Criterion Name' }),
      mk('th', { scope: 'col', class: 'ptscol', text: 'Points' }),
    );
    m.ratings.forEach((rt, c) => {
      const rule = mk('span', { class: 'rpts', text: ruleText(rt.name) });
      ruleRefs.push(rule);
      if (!unlocked) {
        headRow.append(mk('th', { scope: 'col', 'data-test': 'rhead-' + c }, mk('span', { class: 'rname', text: rt.name || '(no name)' }), rule));
      } else {
        const ni = mk('input', { type: 'text', class: 'keep-paste', 'aria-label': 'Rating name, column ' + (c + 1), 'data-test': 'rname-' + c });
        ni.value = rt.name;
        ni.addEventListener('focus', () => { editStarted = false; });
        ni.addEventListener('input', () => {
          if (!store) return;
          if (!editStarted) { store.checkpoint(); editStarted = true; }
          store.model.ratings[c].name = ni.value;
          refresh();
        });
        headRow.append(mk('th', { scope: 'col' }, ni, rule));
      }
    });
    headRow.append(mk('th', { scope: 'col', text: 'Checks' }), mk('th', { scope: 'col', text: 'Row' }));
    const body = mk('tbody');
    m.criteria.forEach((cr, r) => {
      const refs: any[] = [];
      const nameCtl = mk('input', { type: 'text', class: 'keep-paste', 'aria-label': 'Criterion name, row ' + (r + 1), 'data-test': 'cname-' + r });
      bind(nameCtl, (mm) => mm.criteria[r].name, (mm, v) => { mm.criteria[r].name = v; }, r, 0);
      const ptsCtl = mk('input', { type: 'text', class: 'keep-paste pts', inputmode: 'decimal', 'aria-label': 'Maximum points for row ' + (r + 1), 'data-test': 'cpts-' + r });
      bind(ptsCtl, (mm) => mm.criteria[r].points, (mm, v) => { mm.criteria[r].points = v; }, r, 1);
      const calc = mk('span', { class: 'calc', 'data-test': 'calc-' + r });
      calcRefs.push(calc);
      refs.push(nameCtl, ptsCtl);
      const tr = mk('tr', {}, mk('th', { scope: 'row', class: 'namecol' }, nameCtl), mk('td', { class: 'ptscol' }, ptsCtl, calc));
      m.ratings.forEach((rt, c) => {
        const ta = mk('textarea', { rows: 4, class: 'keep-paste', 'aria-label': 'Row ' + (r + 1) + ', ' + (rt.name || 'column ' + (c + 1)) + ' description', 'data-test': 'cell-' + r + '-' + c });
        bind(ta, (mm) => mm.criteria[r].cells[c] ?? '', (mm, v) => { mm.criteria[r].cells[c] = v; }, r, c + 2);
        refs.push(ta);
        tr.append(mk('td', {}, ta));
      });
      const notes = mk('td', { class: 'notes', 'data-test': 'notes-' + r });
      noteCells.push(notes);
      tr.append(notes);
      const btn = (label: string, aria: string, fn: () => void, test: string) => mk('button', { type: 'button', class: 'btn tiny', 'aria-label': aria, title: aria, 'data-test': test + '-' + r, onclick: fn }, label);
      tr.append(mk('td', { class: 'rowtools' },
        btn('Up', 'Move row ' + (r + 1) + ' up', () => structural((mm) => moveCriterion(mm, r, -1)), 'up'),
        btn('Down', 'Move row ' + (r + 1) + ' down', () => structural((mm) => moveCriterion(mm, r, 1)), 'down'),
        btn('Add below', 'Add a criterion below row ' + (r + 1), () => structural((mm) => { addCriterion(mm, r + 1); return r + 1; }), 'add'),
        btn('Remove', 'Remove row ' + (r + 1), () => structural((mm) => { removeCriterion(mm, r); return Math.min(r, mm.criteria.length - 1); }), 'remove'),
      ));
      cellRefs.push(refs);
      body.append(tr);
    });
    gridHost.append(mk('table', { class: 'grid' }, mk('caption', { class: 'sr-only', text: 'Rubric criteria with maximum points and one description per rating level' }), mk('thead', {}, headRow), body));
    if (m.criteria.length === 0) gridHost.append(mk('p', { class: 'small pad', text: 'There are no criteria. Use "Add criterion" to add one.' }));
  }

  function issueItem(i: Issue, withRow = false): any {
    return mk('li', { class: 'issue ' + i.severity }, mk('span', { class: 'tag ' + i.severity, text: SEV_LABEL[i.severity] }), ' ', (withRow && i.row !== undefined ? 'Row ' + (i.row + 1) + ': ' : '') + i.message);
  }

  /** Edits made in the app that the ADS has not marked "parsing only". */
  function editsToFix(): { all: ReturnType<typeof diffModels>; open: ReturnType<typeof diffModels> } {
    if (!store) return { all: [], open: [] };
    const all = diffModels(store.baseline, store.model);
    return { all, open: all.filter((e) => !parsingOnly.has(e.key)) };
  }

  function renderChanges(issues: Issue[]): void {
    parsingRefs.clear();
    if (!store) { changesBox.replaceChildren(); return; }
    const { all, open } = editsToFix();
    const warns = issues.filter((i) => i.severity === 'warn');

    const editList = all.length
      ? mk('ul', { class: 'changelist', 'data-test': 'edit-list' }, ...all.map((e, idx) => {
          const id = 'po-' + idx;
          const cb = mk('input', { type: 'checkbox', id, checked: parsingOnly.has(e.key), 'data-test': 'po-' + e.key });
          cb.addEventListener('change', () => {
            if (cb.checked) parsingOnly.add(e.key); else parsingOnly.delete(e.key);
            refresh();
            const again = parsingRefs.get(e.key);
            if (again) again.focus();
          });
          parsingRefs.set(e.key, cb);
          return mk('li', { class: 'change' + (parsingOnly.has(e.key) ? ' done' : '') },
            mk('span', { class: 'changetext', text: e.message }),
            mk('label', { class: 'check po', for: id }, cb, ' Parsing fix only (the RBA is fine)'));
        }))
      : null;

    const headline = all.length === 0
      ? mk('p', { class: 'okline', 'data-test': 'edit-summary', text: 'No edits were made in the app, so the CSV matches the RBA as loaded.' })
      : open.length === 0
        ? mk('p', { class: 'okline', 'data-test': 'edit-summary', text: 'All ' + all.length + ' ' + plural(all.length, 'edit is', 'edits are') + ' marked as parsing fixes. Nothing to change in the RBA.' })
        : mk('p', { class: 'needs', 'data-test': 'edit-summary', text: open.length + ' ' + plural(open.length, 'edit', 'edits') + ' made in the app ' + plural(open.length, 'is', 'are') + ' not in the RBA yet. Make the same ' + plural(open.length, 'fix', 'fixes') + ' there so it matches the CSV.' });

    const warnList = warns.length
      ? mk('div', { class: 'stillflagged' },
          mk('h4', { text: 'Still flagged, and not fixed in the app' }),
          mk('ul', { class: 'issuelist', 'data-test': 'warn-list' }, ...warns.map((i) => issueItem(i, true))))
      : mk('p', { class: 'okline', text: 'No open warnings.' });

    changesBox.replaceChildren(...compact([
      mk('h3', { text: 'Before you hand off: keep the RBA in sync' }),
      mk('p', { class: 'small', text: 'The CSV is built from what is on screen. If you changed anything here, the RBA still has the old version. Fix real errors in the RBA so the next export matches. If an edit only fixed how the app read the file, tick "Parsing fix only".' }),
      headline, editList, warnList,
      mk('button', { type: 'button', class: 'btn', 'data-test': 'copy-notes', onclick: () => doCopyNotes() }, 'Copy these notes'),
    ]));
  }

  function refresh(): void {
    if (!store) return;
    const m = store.model;
    const issues = validate(m);
    const blocks = blockers(issues);

    if (nameInput.value !== m.rubricName) nameInput.value = m.rubricName;
    if (expectedInput.value !== m.expectedPoints) expectedInput.value = m.expectedPoints;
    if (rangeBox.checked !== m.enableRange) rangeBox.checked = m.enableRange;
    const blank = m.rubricName.trim() === '';
    nameInput.setAttribute('aria-invalid', blank ? 'true' : 'false');
    nameHelp.textContent = blank
      ? 'Required before export. It normally comes from "Assessment Name (learner facing)" in the Metadata table, which was blank or missing. Type the name here.'
      : 'Taken from "Assessment Name (learner facing)" in the Metadata table. You can edit it.';

    // Summary
    const total = totalPoints(m);
    const n = (s: string) => issues.filter((i) => i.severity === s).length;
    const rows: [string, string][] = [
      ['Source', m.sourceLabel],
      ['Rubric name', blank ? 'Missing' : m.rubricName.trim()],
      ['Criteria', String(m.criteria.length)],
      ['Rating levels', m.ratings.map((r) => r.name || '(no name)').join(', ') || 'None'],
      ['Total points', total === null ? 'Cannot be added up yet' : formatNumber(total) + ' (expected ' + formatNumber(expectedTotal(m)) + ')'],
      ['Checks', n('block') + ' blocking, ' + n('warn') + ' warnings, ' + n('info') + ' notes'],
    ];
    summary.replaceChildren(...rows.map(([k, v]) => mk('div', { class: 'srow' }, mk('dt', { text: k }), mk('dd', { text: v }))));

    // Issues
    const global = issues.filter((i) => i.row === undefined);
    const important = global.filter((i) => i.severity !== 'info');
    const notes = global.filter((i) => i.severity === 'info');
    issuesBox.replaceChildren(...compact([
      important.length
        ? mk('ul', { class: 'issuelist' }, ...important.map((i) => issueItem(i)))
        : mk('p', { class: 'okline', text: 'No document-level problems found.' }),
      notes.length ? mk('details', { class: 'notesbox' }, mk('summary', { text: notes.length + ' note' + (notes.length === 1 ? '' : 's') }), mk('ul', { class: 'issuelist' }, ...notes.map((i) => issueItem(i)))) : null,
    ]));

    // Row notes, computed scores, rating rules
    noteCells.forEach((cell, r) => {
      const mine = issues.filter((i) => i.row === r);
      cell.replaceChildren(...(mine.length ? [mk('ul', { class: 'issuelist small' }, ...mine.map((i) => issueItem(i)))] : []));
    });
    calcRefs.forEach((el, r) => {
      const c = m.criteria[r];
      const sc = c ? criterionScores(m.ratings, c.points) : [];
      el.textContent = sc.length && sc.every((x) => x !== null) ? 'Scores: ' + sc.join(' / ') : '';
    });
    ruleRefs.forEach((el, c) => { if (m.ratings[c]) el.textContent = ruleText(m.ratings[c].name); });

    // Toolbar and export
    undoBtn.disabled = !store.canUndo;
    redoBtn.disabled = !store.canRedo;
    downloadBtn.disabled = blocks.length > 0;
    copyBtn.disabled = blocks.length > 0;
    fileNote.textContent = blocks.length ? '' : 'Saves as: ' + csvFileName(m.rubricName);
    exportReason.replaceChildren(...(blocks.length ? compact([
      mk('p', { class: 'msg error', text: 'Export is turned off until these are fixed:' }),
      mk('ul', { class: 'issuelist' }, ...blocks.map((i) => issueItem(i, true))),
      blank ? mk('button', { type: 'button', class: 'btn', 'data-test': 'goto-name', onclick: () => nameInput.focus() }, 'Go to the rubric name field') : null,
    ]) : []));
    preview.textContent = buildCsv(m);
    renderChanges(issues);
  }

  // ---------- export ----------
  function currentCsv(): string | null {
    if (!store) return null;
    if (blockers(validate(store.model)).length) { setExportStatus('Fix the items listed above first.', 'error'); return null; }
    return buildCsv(store.model);
  }

  /** The line shown after an export: what happened, and whether the RBA needs updating. */
  function afterExport(done: string): void {
    const { all, open } = editsToFix();
    if (open.length > 0) {
      setExportStatus(done + ' ' + open.length + ' ' + plural(open.length, 'edit', 'edits') + ' made in the app ' + plural(open.length, 'is', 'are') + ' not in the RBA yet. Update the RBA using the list below.', 'warn');
      changesBox.classList.add('exported');
      if (typeof changesBox.scrollIntoView === 'function') changesBox.scrollIntoView({ block: 'nearest' });
    } else if (all.length > 0) {
      setExportStatus(done + ' All edits are marked as parsing fixes, so the RBA needs no changes.', 'success');
    } else {
      setExportStatus(done + ' No edits were made in the app, so the CSV matches the RBA.', 'success');
    }
  }

  function downloadBlob(blob: Blob, name: string): void {
    const url = URL.createObjectURL(blob);
    const a: any = mk('a', { href: url, download: name });
    root.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  function copyText(text: string, onDone: () => void, onFail: () => void): void {
    const fallback = () => {
      try {
        const ta: any = mk('textarea', { class: 'sr-only', 'aria-hidden': 'true' });
        ta.value = text;
        root.append(ta);
        ta.select();
        const ok = typeof d.execCommand === 'function' && d.execCommand('copy');
        ta.remove();
        if (ok) onDone(); else onFail();
      } catch { onFail(); }
    };
    const nav: any = globalThis.navigator;
    if (nav && nav.clipboard && nav.clipboard.writeText) nav.clipboard.writeText(text).then(onDone, fallback);
    else fallback();
  }

  function doDownload(): void {
    const csv = currentCsv();
    if (csv === null || !store) return;
    const name = csvFileName(store.model.rubricName);
    downloadBlob(new Blob([csv], { type: 'text/csv;charset=utf-8' }), name);
    afterExport('Downloaded ' + name + ' (' + store.model.criteria.length + ' ' + plural(store.model.criteria.length, 'criterion', 'criteria') + ').');
  }

  function doCopy(): void {
    const csv = currentCsv();
    if (csv === null) return;
    copyText(csv, () => afterExport('Copied the CSV to the clipboard.'), () => setExportStatus('Your browser blocked copying. Open the preview below, select the text and copy it, or use Download.', 'error'));
  }

  function doCopyNotes(): void {
    if (!store) return;
    const { all } = editsToFix();
    const text = formatNotes({
      rubricName: store.model.rubricName,
      source: store.model.sourceLabel,
      edits: all,
      parsingOnly,
      warnings: validate(store.model).filter((i) => i.severity === 'warn'),
    });
    copyText(text, () => setExportStatus('Copied the RBA update notes. Paste them into a ticket, an email or the RBA\'s Development Notes.', 'success'), () => setExportStatus('Your browser blocked copying the notes. Select the list above and copy it by hand.', 'error'));
  }
}
