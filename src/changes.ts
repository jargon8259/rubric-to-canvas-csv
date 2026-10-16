import type { Issue, RubricModel } from './types.ts';
import { squash } from './text.ts';

/** One difference between the rubric as loaded from the RBA and the rubric as it is on screen now. */
export interface Change {
  /** Stable key, used to remember "parsing only" ticks. */
  key: string;
  message: string;
}

function shorten(s: string, n = 60): string {
  const one = squash(s);
  return one.length > n ? one.slice(0, n - 1) + '\u2026' : one;
}
const quoted = (s: string) => '"' + shorten(s) + '"';
const shownText = (s: string) => (s.trim() === '' ? '(blank)' : quoted(s));
const sameText = (a: string, b: string) => a.replace(/\r\n?/g, '\n').trim() === b.replace(/\r\n?/g, '\n').trim();

/**
 * Compares the rubric as loaded (baseline) with the rubric now. It compares end states, not clicks,
 * so an edit that was typed and then undone or retyped to the original does not show up.
 * Export-only settings (Criteria Enable Range) are not in the RBA, so they are not reported.
 */
export function diffModels(base: RubricModel, cur: RubricModel): Change[] {
  const out: Change[] = [];

  if (!sameText(base.rubricName, cur.rubricName)) {
    out.push({
      key: 'name',
      message: base.rubricName.trim() === ''
        ? `Rubric name typed in the app: ${quoted(cur.rubricName)}. Enter it as the Assessment Name (learner facing) in the Metadata table.`
        : `Rubric name changed from ${quoted(base.rubricName)} to ${shownText(cur.rubricName)}. The rubric name comes from the Assessment Name (learner facing) in Metadata, so update it there.`,
    });
  }
  if (!sameText(base.expectedPoints, cur.expectedPoints)) {
    out.push({ key: 'total', message: `Expected total points changed from ${shownText(base.expectedPoints)} to ${shownText(cur.expectedPoints)}. Update Assessment Points in the Metadata table.` });
  }
  const shared = Math.min(base.ratings.length, cur.ratings.length);
  for (let c = 0; c < shared; c++) {
    if (!sameText(base.ratings[c].name, cur.ratings[c].name)) {
      out.push({ key: 'rating:' + c, message: `Rating column ${c + 1} renamed from ${shownText(base.ratings[c].name)} to ${shownText(cur.ratings[c].name)}. Update the rubric table header.` });
    }
  }

  const baseById = new Map(base.criteria.map((c) => [c.id, c]));
  const curIds = new Set(cur.criteria.map((c) => c.id));
  for (const c of base.criteria) {
    if (!curIds.has(c.id)) out.push({ key: 'rem:' + c.id, message: `Criterion removed: ${shownText(c.name)}. Delete that row from the RBA rubric table (and its prompt, if it has none left).` });
  }
  cur.criteria.forEach((c, idx) => {
    const was = baseById.get(c.id);
    const label = c.name.trim() === '' ? 'Row ' + (idx + 1) : quoted(c.name);
    if (!was) {
      out.push({ key: 'add:' + c.id, message: `Criterion added: ${label} (row ${idx + 1}, ${c.points.trim() === '' ? 'no points' : c.points.trim() + ' points'}). Add it to the RBA rubric table.` });
      return;
    }
    if (!sameText(was.name, c.name)) out.push({ key: 'ren:' + c.id, message: `Criterion name changed from ${shownText(was.name)} to ${shownText(c.name)}. The names must match the prompts in the RBA.` });
    if (!sameText(was.points, c.points)) out.push({ key: 'pts:' + c.id, message: `${label}: points changed from ${shownText(was.points)} to ${shownText(c.points)}. Update the Points column.` });
    const cols = Math.min(was.cells.length, c.cells.length);
    for (let k = 0; k < cols; k++) {
      if (sameText(was.cells[k], c.cells[k])) continue;
      const level = cur.ratings[k] ? shownText(cur.ratings[k].name).replace(/^"|"$/g, '') : 'column ' + (k + 1);
      const what = was.cells[k].trim() === '' ? 'filled in' : c.cells[k].trim() === '' ? 'cleared' : 'edited';
      out.push({ key: `txt:${c.id}:${k}`, message: `${label}: ${level} description ${what}. Was: ${shownText(was.cells[k])} Now: ${shownText(c.cells[k])}` });
    }
  });

  const baseOrder = base.criteria.filter((c) => curIds.has(c.id)).map((c) => c.id);
  const baseIds = new Set(base.criteria.map((c) => c.id));
  const curOrder = cur.criteria.filter((c) => baseIds.has(c.id)).map((c) => c.id);
  if (baseOrder.join('|') !== curOrder.join('|')) {
    out.push({ key: 'order', message: 'Criteria order was changed. Put the rows in the same order in the RBA rubric table.' });
  }
  return out;
}

export interface NotesInput {
  rubricName: string;
  source: string;
  edits: Change[];
  parsingOnly: Set<string>;
  warnings: Issue[];
}

/** Plain-text version of the hand-off reminder, for pasting into a ticket, email or the RBA's Development Notes. */
export function formatNotes(n: NotesInput): string {
  const needs = n.edits.filter((e) => !n.parsingOnly.has(e.key));
  const parsing = n.edits.filter((e) => n.parsingOnly.has(e.key));
  const lines: string[] = [];
  lines.push('RBA update notes');
  lines.push('Rubric: ' + (n.rubricName.trim() || '(no name)'));
  lines.push('Loaded from: ' + n.source);
  lines.push('');
  lines.push('Edits made in the app. Make the same fix in the RBA so it matches the exported CSV:');
  if (needs.length === 0) lines.push('None.');
  needs.forEach((e, i) => lines.push(`${i + 1}. ${e.message}`));
  if (parsing.length) {
    lines.push('');
    lines.push('Marked as parsing only (the RBA itself is fine):');
    parsing.forEach((e) => lines.push('- ' + e.message));
  }
  lines.push('');
  lines.push('Warnings still showing when the CSV was exported (not fixed in the app):');
  if (n.warnings.length === 0) lines.push('None.');
  n.warnings.forEach((w) => lines.push('- ' + (w.row !== undefined ? `Row ${w.row + 1}: ` : '') + w.message));
  return lines.join('\n') + '\n';
}
