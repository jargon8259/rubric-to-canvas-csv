import type { Block, Issue, ParsedDoc, RubricModel } from './types.ts';
import { norm, squash } from './text.ts';
import { STANDARD_RATING_NAMES, DEFAULT_TOTAL_POINTS, parsePoints } from './scoring.ts';
import { newCriterionId } from './state.ts';

const RUBRIC_HEADS = new Set(['criterion name', 'criteria name', 'criteria', 'criterion']);
const SCALE_HEADS = new Set(['rating name', 'rating', 'rating level']);
const POINTS_HEAD = /^(?:max(?:imum)?\s+)?(?:points?|pts\.?)(?:\s+possible)?$/;

type TableBlock = Extract<Block, { kind: 'table' }>;
export type Kind = 'metadata' | 'scale' | 'rubric' | 'approval' | 'other';

/** What kind of table a header row introduces. 'other' means it is not a known header. */
export function classifyHeader(row: string[]): Kind {
  const h0 = norm(row[0] ?? '');
  const h1 = norm(row[1] ?? '');
  if (h0 === 'field' && h1 === 'value') return 'metadata';
  if (SCALE_HEADS.has(h0) && h1 === 'points') return 'scale'; // old template only; no longer used
  if (RUBRIC_HEADS.has(h0)) return 'rubric';
  if (h0 === 'role' && h1 === 'name') return 'approval';
  return 'other';
}

/** Pulls a rating name out of a header. Also strips old-style "(10 points)" text and reports it. */
export function parseRatingHeader(raw: string): { name: string; legacyPoints: string | null } {
  const s = squash(raw);
  const m = /^(.*?)\s*[(\[]?\s*(-?\d+(?:\.\d+)?)\s*(?:points?|pts?\.?)\s*[)\]]?$/i.exec(s);
  if (m && m[1] !== '') return { name: m[1].trim(), legacyPoints: m[2] };
  return { name: s, legacyPoints: null };
}

/** Turns "20", "20 points" or " 20 " into "20". Anything else is kept as typed so it can be shown and fixed. */
export function normalizePoints(raw: string): string {
  const s = squash(raw);
  const m = /^(\d+(?:\.\d+)?)\s*(?:points?|pts?\.?)?$/i.exec(s);
  return m ? m[1] : s;
}

/**
 * The one place that knows where rubric-level fields live in the document.
 * The rubric name is the learner-facing assessment name from Metadata.
 * assessmentPoints is null when the Metadata table has no such row.
 */
export function extractMetadata(tables: TableBlock[]): { rubricName: string; assessmentPoints: string | null } {
  const map = new Map<string, string>();
  for (const t of tables) {
    for (const r of t.rows.slice(1)) {
      const key = norm(r[0] ?? '');
      if (key && !map.has(key)) map.set(key, squash(r[1] ?? ''));
    }
  }
  const ap = map.get('assessment points');
  return { rubricName: map.get('assessment name (learner facing)') ?? '', assessmentPoints: ap === undefined ? null : ap };
}

export interface Interpreted { model?: RubricModel; error?: string }

export function interpret(doc: ParsedDoc, sourceLabel: string): Interpreted {
  const tables = doc.blocks.filter((b): b is TableBlock => b.kind === 'table' && b.rows.length > 0);
  if (tables.length === 0) {
    return { error: "I couldn't find any table in that content. Load the Word RBA (.docx), paste a table copied from Word or Excel, or load a CSV with the rubric columns." };
  }
  const issues: Issue[] = [];
  const kinds = tables.map((t) => classifyHeader(t.rows[0] ?? []));
  const metaTables = tables.filter((_, i) => kinds[i] === 'metadata');
  const scaleTables = tables.filter((_, i) => kinds[i] === 'scale');
  const rubricTables = tables.filter((_, i) => kinds[i] === 'rubric');

  let rubric: TableBlock | undefined = rubricTables[0];
  let headerless = false;
  if (!rubric) {
    const others = tables.filter((_, i) => kinds[i] === 'other' && tables[i].rows[0].length >= 2);
    if (others.length === 0) {
      return { error: 'I found tables, but none has a "Criterion Name" header. The rubric table needs that as the first header cell, then "Points", then one column per rating level.' };
    }
    rubric = others[0];
    headerless = true;
    issues.push({ id: 'no-header', severity: 'warn', message: 'No header row starting with "Criterion Name" was found. I assumed the standard columns: Criterion Name, Points, Competent, Approaching Competency, Not Competent, Not Submitted (or the same without Points if there are only five columns). Check that the columns line up.' });
    if (others.length > 1) issues.push({ id: 'many-tables', severity: 'warn', message: `Found ${others.length} tables without a recognizable header. I used the first one.` });
  } else if (rubricTables.length > 1) {
    issues.push({ id: 'many-rubrics', severity: 'warn', message: `Found ${rubricTables.length} rubric tables ("Criterion Name" header). I used the first one and ignored the others.` });
  }

  const bodyRows = headerless ? rubric.rows : rubric.rows.slice(1);
  const widest = Math.max(rubric.rows[0].length, ...bodyRows.map((r) => r.length));
  const hasPoints = headerless ? widest >= 6 : POINTS_HEAD.test(norm(rubric.rows[0][1] ?? ''));
  const ratingStart = hasPoints ? 2 : 1;
  const ratingCount = headerless ? Math.min(STANDARD_RATING_NAMES.length, Math.max(widest - ratingStart, 0)) : Math.max(widest - ratingStart, 0);
  if (headerless && widest - ratingStart > STANDARD_RATING_NAMES.length) {
    issues.push({ id: 'extra-cols', severity: 'warn', message: `The table has ${widest - ratingStart} rating columns but I only assumed ${STANDARD_RATING_NAMES.length} without a header row. Extra columns were left out.` });
  }
  if (!hasPoints) {
    issues.push({ id: 'no-points-col', severity: 'warn', message: 'The rubric table has no "Points" column, so I could not read how many points each criterion is worth. Enter each criterion\'s points in the grid.' });
  }

  const ratings = [];
  for (let c = 0; c < ratingCount; c++) {
    if (headerless) { ratings.push({ name: STANDARD_RATING_NAMES[c] }); continue; }
    const p = parseRatingHeader(rubric.rows[0][ratingStart + c] ?? '');
    ratings.push({ name: p.name });
    if (p.legacyPoints !== null) {
      issues.push({ id: `legacy-${c}`, severity: 'info', col: c, message: `The "${p.name}" header included "${p.legacyPoints} points". I ignored that. Points now come from each criterion's Points column.` });
    }
  }
  if (scaleTables.length > 0) {
    issues.push({ id: 'old-scale', severity: 'info', message: 'Found a Rating Scale table (old template). I ignored it. Points now come from each criterion\'s Points column.' });
  }

  const criteria = [];
  let dropped = 0;
  for (const r of bodyRows) {
    const name = squash(r[0] ?? '');
    const points = hasPoints ? normalizePoints(r[1] ?? '') : '';
    const cells = Array.from({ length: ratingCount }, (_, i) => r[ratingStart + i] ?? '');
    if (r.slice(ratingStart + ratingCount).some((x) => x.trim() !== '')) dropped++;
    if (name === '' && points === '' && cells.every((x) => x === '')) continue;
    criteria.push({ id: newCriterionId(), name, points, cells });
  }
  if (dropped) issues.push({ id: 'dropped-cells', severity: 'warn', message: `${dropped} row(s) had text in extra columns beyond the rating levels. That text was not loaded.` });

  if (doc.kind === 'docx') {
    if (doc.openComments > 0) issues.push({ id: 'comments', severity: 'warn', message: `The file has ${doc.openComments} unresolved comment${doc.openComments === 1 ? '' : 's'}. You can continue, but check the rubric is final.` });
    else if (doc.comments > 0) issues.push({ id: 'comments-resolved', severity: 'info', message: `The file still contains ${doc.comments} resolved comment${doc.comments === 1 ? '' : 's'}. They are ignored.` });
    if (doc.trackedInserts + doc.trackedDeletes > 0) {
      issues.push({ id: 'tracked', severity: 'warn', message: `The file has unresolved tracked changes (${doc.trackedInserts} insertion${doc.trackedInserts === 1 ? '' : 's'}, ${doc.trackedDeletes} deletion${doc.trackedDeletes === 1 ? '' : 's'}). I kept inserted text and ignored deleted text. Accept or reject them in Word to be sure.` });
    }
  } else if (doc.kind === 'html') {
    issues.push({ id: 'paste-limits', severity: 'info', message: 'Pasted content cannot show comments or tracked changes. Load the .docx if you need those checked.' });
  } else if (doc.kind === 'text') {
    issues.push({ id: 'plain-paste', severity: 'warn', message: 'This was pasted as plain text, which loses multi-line cells and formatting. If descriptions look cut off, load the .docx or paste again straight from Word.' });
  }
  if (rubric.merged) issues.push({ id: 'merged', severity: 'warn', message: 'The rubric table has merged cells. I filled the covered cells with blanks, so check that the text sits in the right columns.' });
  if (rubric.nested) issues.push({ id: 'nested', severity: 'warn', message: 'The rubric table has a table inside a cell. The inner table was ignored.' });

  // Rubric name and Assessment Points come from the Metadata table. Points default to 100 if not found.
  const meta = extractMetadata(metaTables);
  let expectedPoints = String(DEFAULT_TOTAL_POINTS);
  if (metaTables.length === 0) {
    if (!headerless && doc.kind !== 'delimited') {
      issues.push({ id: 'no-meta', severity: 'info', message: `No Metadata table found, so the rubric name was not loaded and I expected ${DEFAULT_TOTAL_POINTS} points. Enter the rubric name below.` });
    }
  } else if (meta.assessmentPoints === null || meta.assessmentPoints.trim() === '') {
    issues.push({ id: 'ap-default', severity: 'info', message: `Assessment Points was not found in the Metadata table, so I expected ${DEFAULT_TOTAL_POINTS} points.` });
  } else {
    const text = normalizePoints(meta.assessmentPoints);
    const value = parsePoints(text);
    if (value === null) {
      issues.push({ id: 'ap-text', severity: 'warn', message: `Assessment Points ("${meta.assessmentPoints}") is not a number, so I expected ${DEFAULT_TOTAL_POINTS} points. Fix it in the RBA, or type the right total below.` });
    } else {
      expectedPoints = text;
      if (value !== DEFAULT_TOTAL_POINTS) issues.push({ id: 'ap-nonstandard', severity: 'info', message: `Assessment Points is ${text}, not the usual ${DEFAULT_TOTAL_POINTS}. I used ${text} for the total check.` });
    }
  }

  return {
    model: { rubricName: meta.rubricName, expectedPoints, enableRange: false, ratings, criteria, loadIssues: issues, sourceLabel },
  };
}

/** True when the first header is "Rubric Name": the file is already in Canvas's own import format. */
export function looksLikeCanvasCsv(rows: string[][]): boolean {
  return rows.length > 0 && norm(rows[0][0] ?? '') === 'rubric name' && norm(rows[0][1] ?? '') === 'criteria name';
}
