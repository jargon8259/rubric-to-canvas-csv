import type { Issue, RubricModel } from './types.ts';
import { norm } from './text.ts';
import { DEFAULT_TOTAL_POINTS, parsePoints, ruleFor, formatNumber } from './scoring.ts';

const PLACEHOLDER_NAME = /^#\.|title of sub-prompt|^\[.*\]$/i;
const PLACEHOLDER_TEXT = /goes here/i;
const LONG = 1000;
const CANVAS_TITLE_LIMIT = 255;

/** Sum of the criterion points. Null if there are no criteria or any points value is not a number. */
export function totalPoints(m: RubricModel): number | null {
  if (m.criteria.length === 0) return null;
  let sum = 0;
  for (const c of m.criteria) {
    const p = parsePoints(c.points);
    if (p === null) return null;
    sum += p;
  }
  return Math.round(sum * 1e6) / 1e6;
}

/** The total the criterion points should reach: the RBA's Assessment Points, or 100 if that is blank or not a number. */
export function expectedTotal(m: RubricModel): number {
  return parsePoints(m.expectedPoints) ?? DEFAULT_TOTAL_POINTS;
}

const ORDER = { block: 0, warn: 1, info: 2 } as const;

export function validate(m: RubricModel): Issue[] {
  const out: Issue[] = [...m.loadIssues];
  const add = (i: Issue) => out.push(i);

  if (m.rubricName.trim() === '') {
    add({ id: 'name-blank', severity: 'block', field: 'rubricName', message: 'Add a rubric name to enable export.' });
  } else {
    if (m.rubricName.trim().length > CANVAS_TITLE_LIMIT) add({ id: 'name-long', severity: 'warn', field: 'rubricName', message: `The rubric name is over ${CANVAS_TITLE_LIMIT} characters. Canvas may reject or cut it. (Limit not verified.)` });
    if (PLACEHOLDER_NAME.test(m.rubricName.trim())) add({ id: 'name-placeholder', severity: 'warn', field: 'rubricName', message: 'The rubric name looks like template placeholder text.' });
  }
  if (m.criteria.length === 0) add({ id: 'no-criteria', severity: 'block', message: 'No criteria to export. Add at least one criterion row.' });
  if (m.ratings.length === 0) add({ id: 'no-ratings', severity: 'block', message: 'No rating columns found. The rubric needs at least one rating level.' });

  m.ratings.forEach((r, c) => {
    if (r.name.trim() === '') add({ id: `rname-${c}`, severity: 'warn', col: c, message: `Rating column ${c + 1} has no name.` });
    if (ruleFor(r.name) === null) {
      add({ id: `rrule-${c}`, severity: 'block', col: c, message: `I can't work out points for the rating "${r.name || '(blank)'}". Use the standard names: Competent, Approaching Competency, Not Competent or Not Submitted. (Tick "Edit rating names" to fix it here.)` });
    }
  });

  const seen = new Map<string, number>();
  let emptyCells = 0;
  m.criteria.forEach((c, row) => {
    const label = c.name.trim();
    if (label === '') add({ id: `cname-${row}`, severity: 'block', row, message: 'This criterion has no name. Canvas needs one.' });
    else {
      const key = norm(label);
      if (seen.has(key)) add({ id: `dup-${row}`, severity: 'warn', row, message: `Same name as row ${seen.get(key)! + 1}.` });
      else seen.set(key, row);
      if (PLACEHOLDER_NAME.test(label)) add({ id: `ph-${row}`, severity: 'warn', row, message: 'Looks like template placeholder text.' });
      if (label.length > CANVAS_TITLE_LIMIT) add({ id: `clong-${row}`, severity: 'warn', row, message: `Name is over ${CANVAS_TITLE_LIMIT} characters. Canvas may cut it. (Limit not verified.)` });
    }

    const pts = c.points.trim();
    const max = parsePoints(pts);
    if (pts === '') add({ id: `cpts-${row}`, severity: 'block', row, message: 'This criterion has no points. Enter its maximum points.' });
    else if (max === null) add({ id: `cpts-${row}`, severity: 'block', row, message: `Points must be a number, 0 or higher (for example 15). Found "${c.points}".` });
    else {
      if (max === 0) add({ id: `czero-${row}`, severity: 'warn', row, message: 'This criterion is worth 0 points.' });
      if (!Number.isInteger(max)) add({ id: `cdec-${row}`, severity: 'warn', row, message: 'Points are not a whole number. "Half, rounded up" is meant for whole numbers, so check the Approaching Competency points.' });
      else if (max === 1 && m.ratings.some((r) => ruleFor(r.name) === 'half')) add({ id: `chalf-${row}`, severity: 'info', row, message: '1 point rounds up to 1 for Approaching Competency, the same as Competent.' });
    }

    if (c.cells.some((x) => PLACEHOLDER_TEXT.test(x))) add({ id: `phc-${row}`, severity: 'warn', row, message: 'A description still contains template text ("goes here").' });
    if (c.cells.some((x) => x.length > LONG)) add({ id: `long-${row}`, severity: 'info', row, message: `A description is over ${LONG} characters.` });
    emptyCells += c.cells.filter((x) => x.trim() === '').length;
  });
  const totalCells = m.criteria.length * m.ratings.length;
  if (emptyCells > 0) add({ id: 'empty', severity: 'info', message: `${emptyCells} of ${totalCells} rating descriptions are empty. That is allowed.` });

  if (m.expectedPoints.trim() !== '' && parsePoints(m.expectedPoints) === null) {
    add({ id: 'ap-invalid', severity: 'warn', message: `The expected total ("${m.expectedPoints}") is not a number, so I checked against ${DEFAULT_TOTAL_POINTS}.` });
  }
  const total = totalPoints(m);
  const expected = expectedTotal(m);
  if (total !== null && total !== expected) {
    add({ id: 'ap-mismatch', severity: 'warn', message: `The criterion points add up to ${formatNumber(total)}, but the expected total (Assessment Points) is ${formatNumber(expected)}.` });
  }

  return out.sort((a, b) => ORDER[a.severity] - ORDER[b.severity]);
}

export function blockers(issues: Issue[]): Issue[] {
  return issues.filter((i) => i.severity === 'block');
}
