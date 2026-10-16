import type { RubricModel } from './types.ts';
import { toCsv } from './csv.ts';
import { criterionScores, formatNumber } from './scoring.ts';

const FIXED = ['Rubric Name', 'Criteria Name', 'Criteria Description', 'Criteria Enable Range'];
const GROUP = ['Rating Name', 'Rating Description', 'Rating Points'];

export function buildCsvRows(m: RubricModel): string[][] {
  const header = [...FIXED];
  for (let i = 0; i < m.ratings.length; i++) header.push(...GROUP);
  const rows = [header];
  for (const c of m.criteria) {
    const scores = criterionScores(m.ratings, c.points);
    const row = [m.rubricName.trim(), c.name.trim(), '', m.enableRange ? 'true' : 'false'];
    m.ratings.forEach((r, i) => row.push(r.name.trim(), c.cells[i] ?? '', scores[i] === null ? '' : formatNumber(scores[i] as number)));
    rows.push(row);
  }
  return rows;
}

export function buildCsv(m: RubricModel): string {
  return toCsv(buildCsvRows(m));
}

/** A safe file name built from the rubric name. */
export function csvFileName(rubricName: string): string {
  const base = rubricName
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/, '')
    .slice(0, 100)
    .trim();
  return (base || 'rubric') + '.csv';
}
