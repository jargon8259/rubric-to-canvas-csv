import { table, h, p } from './docxbuilder.ts';

export const RUBRIC_HEAD = ['Criterion Name', 'Points', 'Competent', 'Approaching Competency', 'Not Competent', 'Not Submitted'];

export function metaRows(over: Record<string, string> = {}): string[][] {
  const base: Record<string, string> = {
    'Course Code': 'SAMPLE101', 'Assessment Name (learner facing)': 'Sample Assessment', 'Assessment Points': '100',
  };
  return [['Field', 'Value'], ...Object.entries({ ...base, ...over })];
}

/** Two criteria worth 60 + 40 = 100. */
export function criteriaRows(): string[][] {
  return [
    ['A. First Prompt', '60', 'Meets all of A.', 'Meets some of A.', 'Does not meet A.', 'No submission.'],
    ['B. Second Prompt', '40', 'Meets all of B.', 'Meets some of B.', 'Does not meet B.', 'No submission.'],
  ];
}

export function rbaBody(opts: { meta?: Record<string, string>; metaRows?: string[][]; criteria?: string[][]; head?: string[] } = {}): string {
  return [
    h(1, 'Metadata'), table(opts.metaRows ?? metaRows(opts.meta)),
    h(1, 'Rubric'),
    table([opts.head ?? RUBRIC_HEAD, ...(opts.criteria ?? criteriaRows())]),
    h(1, 'Approval Record'), table([['Role', 'Name', 'Date of Approval'], ['SME', '', '']]),
    p('Template updated 2026-10-16'),
  ].join('');
}
