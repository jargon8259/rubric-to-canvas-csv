import { norm } from './text.ts';
import type { Rating } from './types.ts';

/** What the criterion points should add up to when the RBA doesn't say otherwise. */
export const DEFAULT_TOTAL_POINTS = 100;

export type Rule = 'full' | 'half' | 'zero';

/**
 * How each rating level is scored, from the criterion's maximum points:
 * Competent = full points, Approaching Competency = half rounded up, Not Competent and Not Submitted = 0.
 * Change the scoring here and nowhere else.
 */
const RULES = new Map<string, Rule>([
  ['competent', 'full'],
  ['approaching competency', 'half'],
  ['not competent', 'zero'],
  ['not submitted', 'zero'],
]);

export const STANDARD_RATING_NAMES = ['Competent', 'Approaching Competency', 'Not Competent', 'Not Submitted'];

export function ruleFor(name: string): Rule | null {
  return RULES.get(norm(name)) ?? null;
}

/** Strict numeric parse for points. Returns null for blanks, text and negative numbers. */
export function parsePoints(s: string): number | null {
  const t = s.trim();
  return /^\d+(\.\d+)?$/.test(t) ? Number(t) : null;
}

export function scoreFor(rule: Rule, max: number): number {
  if (rule === 'full') return max;
  if (rule === 'half') return Math.ceil(max / 2);
  return 0;
}

/** Points for each rating of one criterion. An entry is null if the criterion points or the rating name can't be scored. */
export function criterionScores(ratings: Rating[], pointsText: string): (number | null)[] {
  const max = parsePoints(pointsText);
  return ratings.map((r) => {
    const rule = ruleFor(r.name);
    return max === null || rule === null ? null : scoreFor(rule, max);
  });
}

export function formatNumber(n: number): string {
  return String(Math.round(n * 1e6) / 1e6);
}
