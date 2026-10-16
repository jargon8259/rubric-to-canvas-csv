export type Severity = 'block' | 'warn' | 'info';

export interface Issue {
  id: string;
  severity: Severity;
  message: string;
  /** Zero-based criterion row this issue belongs to, if any. */
  row?: number;
  /** Zero-based rating column this issue belongs to, if any. */
  col?: number;
  field?: 'rubricName';
}

export interface Rating {
  name: string;
}

export interface Criterion {
  /** Internal id so edits can be tracked even when rows are renamed or reordered. Never exported. */
  id: string;
  name: string;
  /** Maximum points for this criterion, kept as text so a bad value can be shown and fixed. */
  points: string;
  /** One description per rating column, same order as `ratings`. */
  cells: string[];
}

export interface RubricModel {
  rubricName: string;
  /** What the criterion points should add up to. From Metadata "Assessment Points"; 100 if not found. Kept as text. */
  expectedPoints: string;
  enableRange: boolean;
  ratings: Rating[];
  criteria: Criterion[];
  /** Issues found while reading the source. Not recomputed on edit. */
  loadIssues: Issue[];
  sourceLabel: string;
}

export type Block =
  | { kind: 'heading'; level: number; text: string }
  | { kind: 'text'; text: string }
  | { kind: 'table'; rows: string[][]; merged: boolean; nested: boolean };

export interface ParsedDoc {
  blocks: Block[];
  kind: 'docx' | 'html' | 'text' | 'delimited';
  /** Total comments in the file (docx only). */
  comments: number;
  /** Comments not marked resolved (docx only). */
  openComments: number;
  trackedInserts: number;
  trackedDeletes: number;
}
