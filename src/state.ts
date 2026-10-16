import type { RubricModel } from './types.ts';

const MAX_HISTORY = 200;
let criterionCounter = 0;

/** A new internal id for a criterion row. Used to track edits across renames and reordering. */
export function newCriterionId(): string {
  criterionCounter += 1;
  return 'c' + criterionCounter;
}

/**
 * Holds the current rubric with simple undo and redo, plus a snapshot of the rubric exactly as it was
 * loaded (the baseline) so the app can report what was changed afterwards.
 * Call checkpoint() just before each change.
 */
export class Store {
  model: RubricModel;
  readonly baseline: RubricModel;
  private undoStack: string[] = [];
  private redoStack: string[] = [];

  constructor(model: RubricModel) {
    this.model = model;
    this.baseline = JSON.parse(JSON.stringify(model));
  }

  checkpoint(): void {
    this.undoStack.push(JSON.stringify(this.model));
    if (this.undoStack.length > MAX_HISTORY) this.undoStack.shift();
    this.redoStack = [];
  }
  get canUndo(): boolean { return this.undoStack.length > 0; }
  get canRedo(): boolean { return this.redoStack.length > 0; }
  undo(): boolean {
    const prev = this.undoStack.pop();
    if (prev === undefined) return false;
    this.redoStack.push(JSON.stringify(this.model));
    this.model = JSON.parse(prev);
    return true;
  }
  redo(): boolean {
    const next = this.redoStack.pop();
    if (next === undefined) return false;
    this.undoStack.push(JSON.stringify(this.model));
    this.model = JSON.parse(next);
    return true;
  }
}

export function addCriterion(m: RubricModel, at: number): void {
  const i = Math.max(0, Math.min(at, m.criteria.length));
  m.criteria.splice(i, 0, { id: newCriterionId(), name: '', points: '', cells: m.ratings.map(() => '') });
}
export function removeCriterion(m: RubricModel, at: number): void {
  if (at >= 0 && at < m.criteria.length) m.criteria.splice(at, 1);
}
export function moveCriterion(m: RubricModel, at: number, delta: -1 | 1): number {
  const to = at + delta;
  if (at < 0 || at >= m.criteria.length || to < 0 || to >= m.criteria.length) return at;
  const [c] = m.criteria.splice(at, 1);
  m.criteria.splice(to, 0, c);
  return to;
}
