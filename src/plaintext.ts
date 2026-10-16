import type { Block } from './types.ts';
import { parseDelimited, detectDelimiter } from './csv.ts';
import { cleanCell } from './text.ts';
import { classifyHeader } from './rba.ts';

export interface PlainResult { blocks: Block[]; usedFallback: boolean }

/**
 * Reads pasted plain text (tab-separated from Word or Excel, or comma-separated) into blocks.
 * Rows with two or more fields become tables; single-field rows become text.
 */
export function textToBlocks(text: string): PlainResult {
  const delim = detectDelimiter(text);
  let usedFallback = false;
  let rows: string[][];
  const first = parseDelimited(text, delim);
  if (first.unterminatedQuote) {
    // A stray quote mark swallowed the rest. Fall back to a simple split with no quote handling.
    usedFallback = true;
    rows = text.replace(/^\ufeff/, '').split(/\r\n?|\n/).filter((l) => l.trim() !== '').map((l) => l.split(delim));
  } else {
    rows = first.rows;
  }
  const blocks: Block[] = [];
  let table: string[][] | null = null;
  for (const r of rows) {
    const cells = r.map((c) => cleanCell(c));
    if (cells.length >= 2) {
      // A known header row (Field/Value, Criterion Name, Role/Name) starts a new table.
      if (table && classifyHeader(cells) !== 'other') table = null;
      if (!table) { table = []; blocks.push({ kind: 'table', rows: table, merged: false, nested: false }); }
      table.push(cells);
    } else {
      table = null;
      if (cells[0]) blocks.push({ kind: 'text', text: cells[0] });
    }
  }
  return { blocks, usedFallback };
}
