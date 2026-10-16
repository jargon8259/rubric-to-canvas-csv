export function detectDelimiter(text: string): ',' | '\t' {
  const head = text.split(/\r?\n/).slice(0, 5).join('\n');
  return head.includes('\t') ? '\t' : ',';
}

export interface DelimitedResult {
  rows: string[][];
  /** True if a quoted field was never closed. */
  unterminatedQuote: boolean;
}

/** Standard CSV/TSV reader: quotes, doubled quotes, embedded newlines. Fully blank rows are dropped. */
export function parseDelimited(input: string, delimiter: string): DelimitedResult {
  const text = input.charCodeAt(0) === 0xfeff ? input.slice(1) : input;
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  let fieldStarted = false;
  let unterminated = false;

  const endField = () => { row.push(field); field = ''; fieldStarted = false; };
  const endRow = () => {
    endField();
    if (!row.every((c) => c === '')) rows.push(row);
    row = [];
  };

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; } else { inQuotes = false; }
      } else {
        field += ch;
      }
      continue;
    }
    if (ch === '"' && !fieldStarted) { inQuotes = true; fieldStarted = true; continue; }
    if (ch === delimiter) { endField(); continue; }
    if (ch === '\r') { if (text[i + 1] === '\n') i++; endRow(); continue; }
    if (ch === '\n') { endRow(); continue; }
    field += ch;
    fieldStarted = true;
  }
  if (inQuotes) unterminated = true;
  if (field !== '' || row.length > 0) endRow();
  return { rows: rows.map((r) => r.map((c) => c.replace(/\r\n?/g, '\n'))), unterminatedQuote: unterminated };
}

export function escapeField(value: string): string {
  const s = value.replace(/\r\n?/g, '\n');
  return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

/** Rows to CSV text: \n line endings, trailing newline, no BOM. */
export function toCsv(rows: string[][]): string {
  return rows.map((r) => r.map(escapeField).join(',')).join('\n') + '\n';
}
