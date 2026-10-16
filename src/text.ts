const NAMED: Record<string, string> = {
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ',
  ldquo: '\u201c', rdquo: '\u201d', lsquo: '\u2018', rsquo: '\u2019',
  ndash: '\u2013', mdash: '\u2014', hellip: '\u2026', bull: '\u2022',
  copy: '\u00a9', reg: '\u00ae', trade: '\u2122', deg: '\u00b0',
};

export function decodeEntities(s: string): string {
  return s.replace(/&(#x[0-9a-f]+|#\d+|[a-z][a-z0-9]*);/gi, (whole, body: string) => {
    if (body[0] === '#') {
      const code = body[1].toLowerCase() === 'x' ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
      if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) return whole;
      try { return String.fromCodePoint(code); } catch { return whole; }
    }
    const v = Object.prototype.hasOwnProperty.call(NAMED, body.toLowerCase()) ? NAMED[body.toLowerCase()] : undefined;
    return v === undefined ? whole : v;
  });
}

/** Lowercase, collapse all whitespace (including non-breaking spaces), trim. */
export function norm(s: string): string {
  return s.replace(/[\s\u00a0]+/g, ' ').trim().toLowerCase();
}

/** Collapse whitespace to single spaces and trim. Keeps case. */
export function squash(s: string): string {
  return s.replace(/[\s\u00a0]+/g, ' ').trim();
}

/** Clean the text of one table cell: normalise line endings and spaces, trim each line, drop empty lines. */
export function cleanCell(s: string): string {
  return s
    .replace(/\r\n?/g, '\n')
    .replace(/\u00a0/g, ' ')
    .split('\n')
    .map((line) => line.replace(/[ \t\f\v]+/g, ' ').trim())
    .filter((line) => line !== '')
    .join('\n');
}
