/** A1-notation helpers. All row/col indices in this app are 0-based internally. */

export interface CellRef {
  row: number;
  col: number;
}

export interface RangeRef {
  start: CellRef;
  end: CellRef;
}

export function colToLetter(col: number): string {
  let n = col + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

export function letterToCol(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

export function encodeCell(ref: CellRef): string {
  return `${colToLetter(ref.col)}${ref.row + 1}`;
}

export function decodeCell(a1: string): CellRef {
  const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(a1.trim());
  if (!m) throw new Error(`Invalid cell reference: "${a1}"`);
  return { col: letterToCol(m[1]), row: parseInt(m[2], 10) - 1 };
}

export function encodeRange(range: RangeRef): string {
  return `${encodeCell(range.start)}:${encodeCell(range.end)}`;
}

export function decodeRange(a1: string): RangeRef {
  const parts = a1.trim().split(':');
  if (parts.length === 1) {
    const ref = decodeCell(parts[0]);
    return { start: ref, end: ref };
  }
  if (parts.length !== 2) throw new Error(`Invalid range: "${a1}"`);
  return normalizeRange({ start: decodeCell(parts[0]), end: decodeCell(parts[1]) });
}

/** Reorders a range so start is always the top-left corner. */
export function normalizeRange(range: RangeRef): RangeRef {
  return {
    start: {
      row: Math.min(range.start.row, range.end.row),
      col: Math.min(range.start.col, range.end.col),
    },
    end: {
      row: Math.max(range.start.row, range.end.row),
      col: Math.max(range.start.col, range.end.col),
    },
  };
}

export function rangeHeight(range: RangeRef): number {
  return range.end.row - range.start.row + 1;
}

export function rangeWidth(range: RangeRef): number {
  return range.end.col - range.start.col + 1;
}

export function isInRange(ref: CellRef, range: RangeRef): boolean {
  return (
    ref.row >= range.start.row &&
    ref.row <= range.end.row &&
    ref.col >= range.start.col &&
    ref.col <= range.end.col
  );
}

/** "Sheet1!B2:D5", quoting the sheet name only when it needs it. */
export function qualifyRange(sheet: string, range: string): string {
  const needsQuotes = /[^A-Za-z0-9_]/.test(sheet);
  return `${needsQuotes ? `'${sheet.replace(/'/g, "''")}'` : sheet}!${range}`;
}
