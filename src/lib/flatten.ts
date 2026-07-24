import type { CellValue, ExtractionResult, FieldResult } from './types';

export interface FlatTable {
  headers: string[];
  /**
   * One row per extraction. Kept as a matrix rather than a single row so several
   * files can be stacked into one table later without reshaping anything.
   */
  rows: CellValue[][];
}

/**
 * Collapses a field to one cell.
 *
 * The shape is decided by the field's declared output type, never by how much
 * data this particular file happened to contain — a column that is a bare
 * number in one file and a JSON array in the next is unusable downstream.
 */
export function flattenFieldValue(field: FieldResult): CellValue {
  if (!field.ok) return null;
  if (field.output === 'value') return field.value ?? null;

  if (field.output === 'list') {
    return JSON.stringify(field.rows.map((row) => row[0] ?? null));
  }

  // table: objects when the field captured headers, plain arrays otherwise.
  const data = field.headers
    ? field.rows.map((row) =>
        Object.fromEntries(field.headers!.map((h, i) => [h, row[i] ?? null])),
      )
    : field.rows;

  return JSON.stringify(data);
}

export function flattenResult(result: ExtractionResult): FlatTable {
  return {
    headers: result.fields.map((f) => f.fieldName),
    rows: [result.fields.map(flattenFieldValue)],
  };
}

/** True when a cell holds an embedded JSON document rather than a plain value. */
export function isJsonCell(field: FieldResult): boolean {
  return field.ok && field.output !== 'value';
}
