import type { CellValue, ExtractionResult, FieldResult } from './types';

export const FILE_COLUMN = 'File';

export interface FlatTable {
  headers: string[];
  /** One row per extracted file. */
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

/**
 * One row per file, one column per template field, prefixed with the file name.
 *
 * Columns come from the first result rather than the union of all of them: every
 * file in a batch is extracted with the same template, so the field list is
 * identical across results and the table stays rectangular.
 */
export function flattenResults(results: ExtractionResult[]): FlatTable {
  const headers = [FILE_COLUMN, ...(results[0]?.fields ?? []).map((f) => f.fieldName)];

  return {
    headers,
    rows: results.map((result) => [
      result.fileName,
      ...result.fields.map(flattenFieldValue),
    ]),
  };
}

export function flattenResult(result: ExtractionResult): FlatTable {
  return flattenResults([result]);
}

/** True when a cell holds an embedded JSON document rather than a plain value. */
export function isJsonCell(field: FieldResult): boolean {
  return field.ok && field.output !== 'value';
}
