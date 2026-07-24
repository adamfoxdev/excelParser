import { describe, expect, it } from 'vitest';
import { FILE_COLUMN, flattenFieldValue, flattenResults } from './flatten';
import { resultsToFlatCsv } from './download';
import type { ExtractionResult, FieldResult } from './types';

const table: FieldResult = {
  fieldId: 'a',
  fieldName: 'Line items',
  output: 'table',
  ok: true,
  resolvedRange: 'Invoice!A9:C11',
  headers: ['SKU', 'Description', 'Qty'],
  rows: [
    ['A-1001', 'Widget, standard', 12],
    ['A-1002', 'Widget, reinforced', 4],
  ],
};

const total: FieldResult = {
  fieldId: 'b',
  fieldName: 'Total',
  output: 'value',
  ok: true,
  resolvedRange: 'Invoice!E17:E17',
  rows: [],
  value: 231.66,
};

const skus: FieldResult = {
  fieldId: 'c',
  fieldName: 'SKUs',
  output: 'list',
  ok: true,
  resolvedRange: 'Invoice!A10:A11',
  rows: [['A-1001'], ['A-1002']],
};

const broken: FieldResult = {
  fieldId: 'd',
  fieldName: 'Missing',
  output: 'table',
  ok: false,
  error: 'Anchor "Freight" not found',
  rows: [],
};

const result: ExtractionResult = {
  fileName: 'invoice.xlsx',
  templateName: 'Invoice extract',
  fields: [table, total, skus, broken],
};

describe('flattenFieldValue', () => {
  it('JSON-encodes a table as objects when it has headers', () => {
    expect(JSON.parse(flattenFieldValue(table) as string)).toEqual([
      { SKU: 'A-1001', Description: 'Widget, standard', Qty: 12 },
      { SKU: 'A-1002', Description: 'Widget, reinforced', Qty: 4 },
    ]);
  });

  it('JSON-encodes a headerless table as arrays', () => {
    const value = flattenFieldValue({ ...table, headers: undefined });
    expect(JSON.parse(value as string)).toEqual([
      ['A-1001', 'Widget, standard', 12],
      ['A-1002', 'Widget, reinforced', 4],
    ]);
  });

  it('JSON-encodes a list as a flat array', () => {
    expect(JSON.parse(flattenFieldValue(skus) as string)).toEqual(['A-1001', 'A-1002']);
  });

  it('leaves a single value as a bare scalar', () => {
    expect(flattenFieldValue(total)).toBe(231.66);
  });

  it('emits null for a failed field', () => {
    expect(flattenFieldValue(broken)).toBeNull();
  });

  it('keeps a column type stable regardless of how much data a file holds', () => {
    // A one-row table stays JSON — a column that is scalar in one file and an
    // array in the next would break anything consuming the export.
    const oneRow = flattenFieldValue({ ...table, rows: [['A-1001', 'Widget', 1]] });
    const empty = flattenFieldValue({ ...table, rows: [] });

    expect(typeof oneRow).toBe('string');
    expect(empty).toBe('[]');
  });
});

describe('flattenResults', () => {
  it('produces a file column plus one column per field', () => {
    const flat = flattenResults([result]);

    expect(flat.headers).toEqual([FILE_COLUMN, 'Line items', 'Total', 'SKUs', 'Missing']);
    expect(flat.rows).toHaveLength(1);
    expect(flat.rows[0][0]).toBe('invoice.xlsx');
    expect(flat.rows[0][2]).toBe(231.66);
  });

  it('gives each file its own row under one shared header', () => {
    const second: ExtractionResult = {
      ...result,
      fileName: 'invoice-b.xlsx',
      fields: [table, { ...total, value: 99.5 }, skus, broken],
    };
    const flat = flattenResults([result, second]);

    expect(flat.rows).toHaveLength(2);
    expect(flat.rows.map((r) => r[0])).toEqual(['invoice.xlsx', 'invoice-b.xlsx']);
    expect(flat.rows.map((r) => r[2])).toEqual([231.66, 99.5]);
    // Every row must be the same width as the header, or the CSV misaligns.
    expect(new Set(flat.rows.map((r) => r.length))).toEqual(new Set([flat.headers.length]));
  });

  it('keeps rows rectangular when a file fails a field the others resolved', () => {
    const partial: ExtractionResult = {
      ...result,
      fileName: 'partial.xlsx',
      fields: [table, { ...total, ok: false, error: 'Anchor not found', value: undefined }, skus, broken],
    };
    const flat = flattenResults([result, partial]);

    expect(flat.rows[1]).toHaveLength(flat.headers.length);
    expect(flat.rows[1][2]).toBeNull();
  });

  it('handles an empty batch and a fieldless template', () => {
    expect(flattenResults([])).toEqual({ headers: [FILE_COLUMN], rows: [] });
    expect(flattenResults([{ ...result, fields: [] }])).toEqual({
      headers: [FILE_COLUMN],
      rows: [['invoice.xlsx']],
    });
  });
});

describe('flat CSV', () => {
  const csv = resultsToFlatCsv([result]);
  const [header, row] = csv.split('\r\n');

  it('writes one header row and one row per file', () => {
    expect(csv.split('\r\n')).toHaveLength(2);
    expect(header).toBe('File,Line items,Total,SKUs,Missing');
    expect(resultsToFlatCsv([result, result]).split('\r\n')).toHaveLength(3);
  });

  it('quotes the embedded JSON so it survives a CSV round trip', () => {
    // JSON is full of commas and double quotes; both must be escaped.
    expect(row).toContain('"[{""SKU"":""A-1001""');
    expect(row.endsWith(',')).toBe(true); // failed field lands as an empty cell
  });

  it('keeps the scalar column unquoted', () => {
    expect(row).toContain(',231.66,');
  });

  it('every data row has the same column count as the header', () => {
    // Naive splitting would break on the embedded JSON, so count quoted fields.
    const countCols = (line: string) =>
      (line.match(/(^|,)(?:"(?:[^"]|"")*"|[^,]*)/g) ?? []).length;
    const lines = resultsToFlatCsv([result, result]).split('\r\n');

    expect(new Set(lines.map(countCols)).size).toBe(1);
  });
});
