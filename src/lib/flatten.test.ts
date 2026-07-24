import { describe, expect, it } from 'vitest';
import { flattenFieldValue, flattenResult } from './flatten';
import { resultToFlatCsv } from './download';
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

describe('flattenResult', () => {
  it('produces one column per field and a single row', () => {
    const flat = flattenResult(result);

    expect(flat.headers).toEqual(['Line items', 'Total', 'SKUs', 'Missing']);
    expect(flat.rows).toHaveLength(1);
    expect(flat.rows[0]).toHaveLength(4);
    expect(flat.rows[0][1]).toBe(231.66);
  });

  it('handles a template with no fields', () => {
    expect(flattenResult({ ...result, fields: [] })).toEqual({ headers: [], rows: [[]] });
  });
});

describe('flat CSV', () => {
  const csv = resultToFlatCsv(result);
  const [header, row] = csv.split('\r\n');

  it('writes exactly one header row and one data row', () => {
    expect(csv.split('\r\n')).toHaveLength(2);
    expect(header).toBe('Line items,Total,SKUs,Missing');
  });

  it('quotes the embedded JSON so it survives a CSV round trip', () => {
    // JSON is full of commas and double quotes; both must be escaped.
    expect(row).toContain('"[{""SKU"":""A-1001""');
    expect(row.endsWith(',')).toBe(true); // failed field lands as an empty cell
  });

  it('keeps the scalar column unquoted', () => {
    expect(row).toContain(',231.66,');
  });
});
