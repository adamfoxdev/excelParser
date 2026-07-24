import { describe, expect, it } from 'vitest';
import { deserializeTemplate, serializeTemplate } from './storage';
import { resultToCsv, resultToJson, safeFileName } from './download';
import { ANY_SHEET, type ExtractionResult, type Template } from './types';

const template: Template = {
  id: 'orig-id',
  name: 'Invoice extract',
  description: 'Supplier invoices',
  flatten: true,
  createdAt: 1700000000000,
  updatedAt: 1700000000000,
  fields: [
    {
      id: 'a',
      name: 'Line items',
      selector: {
        kind: 'anchor',
        sheet: ANY_SHEET,
        anchorText: 'SKU',
        matchMode: 'exact',
        caseSensitive: false,
        occurrence: 1,
        offsetRows: 0,
        offsetCols: 0,
        height: 'auto',
        width: 'auto',
      },
      headerRow: true,
      output: 'table',
    },
    {
      id: 'b',
      name: 'Total',
      selector: { kind: 'fixed', sheet: 'Invoice', range: 'D17:D17' },
      headerRow: false,
      output: 'value',
    },
  ],
};

describe('template import/export', () => {
  it('round-trips every selector setting', () => {
    const back = deserializeTemplate(serializeTemplate(template));

    expect(back.name).toBe(template.name);
    expect(back.fields.map((f) => f.selector)).toEqual(template.fields.map((f) => f.selector));
    expect(back.fields.map((f) => f.output)).toEqual(['table', 'value']);
    expect(back.flatten).toBe(true);
  });

  it('defaults flatten to false for templates exported before the option existed', () => {
    expect(deserializeTemplate('{"name":"x","fields":[]}').flatten).toBe(false);
  });

  it('assigns a new id so an import cannot overwrite an existing template', () => {
    expect(deserializeTemplate(serializeTemplate(template)).id).not.toBe(template.id);
  });

  it('accepts a bare template object without the version wrapper', () => {
    expect(deserializeTemplate(JSON.stringify(template)).fields).toHaveLength(2);
  });

  it('rejects malformed input with a readable message', () => {
    expect(() => deserializeTemplate('not json')).toThrow(/not valid JSON/);
    expect(() => deserializeTemplate('{"name":"x"}')).toThrow(/no fields/i);
    expect(() =>
      deserializeTemplate('{"name":"x","fields":[{"selector":{"kind":"wat"}}]}'),
    ).toThrow(/Field 1: unknown selector kind/);
  });

  it('defaults missing anchor settings rather than failing', () => {
    const parsed = deserializeTemplate(
      '{"name":"x","fields":[{"name":"f","selector":{"kind":"anchor","anchorText":"Total"}}]}',
    );
    expect(parsed.fields[0].selector).toMatchObject({
      sheet: ANY_SHEET,
      matchMode: 'exact',
      occurrence: 1,
      height: 'auto',
      width: 'auto',
    });
  });
});

const result: ExtractionResult = {
  fileName: 'invoice.xlsx',
  templateName: 'Invoice extract',
  fields: [
    {
      fieldId: 'a',
      fieldName: 'Line items',
      output: 'table',
      ok: true,
      resolvedRange: 'Invoice!A9:C10',
      headers: ['SKU', 'Description', 'Qty'],
      rows: [['A-1001', 'Widget, "standard"', 12]],
    },
    { fieldId: 'b', fieldName: 'Total', output: 'value', ok: true, rows: [], value: 231.66 },
    { fieldId: 'c', fieldName: 'Missing', output: 'table', ok: false, error: 'Anchor not found', rows: [] },
  ],
};

describe('exports', () => {
  it('shapes JSON by output type and keeps failures visible', () => {
    const json = JSON.parse(resultToJson(result));

    expect(json.data['Line items']).toEqual([
      { SKU: 'A-1001', Description: 'Widget, "standard"', Qty: 12 },
    ]);
    expect(json.data.Total).toBe(231.66);
    expect(json.data.Missing).toEqual({ error: 'Anchor not found' });
  });

  it('escapes quotes and commas in CSV', () => {
    const csv = resultToCsv(result);

    expect(csv).toContain('"Widget, ""standard"""');
    expect(csv).toContain('# Line items (Invoice!A9:C10)');
    expect(csv).toContain('# Missing — FAILED');
  });

  it('sanitises file names', () => {
    expect(safeFileName('Q4 report: final/draft')).toBe('Q4_report_final_draft');
    expect(safeFileName('***')).toBe('extraction');
  });
});
