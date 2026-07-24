import { describe, expect, it } from 'vitest';
import {
  applyTemplate,
  applyTemplateToAll,
  extractField,
  failedFieldCount,
  resolveSelector,
} from './extract';
import { encodeRange } from './range';
import { ANY_SHEET, type CellValue, type Field, type SheetData, type WorkbookData } from './types';

function sheet(name: string, cells: CellValue[][]): SheetData {
  const colCount = Math.max(...cells.map((r) => r.length), 0);
  const padded = cells.map((r) => [...r, ...Array(colCount - r.length).fill(null)]);
  return { name, cells: padded, rowCount: padded.length, colCount };
}

/**
 * Invoice-shaped fixture. `padTop` shifts everything down, standing in for a
 * new file that has extra header rows.
 */
function invoice(padTop: number): WorkbookData {
  const pad: CellValue[][] = Array.from({ length: padTop }, () => []);
  return {
    id: `wb-${padTop}`,
    fileName: 'invoice.xlsx',
    sheets: [
      sheet('Invoice', [
        ...pad,
        ['ACME SUPPLY CO.'],
        [],
        ['Invoice Number', 'INV-2024-0871'],
        ['Customer', 'Northwind Traders'],
        [],
        ['Line Items'],
        ['SKU', 'Description', 'Qty', 'Unit Price'],
        ['A-1001', 'Widget, standard', 12, 4.5],
        ['A-1002', 'Widget, reinforced', 4, 11.25],
        ['B-2010', 'Bracket set', 30, 2.1],
        [],
        [null, null, 'Invoice Total', 231.66],
      ]),
      sheet('Regional Summary', [
        ['Region', 'Q1', 'Q2'],
        ['North', 120500, 133200],
        ['South', 98700, 101300],
      ]),
    ],
  };
}

const tableField = (selector: Field['selector']): Field => ({
  id: 'f1',
  name: 'Line items',
  selector,
  headerRow: true,
  output: 'table',
});

describe('fixed selectors', () => {
  it('reads the requested range', () => {
    const result = extractField(
      invoice(0),
      tableField({ kind: 'fixed', sheet: 'Invoice', range: 'A7:D10' }),
    );

    expect(result.ok).toBe(true);
    expect(result.headers).toEqual(['SKU', 'Description', 'Qty', 'Unit Price']);
    expect(result.rows).toHaveLength(3);
    expect(result.rows[0]).toEqual(['A-1001', 'Widget, standard', 12, 4.5]);
  });

  it('reports a missing sheet instead of throwing', () => {
    const result = extractField(
      invoice(0),
      tableField({ kind: 'fixed', sheet: 'Nope', range: 'A1:B2' }),
    );

    expect(result.ok).toBe(false);
    expect(result.error).toContain('Nope');
  });

  it('drifts onto the wrong data when the file shifts', () => {
    const result = extractField(
      invoice(4),
      tableField({ kind: 'fixed', sheet: 'Invoice', range: 'A7:D10' }),
    );

    // Same range, different file: this is exactly the failure anchoring solves.
    expect(result.rows[0]).not.toEqual(['A-1001', 'Widget, standard', 12, 4.5]);
  });
});

describe('anchored selectors', () => {
  const lineItemsAnchor = tableField({
    kind: 'anchor',
    sheet: 'Invoice',
    anchorText: 'SKU',
    matchMode: 'exact',
    caseSensitive: false,
    occurrence: 1,
    offsetRows: 0,
    offsetCols: 0,
    height: 'auto',
    width: 'auto',
  });

  it('finds the same table regardless of row shift', () => {
    for (const padTop of [0, 4, 17]) {
      const result = extractField(invoice(padTop), lineItemsAnchor);

      expect(result.ok).toBe(true);
      expect(result.headers).toEqual(['SKU', 'Description', 'Qty', 'Unit Price']);
      expect(result.rows).toHaveLength(3);
      expect(result.rows.at(-1)).toEqual(['B-2010', 'Bracket set', 30, 2.1]);
    }
  });

  it('stops the auto region at the blank row', () => {
    const region = resolveSelector(invoice(0), lineItemsAnchor.selector);
    expect(encodeRange(region.range)).toBe('A7:D10');
  });

  it('pulls a single value to the right of a label', () => {
    const field: Field = {
      id: 'total',
      name: 'Invoice total',
      selector: {
        kind: 'anchor',
        sheet: ANY_SHEET,
        anchorText: 'Invoice Total',
        matchMode: 'exact',
        caseSensitive: false,
        occurrence: 1,
        offsetRows: 0,
        offsetCols: 1,
        height: 1,
        width: 1,
      },
      headerRow: false,
      output: 'value',
    };

    expect(extractField(invoice(9), field).value).toBe(231.66);
  });

  it('distinguishes occurrences of a repeated label', () => {
    const wb: WorkbookData = {
      id: 'wb-repeat',
      fileName: 'x.xlsx',
      sheets: [sheet('S', [['Total', 1], ['Total', 2], ['Total', 3]])],
    };
    const at = (occurrence: number) =>
      extractField(wb, {
        id: 'v',
        name: 'v',
        selector: {
          kind: 'anchor',
          sheet: 'S',
          anchorText: 'Total',
          matchMode: 'exact',
          caseSensitive: false,
          occurrence,
          offsetRows: 0,
          offsetCols: 1,
          height: 1,
          width: 1,
        },
        headerRow: false,
        output: 'value',
      }).value;

    expect([at(1), at(2), at(3)]).toEqual([1, 2, 3]);
  });

  it('searches every sheet when the sheet is ANY_SHEET', () => {
    const region = resolveSelector(invoice(0), {
      kind: 'anchor',
      sheet: ANY_SHEET,
      anchorText: 'Region',
      matchMode: 'exact',
      caseSensitive: false,
      occurrence: 1,
      offsetRows: 0,
      offsetCols: 0,
      height: 'auto',
      width: 'auto',
    });

    expect(region.sheet.name).toBe('Regional Summary');
  });

  it('honours contains and regex matching', () => {
    const find = (matchMode: 'contains' | 'regex', anchorText: string) =>
      resolveSelector(invoice(0), {
        kind: 'anchor',
        sheet: 'Invoice',
        anchorText,
        matchMode,
        caseSensitive: false,
        occurrence: 1,
        offsetRows: 0,
        offsetCols: 0,
        height: 1,
        width: 1,
      });

    expect(encodeRange(find('contains', 'Invoice Num').range)).toBe('A3:A3');
    expect(encodeRange(find('regex', '^Line Items$').range)).toBe('A6:A6');
  });

  it('reports a missing anchor with the occurrence count', () => {
    const result = extractField(invoice(0), {
      ...lineItemsAnchor,
      selector: { ...lineItemsAnchor.selector, occurrence: 5 } as Field['selector'],
    });

    expect(result.ok).toBe(false);
    expect(result.error).toMatch(/occurrence 5/);
  });

  it('reports an invalid regex rather than crashing', () => {
    const result = extractField(invoice(0), {
      ...lineItemsAnchor,
      selector: {
        ...lineItemsAnchor.selector,
        matchMode: 'regex',
        anchorText: '([',
      } as Field['selector'],
    });

    expect(result.ok).toBe(false);
    expect(result.error).toContain('Invalid regular expression');
  });
});

describe('output shapes', () => {
  const region: Field['selector'] = { kind: 'fixed', sheet: 'Invoice', range: 'A7:D10' };

  it('flattens a list and drops blanks', () => {
    const result = extractField(invoice(0), {
      id: 'l',
      name: 'SKUs',
      selector: { kind: 'fixed', sheet: 'Invoice', range: 'A8:A12' },
      headerRow: false,
      output: 'list',
    });

    expect(result.rows.flat()).toEqual(['A-1001', 'A-1002', 'B-2010']);
  });

  it('names unlabelled header columns', () => {
    const result = extractField(invoice(0), {
      id: 't',
      name: 't',
      selector: { kind: 'fixed', sheet: 'Invoice', range: 'A11:D12' },
      headerRow: true,
      output: 'table',
    });

    expect(result.headers).toEqual(['Column 1', 'Column 2', 'Column 3', 'Column 4']);
  });

  it('applies one template across a batch of files', () => {
    const template = {
      id: 't',
      name: 'Invoice',
      description: '',
      flatten: true,
      createdAt: 0,
      updatedAt: 0,
      fields: [
        {
          id: 'items',
          name: 'Line items',
          selector: {
            kind: 'anchor',
            sheet: 'Invoice',
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
        } satisfies Field,
      ],
    };

    // Same template, three files with the table at different offsets.
    const results = applyTemplateToAll([invoice(0), invoice(4), invoice(11)], template);

    expect(results).toHaveLength(3);
    expect(results.every((r) => r.fields[0].ok)).toBe(true);
    expect(results.map((r) => r.fields[0].resolvedRange)).toEqual([
      'Invoice!A7:D10',
      'Invoice!A11:D14',
      'Invoice!A18:D21',
    ]);
    // Different ranges, identical extracted data.
    expect(new Set(results.map((r) => JSON.stringify(r.fields[0].rows))).size).toBe(1);
  });

  it('counts failed fields per file', () => {
    const template = {
      id: 't',
      name: 'T',
      description: '',
      flatten: false,
      createdAt: 0,
      updatedAt: 0,
      fields: [
        tableField({ kind: 'fixed', sheet: 'Invoice', range: 'A7:D10' }),
        { ...tableField({ kind: 'fixed', sheet: 'Ghost', range: 'A1:A1' }), id: 'x', name: 'X' },
      ],
    };

    expect(applyTemplateToAll([invoice(0)], template).map(failedFieldCount)).toEqual([1]);
  });

  it('applies a whole template, isolating per-field failures', () => {
    const result = applyTemplate(invoice(0), {
      id: 't',
      name: 'Invoice',
      description: '',
      flatten: false,
      createdAt: 0,
      updatedAt: 0,
      fields: [
        tableField(region),
        { ...tableField({ kind: 'fixed', sheet: 'Ghost', range: 'A1:A1' }), id: 'bad', name: 'Bad' },
      ],
    });

    expect(result.fields.map((f) => f.ok)).toEqual([true, false]);
  });
});
