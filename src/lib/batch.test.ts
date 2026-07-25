import { describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import { runBatch, templateSignature, toBatchFile, type BatchProgress } from './batch';
import { ANY_SHEET, type Template } from './types';

/** A real .xlsx buffer, so these tests exercise the actual parser. */
function invoiceFile(name: string, padTop: number, total: number): File {
  const rows: unknown[][] = [
    ...Array.from({ length: padTop }, () => []),
    ['ACME SUPPLY CO.'],
    [],
    ['SKU', 'Description', 'Qty'],
    ['A-1001', 'Widget', 12],
    ['B-2010', 'Bracket', 30],
    [],
    ['', '', 'Invoice Total', total],
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Invoice');
  const buffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;
  return new File([buffer], name);
}

const template: Template = {
  id: 't',
  name: 'Invoice',
  description: '',
  flatten: true,
  sql: { schema: 'dbo', table: 'Extraction', dropExisting: true },
  createdAt: 0,
  updatedAt: 0,
  fields: [
    {
      id: 'items',
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
      id: 'total',
      name: 'Total',
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
    },
  ],
};

describe('runBatch', () => {
  it('extracts every file, at differing offsets', async () => {
    const files = [
      invoiceFile('a.xlsx', 0, 100),
      invoiceFile('b.xlsx', 3, 200),
      invoiceFile('c.xlsx', 9, 300),
    ].map(toBatchFile);

    const outcome = await runBatch(files, template);

    expect(outcome.failures).toEqual([]);
    expect(outcome.cancelled).toBe(false);
    expect(outcome.results.map((r) => r.fileName)).toEqual(['a.xlsx', 'b.xlsx', 'c.xlsx']);
    expect(outcome.results.map((r) => r.fields[1].value)).toEqual([100, 200, 300]);
    // The anchored table resolves to a different range in each file.
    expect(new Set(outcome.results.map((r) => r.fields[0].resolvedRange)).size).toBe(3);
  });

  it('isolates an unreadable file without losing the rest', async () => {
    const files = [
      invoiceFile('good-1.xlsx', 0, 100),
      // A zip header followed by junk: fails to inflate.
      new File([new Uint8Array([0x50, 0x4b, 0x03, 0x04, ...Array(200).fill(7)])], 'bad.xlsx'),
      invoiceFile('good-2.xlsx', 2, 300),
    ].map(toBatchFile);

    const outcome = await runBatch(files, template);

    expect(outcome.results.map((r) => r.fileName)).toEqual(['good-1.xlsx', 'good-2.xlsx']);
    expect(outcome.failures).toHaveLength(1);
    expect(outcome.failures[0].fileName).toBe('bad.xlsx');
    expect(outcome.failures[0].error).toBeTruthy();
  });

  it('reports progress once per file, in order', async () => {
    const files = [invoiceFile('a.xlsx', 0, 1), invoiceFile('b.xlsx', 0, 2)].map(toBatchFile);
    const seen: BatchProgress[] = [];

    await runBatch(files, template, { onProgress: (p) => seen.push(p) });

    expect(seen.map((p) => p.done)).toEqual([0, 1, 2]);
    expect(seen.map((p) => p.fileName)).toEqual(['a.xlsx', 'b.xlsx', '']);
    expect(seen.every((p) => p.total === 2)).toBe(true);
  });

  it('stops early when cancelled, keeping what it already extracted', async () => {
    const files = [
      invoiceFile('a.xlsx', 0, 1),
      invoiceFile('b.xlsx', 0, 2),
      invoiceFile('c.xlsx', 0, 3),
    ].map(toBatchFile);

    const controller = new AbortController();
    const outcome = await runBatch(files, template, {
      signal: controller.signal,
      onProgress: (p) => {
        if (p.done === 1) controller.abort();
      },
    });

    // Cancelling takes effect at the next file boundary: the file already in
    // flight finishes rather than having its parse thrown away, so aborting as
    // b starts yields a and b, and c is never read.
    expect(outcome.cancelled).toBe(true);
    expect(outcome.results.map((r) => r.fileName)).toEqual(['a.xlsx', 'b.xlsx']);
  });

  it('handles an empty batch', async () => {
    expect(await runBatch([], template)).toEqual({ results: [], failures: [], cancelled: false });
  });
});

describe('templateSignature', () => {
  it('changes when a selector changes', () => {
    const edited: Template = {
      ...template,
      fields: [
        { ...template.fields[0], selector: { kind: 'fixed', sheet: 'Invoice', range: 'A1:B2' } },
        template.fields[1],
      ],
    };

    expect(templateSignature(edited)).not.toBe(templateSignature(template));
  });

  it('ignores changes that do not affect extraction', () => {
    // Renaming a template or flipping the view toggle must not invalidate a
    // batch that took minutes to run.
    const renamed: Template = {
      ...template,
      name: 'Something else',
      description: 'notes',
      flatten: !template.flatten,
      updatedAt: Date.now(),
    };

    expect(templateSignature(renamed)).toBe(templateSignature(template));
  });
});
