/**
 * Generates two sample workbooks with the same logical content but different
 * row offsets, so you can prove an anchored template survives a layout shift.
 *
 *   node samples/make-samples.mjs
 */
import * as XLSX from 'xlsx';
import { writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));

const defaultItems = [
  ['SKU', 'Description', 'Qty', 'Unit Price', 'Line Total'],
  ['A-1001', 'Widget, standard', 12, 4.5, 54],
  ['A-1002', 'Widget, reinforced', 4, 11.25, 45],
  ['B-2010', 'Bracket set', 30, 2.1, 63],
  ['C-3300', 'Fastener pack (100)', 6, 8.75, 52.5],
];

function buildInvoice({ padTop, lineItems = defaultItems, number = 'INV-2024-0871', total = 231.66 }) {
  const rows = [];
  const blank = () => rows.push([]);

  for (let i = 0; i < padTop; i++) blank();

  rows.push(['ACME SUPPLY CO.']);
  rows.push(['123 Industrial Way, Springfield']);
  blank();
  rows.push(['Invoice Number', number]);
  rows.push(['Invoice Date', '2024-11-04']);
  rows.push(['Customer', 'Northwind Traders']);
  blank();
  rows.push(['Line Items']);
  for (const item of lineItems) rows.push(item);
  blank();
  rows.push(['', '', '', 'Subtotal', Number((total / 1.08).toFixed(2))]);
  rows.push(['', '', '', 'Tax (8%)', Number((total - total / 1.08).toFixed(2))]);
  rows.push(['', '', '', 'Invoice Total', total]);

  return rows;
}

const summary = [
  ['Region', 'Q1', 'Q2', 'Q3', 'Q4'],
  ['North', 120500, 133200, 128900, 141000],
  ['South', 98700, 101300, 96400, 110800],
  ['East', 143000, 139500, 150200, 162400],
  ['West', 87200, 91900, 88600, 94300],
];

/**
 * Every invoice holds the same logical content at a different vertical offset,
 * with different row counts — a batch that only an anchored template survives.
 */
const invoices = [
  { name: 'invoice-a.xlsx', padTop: 0 },
  { name: 'invoice-b-shifted.xlsx', padTop: 4 },
  {
    name: 'invoice-c.xlsx',
    padTop: 2,
    number: 'INV-2024-0902',
    total: 118.8,
    lineItems: [
      ['SKU', 'Description', 'Qty', 'Unit Price', 'Line Total'],
      ['A-1001', 'Widget, standard', 20, 4.5, 90],
      ['D-4100', 'Gasket, nitrile', 10, 2.0, 20],
    ],
  },
  {
    name: 'invoice-d.xlsx',
    padTop: 9,
    number: 'INV-2024-0915',
    total: 402.15,
    lineItems: [
      ['SKU', 'Description', 'Qty', 'Unit Price', 'Line Total'],
      ['B-2010', 'Bracket set', 60, 2.1, 126],
      ['C-3300', 'Fastener pack (100)', 12, 8.75, 105],
      ['E-5000', 'Rail, 2m', 7, 18.4, 128.8],
      ['F-6200', 'End cap', 25, 1.69, 42.25],
    ],
  },
];

for (const { name, ...spec } of invoices) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(buildInvoice(spec)), 'Invoice');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summary), 'Regional Summary');
  writeFileSync(join(here, name), XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
  console.log(`wrote samples/${name} (${spec.padTop} padding rows)`);
}
