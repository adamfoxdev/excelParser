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

const lineItems = [
  ['SKU', 'Description', 'Qty', 'Unit Price', 'Line Total'],
  ['A-1001', 'Widget, standard', 12, 4.5, 54],
  ['A-1002', 'Widget, reinforced', 4, 11.25, 45],
  ['B-2010', 'Bracket set', 30, 2.1, 63],
  ['C-3300', 'Fastener pack (100)', 6, 8.75, 52.5],
];

function buildInvoice({ padTop }) {
  const rows = [];
  const blank = () => rows.push([]);

  for (let i = 0; i < padTop; i++) blank();

  rows.push(['ACME SUPPLY CO.']);
  rows.push(['123 Industrial Way, Springfield']);
  blank();
  rows.push(['Invoice Number', 'INV-2024-0871']);
  rows.push(['Invoice Date', '2024-11-04']);
  rows.push(['Customer', 'Northwind Traders']);
  blank();
  rows.push(['Line Items']);
  for (const item of lineItems) rows.push(item);
  blank();
  rows.push(['', '', '', 'Subtotal', 214.5]);
  rows.push(['', '', '', 'Tax (8%)', 17.16]);
  rows.push(['', '', '', 'Invoice Total', 231.66]);

  return rows;
}

const summary = [
  ['Region', 'Q1', 'Q2', 'Q3', 'Q4'],
  ['North', 120500, 133200, 128900, 141000],
  ['South', 98700, 101300, 96400, 110800],
  ['East', 143000, 139500, 150200, 162400],
  ['West', 87200, 91900, 88600, 94300],
];

for (const [name, padTop] of [
  ['invoice-a.xlsx', 0],
  ['invoice-b-shifted.xlsx', 4],
]) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(buildInvoice({ padTop })), 'Invoice');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summary), 'Regional Summary');
  writeFileSync(join(here, name), XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
  console.log(`wrote samples/${name} (${padTop} padding rows)`);
}
