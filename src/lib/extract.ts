import type {
  AnchorSelector,
  CellValue,
  ExtractionResult,
  Field,
  FieldResult,
  FixedSelector,
  SheetData,
  Selector,
  Template,
  WorkbookData,
} from './types';
import { ANY_SHEET } from './types';
import { cellAt, isBlank } from './workbook';
import {
  decodeRange,
  encodeRange,
  qualifyRange,
  rangeHeight,
  rangeWidth,
  type CellRef,
  type RangeRef,
} from './range';

export interface ResolvedRegion {
  sheet: SheetData;
  range: RangeRef;
}

/** Thrown when a selector can't be pointed at a real region in this workbook. */
export class ResolveError extends Error {}

function textOf(value: CellValue): string {
  return value === null ? '' : String(value);
}

function matches(value: CellValue, sel: AnchorSelector): boolean {
  if (value === null) return false;
  const haystack = sel.caseSensitive ? textOf(value) : textOf(value).toLowerCase();
  const needle = sel.caseSensitive ? sel.anchorText : sel.anchorText.toLowerCase();

  switch (sel.matchMode) {
    case 'exact':
      return haystack.trim() === needle.trim();
    case 'contains':
      return haystack.includes(needle);
    case 'regex':
      try {
        return new RegExp(sel.anchorText, sel.caseSensitive ? '' : 'i').test(textOf(value));
      } catch {
        throw new ResolveError(`Invalid regular expression: "${sel.anchorText}"`);
      }
  }
}

function findAnchor(sheets: SheetData[], sel: AnchorSelector): { sheet: SheetData; ref: CellRef } {
  let seen = 0;
  for (const sheet of sheets) {
    for (let r = 0; r < sheet.rowCount; r++) {
      for (let c = 0; c < sheet.colCount; c++) {
        if (matches(cellAt(sheet, r, c), sel)) {
          seen++;
          if (seen === sel.occurrence) return { sheet, ref: { row: r, col: c } };
        }
      }
    }
  }
  const suffix = seen > 0 ? ` (found ${seen}, needed occurrence ${sel.occurrence})` : '';
  throw new ResolveError(`Anchor "${sel.anchorText}" not found${suffix}`);
}

/** Grows right from the origin while the origin row keeps producing values. */
function autoWidth(sheet: SheetData, origin: CellRef): number {
  let width = 0;
  while (
    origin.col + width < sheet.colCount &&
    !isBlank(cellAt(sheet, origin.row, origin.col + width))
  ) {
    width++;
  }
  return Math.max(width, 1);
}

/** Grows down while any cell across the region's width is non-blank. */
function autoHeight(sheet: SheetData, origin: CellRef, width: number): number {
  let height = 0;
  while (origin.row + height < sheet.rowCount) {
    let rowHasData = false;
    for (let c = origin.col; c < origin.col + width; c++) {
      if (!isBlank(cellAt(sheet, origin.row + height, c))) {
        rowHasData = true;
        break;
      }
    }
    if (!rowHasData) break;
    height++;
  }
  return Math.max(height, 1);
}

function resolveFixed(wb: WorkbookData, sel: FixedSelector): ResolvedRegion {
  const sheet = wb.sheets.find((s) => s.name === sel.sheet);
  if (!sheet) {
    throw new ResolveError(
      `Sheet "${sel.sheet}" is not in this workbook (has: ${wb.sheets.map((s) => s.name).join(', ')})`,
    );
  }
  return { sheet, range: decodeRange(sel.range) };
}

function resolveAnchor(wb: WorkbookData, sel: AnchorSelector): ResolvedRegion {
  const searchIn =
    sel.sheet === ANY_SHEET ? wb.sheets : wb.sheets.filter((s) => s.name === sel.sheet);

  if (searchIn.length === 0) {
    throw new ResolveError(`Sheet "${sel.sheet}" is not in this workbook`);
  }

  const { sheet, ref } = findAnchor(searchIn, sel);
  const origin: CellRef = { row: ref.row + sel.offsetRows, col: ref.col + sel.offsetCols };

  if (origin.row < 0 || origin.col < 0) {
    throw new ResolveError(
      `Offset puts the region off the top/left of the sheet (anchor at ${encodeRange({ start: ref, end: ref }).split(':')[0]})`,
    );
  }

  const width = sel.width === 'auto' ? autoWidth(sheet, origin) : Math.max(sel.width, 1);
  const height = sel.height === 'auto' ? autoHeight(sheet, origin, width) : Math.max(sel.height, 1);

  return {
    sheet,
    range: {
      start: origin,
      end: { row: origin.row + height - 1, col: origin.col + width - 1 },
    },
  };
}

export function resolveSelector(wb: WorkbookData, selector: Selector): ResolvedRegion {
  return selector.kind === 'fixed' ? resolveFixed(wb, selector) : resolveAnchor(wb, selector);
}

export function readRegion(sheet: SheetData, range: RangeRef): CellValue[][] {
  const rows: CellValue[][] = [];
  for (let r = range.start.row; r <= range.end.row; r++) {
    const row: CellValue[] = [];
    for (let c = range.start.col; c <= range.end.col; c++) {
      row.push(cellAt(sheet, r, c));
    }
    rows.push(row);
  }
  return rows;
}

export function extractField(wb: WorkbookData, field: Field): FieldResult {
  const base: FieldResult = {
    fieldId: field.id,
    fieldName: field.name,
    output: field.output,
    ok: false,
    rows: [],
  };

  let region: ResolvedRegion;
  try {
    region = resolveSelector(wb, field.selector);
  } catch (err) {
    return { ...base, error: err instanceof Error ? err.message : String(err) };
  }

  const { sheet, range } = region;
  const resolvedRange = qualifyRange(sheet.name, encodeRange(range));
  let rows = readRegion(sheet, range);
  let headers: string[] | undefined;

  if (field.headerRow && rows.length > 0) {
    headers = rows[0].map((v, i) => {
      const label = textOf(v).trim();
      return label || `Column ${i + 1}`;
    });
    rows = rows.slice(1);
  }

  if (field.output === 'value') {
    const flat = rows.flat().find((v) => !isBlank(v));
    return { ...base, ok: true, resolvedRange, headers, rows, value: flat ?? null };
  }

  if (field.output === 'list') {
    const flat = rows.flat().filter((v) => !isBlank(v));
    return { ...base, ok: true, resolvedRange, headers, rows: flat.map((v) => [v]) };
  }

  return { ...base, ok: true, resolvedRange, headers, rows };
}

export function applyTemplate(wb: WorkbookData, template: Template): ExtractionResult {
  return {
    fileName: wb.fileName,
    templateName: template.name,
    fields: template.fields.map((f) => extractField(wb, f)),
  };
}

export function applyTemplateToAll(
  workbooks: WorkbookData[],
  template: Template,
): ExtractionResult[] {
  return workbooks.map((wb) => applyTemplate(wb, template));
}

export function failedFieldCount(result: ExtractionResult): number {
  return result.fields.filter((f) => !f.ok).length;
}

/** Human-readable summary of where a selector points, for the fields list. */
export function describeSelector(selector: Selector): string {
  if (selector.kind === 'fixed') return qualifyRange(selector.sheet, selector.range);

  const sheet = selector.sheet === ANY_SHEET ? 'any sheet' : selector.sheet;
  const size = `${selector.height === 'auto' ? 'auto' : selector.height}×${
    selector.width === 'auto' ? 'auto' : selector.width
  }`;
  const offset =
    selector.offsetRows === 0 && selector.offsetCols === 0
      ? 'at anchor'
      : `${selector.offsetRows >= 0 ? '+' : ''}${selector.offsetRows}r ${
          selector.offsetCols >= 0 ? '+' : ''
        }${selector.offsetCols}c`;
  return `${sheet} · "${selector.anchorText}" ${offset} · ${size}`;
}

export function regionSize(range: RangeRef): string {
  return `${rangeHeight(range)} × ${rangeWidth(range)}`;
}
