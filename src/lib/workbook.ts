import type * as XLSX from 'xlsx';
import type { CellValue, SheetData, WorkbookData } from './types';
import { decodeCell } from './range';

/**
 * SheetJS is ~500kB and is only needed once a user actually opens a file, so it
 * is pulled in on demand and cached rather than shipped in the initial bundle.
 */
let sheetJs: Promise<typeof XLSX> | null = null;
const loadSheetJs = () => (sheetJs ??= import('xlsx'));

/** Hard cap so a pathological file can't hang the grid. */
const MAX_ROWS = 20000;
const MAX_COLS = 512;

function toCellValue(cell: XLSX.CellObject | undefined): CellValue {
  if (!cell || cell.v === undefined || cell.v === null) return null;
  // Dates arrive as Date objects (cellDates); prefer Excel's own formatted text
  // so the grid shows what the user sees in Excel.
  if (cell.t === 'd') return cell.w ?? (cell.v as Date).toISOString().slice(0, 10);
  if (cell.t === 'e') return null; // error cells read as empty
  if (typeof cell.v === 'boolean' || typeof cell.v === 'number') return cell.v;
  return String(cell.v);
}

function sheetToData(xlsx: typeof XLSX, name: string, sheet: XLSX.WorkSheet): SheetData {
  const ref = sheet['!ref'];
  if (!ref) return { name, cells: [], rowCount: 0, colCount: 0 };

  const [startA1, endA1] = ref.split(':');
  const start = decodeCell(startA1);
  const end = decodeCell(endA1 ?? startA1);

  // The used range can start below/right of A1; we always materialise from A1 so
  // grid coordinates and A1 references line up without an offset everywhere.
  const rowCount = Math.min(end.row + 1, MAX_ROWS);
  const colCount = Math.min(end.col + 1, MAX_COLS);

  const cells: CellValue[][] = [];
  for (let r = 0; r < rowCount; r++) {
    const row: CellValue[] = new Array(colCount).fill(null);
    for (let c = start.col; c < colCount; c++) {
      row[c] = toCellValue(sheet[xlsx.utils.encode_cell({ r, c })]);
    }
    cells.push(row);
  }

  return { name, cells, rowCount, colCount };
}

export async function loadWorkbook(file: File): Promise<WorkbookData> {
  const xlsx = await loadSheetJs();
  const buffer = await file.arrayBuffer();
  const wb = xlsx.read(buffer, { type: 'array', cellDates: true });

  const sheets = wb.SheetNames.map((name) => sheetToData(xlsx, name, wb.Sheets[name]));
  if (sheets.length === 0) throw new Error('That workbook contains no sheets.');

  return { id: crypto.randomUUID(), fileName: file.name, sheets };
}

export interface LoadFailure {
  fileName: string;
  error: string;
}

export interface BatchLoad {
  loaded: WorkbookData[];
  failed: LoadFailure[];
}

/**
 * Loads a batch one file at a time, reporting each file's outcome separately —
 * one unreadable workbook must not discard the rest of the batch. `onProgress`
 * fires after each file so the UI can show which one is being read.
 */
export async function loadWorkbooks(
  files: File[],
  onProgress?: (done: number, total: number, fileName: string) => void,
): Promise<BatchLoad> {
  const loaded: WorkbookData[] = [];
  const failed: LoadFailure[] = [];

  for (const [i, file] of files.entries()) {
    onProgress?.(i, files.length, file.name);
    try {
      loaded.push(await loadWorkbook(file));
    } catch (err) {
      failed.push({ fileName: file.name, error: err instanceof Error ? err.message : String(err) });
    }
  }

  onProgress?.(files.length, files.length, '');
  return { loaded, failed };
}

export function getSheet(wb: WorkbookData, name: string): SheetData | undefined {
  return wb.sheets.find((s) => s.name === name);
}

export function cellAt(sheet: SheetData, row: number, col: number): CellValue {
  return sheet.cells[row]?.[col] ?? null;
}

export function isBlank(value: CellValue): boolean {
  return value === null || (typeof value === 'string' && value.trim() === '');
}
