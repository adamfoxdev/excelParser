export type CellValue = string | number | boolean | null;

export interface SheetData {
  name: string;
  /** Dense matrix of the sheet's used range, origin at A1. */
  cells: CellValue[][];
  rowCount: number;
  colCount: number;
}

export interface WorkbookData {
  /** Stable per-load id: file names are not unique across a batch. */
  id: string;
  fileName: string;
  sheets: SheetData[];
}

/** A literal sheet + A1 range. Predictable, but breaks if rows shift. */
export interface FixedSelector {
  kind: 'fixed';
  sheet: string;
  /** A1 range, e.g. "B2:D50". Single cells are stored as "B2:B2". */
  range: string;
}

export type MatchMode = 'exact' | 'contains' | 'regex';

/** Finds a label in the sheet, then reads a region relative to it. */
export interface AnchorSelector {
  kind: 'anchor';
  /** Sheet name, or ANY_SHEET to search every sheet in order. */
  sheet: string;
  anchorText: string;
  matchMode: MatchMode;
  caseSensitive: boolean;
  /** 1-based: which matching cell to use when the label appears more than once. */
  occurrence: number;
  /** Region origin relative to the anchor cell. */
  offsetRows: number;
  offsetCols: number;
  /** 'auto' grows until a blank row/column terminates the region. */
  height: number | 'auto';
  width: number | 'auto';
}

export type Selector = FixedSelector | AnchorSelector;

export const ANY_SHEET = '*';

export type FieldOutput = 'table' | 'list' | 'value';

export interface Field {
  id: string;
  name: string;
  selector: Selector;
  /** Treat the region's first row as column headers. */
  headerRow: boolean;
  output: FieldOutput;
}

export interface Template {
  id: string;
  name: string;
  description: string;
  fields: Field[];
  /**
   * Present results and CSV as a single flat row — one column per field, with
   * multi-row fields JSON-encoded into their cell.
   */
  flatten: boolean;
  /** Target for the generated T-SQL load script. */
  sql: SqlTarget;
  createdAt: number;
  updatedAt: number;
}

export interface SqlTarget {
  schema: string;
  table: string;
  dropExisting: boolean;
}

export interface FieldResult {
  fieldId: string;
  fieldName: string;
  output: FieldOutput;
  ok: boolean;
  error?: string;
  /** Where the selector actually landed, e.g. "Sheet1!B2:D50". */
  resolvedRange?: string;
  headers?: string[];
  rows: CellValue[][];
  /** Populated when output is 'value'. */
  value?: CellValue;
}

export interface ExtractionResult {
  fileName: string;
  templateName: string;
  fields: FieldResult[];
}
