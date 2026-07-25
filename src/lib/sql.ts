import type { CellValue, ExtractionResult, FieldResult } from './types';
import { flattenResults, isJsonCell } from './flatten';

/**
 * T-SQL script generation for SQL Server and SQL Server Express LocalDB.
 *
 * The dialect is identical for both — only the connection string differs — so a
 * single script targets either.
 */

export type SqlMode = 'flat' | 'normalized';

export interface SqlOptions {
  schema: string;
  table: string;
  mode: SqlMode;
  /** Emit DROP TABLE IF EXISTS so the script can be re-run. */
  dropExisting: boolean;
}

export const DEFAULT_SQL_OPTIONS: Omit<SqlOptions, 'mode'> = {
  schema: 'dbo',
  table: 'Extraction',
  dropExisting: true,
};

/** SQL Server caps a VALUES clause at 1000 row constructors per INSERT. */
const MAX_ROWS_PER_INSERT = 1000;
/** sysname limit. */
const MAX_IDENT_LEN = 128;
const MAX_NVARCHAR = 4000;

/* -------------------------------- identifiers ------------------------------- */

/**
 * Bracket-quotes an identifier. Spreadsheet headers are arbitrary text, so this
 * has to survive spaces, reserved words, leading digits and embedded brackets —
 * a `]` is escaped by doubling it, which is what closes the injection hole.
 */
export function quoteIdent(name: string): string {
  const trimmed = name.trim().slice(0, MAX_IDENT_LEN) || 'Column';
  return `[${trimmed.replace(/]/g, ']]')}]`;
}

/** Makes a list of column names unique, since two headers can share a name. */
export function uniqueNames(names: string[]): string[] {
  const seen = new Map<string, number>();
  return names.map((raw) => {
    const base = (raw.trim() || 'Column').slice(0, MAX_IDENT_LEN);
    const key = base.toLowerCase();
    const count = seen.get(key) ?? 0;
    seen.set(key, count + 1);
    return count === 0 ? base : `${base}_${count + 1}`;
  });
}

/* --------------------------------- literals -------------------------------- */

export function quoteLiteral(value: CellValue): string {
  if (value === null) return 'NULL';

  if (typeof value === 'boolean') return value ? '1' : '0';

  if (typeof value === 'number') {
    // NaN and Infinity have no SQL representation.
    if (!Number.isFinite(value)) return 'NULL';
    return Number.isInteger(value) ? String(value) : formatFloat(value);
  }

  // N-prefix so non-ASCII survives; doubling quotes is the escape.
  return `N'${value.replace(/'/g, "''")}'`;
}

function formatFloat(value: number): string {
  const text = String(value);
  // Exponent notation is legal in T-SQL only for FLOAT; expand it so the literal
  // is unambiguous whatever the column type turns out to be.
  return text.includes('e') || text.includes('E') ? value.toFixed(10).replace(/0+$/, '0') : text;
}

/* ------------------------------ type inference ------------------------------ */

const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;

/** Decimal places in a number's plain form, or null if it needs an exponent. */
function decimalPlaces(value: number): number | null {
  const text = String(value);
  if (text.includes('e') || text.includes('E')) return null;
  const dot = text.indexOf('.');
  return dot === -1 ? 0 : text.length - dot - 1;
}

/**
 * Fractional columns get DECIMAL, not FLOAT.
 *
 * Spreadsheet decimals are overwhelmingly money, and FLOAT is binary — a total
 * of 128.8 comes back out of SQL Server as 128.80000000000001. DECIMAL stores
 * it exactly. Scale and precision carry headroom so a later file with a longer
 * value still loads.
 */
function decimalType(values: number[]): string {
  const places = values.map(decimalPlaces);
  if (places.some((p) => p === null)) return 'FLOAT';

  const scale = Math.min(Math.max(...(places as number[])) + 2, 10);
  const intDigits = Math.max(...values.map((v) => Math.trunc(Math.abs(v)).toString().length));
  const precision = Math.min(intDigits + scale + 6, 38);

  // No room for the integer part once scale is taken out; FLOAT is the only fit.
  return precision - scale < intDigits ? 'FLOAT' : `DECIMAL(${precision}, ${scale})`;
}

export function inferType(values: CellValue[], forceJson = false): string {
  if (forceJson) return 'NVARCHAR(MAX)';

  const present = values.filter((v) => v !== null && v !== '');
  if (present.length === 0) return 'NVARCHAR(255)';

  if (present.every((v) => typeof v === 'boolean')) return 'BIT';

  if (present.every((v) => typeof v === 'number')) {
    const numbers = present as number[];
    if (numbers.every((v) => Number.isInteger(v))) {
      const fitsInt = numbers.every((v) => Math.abs(v) <= 2147483647);
      return fitsInt ? 'INT' : 'BIGINT';
    }
    return decimalType(numbers);
  }

  // Dates come out of the parser pre-formatted; only trust unambiguous ISO.
  if (present.every((v) => typeof v === 'string' && ISO_DATE.test(v))) return 'DATE';

  const longest = present.reduce<number>((max, v) => Math.max(max, String(v).length), 0);
  if (longest > MAX_NVARCHAR) return 'NVARCHAR(MAX)';
  // Headroom, so a later file with slightly longer values still loads.
  return `NVARCHAR(${Math.min(MAX_NVARCHAR, Math.max(50, Math.ceil((longest * 1.5) / 50) * 50))})`;
}

/* -------------------------------- statements -------------------------------- */

interface TableSpec {
  name: string;
  columns: { name: string; type: string; json: boolean }[];
  rows: CellValue[][];
}

function qualified(schema: string, table: string): string {
  return `${quoteIdent(schema)}.${quoteIdent(table)}`;
}

function createTable(spec: TableSpec, options: SqlOptions): string {
  const target = qualified(options.schema, spec.name);
  const lines: string[] = [];

  if (options.dropExisting) lines.push(`DROP TABLE IF EXISTS ${target};`, 'GO', '');

  const columns = spec.columns.map((col) => {
    const check = col.json ? ` CHECK (ISJSON(${quoteIdent(col.name)}) = 1)` : '';
    return `    ${quoteIdent(col.name)} ${col.type} NULL${check}`;
  });

  lines.push(`CREATE TABLE ${target} (`, columns.join(',\n'), ');', 'GO', '');
  return lines.join('\n');
}

function insertRows(spec: TableSpec, options: SqlOptions): string {
  if (spec.rows.length === 0) return `-- ${spec.name}: no rows\n`;

  const target = qualified(options.schema, spec.name);
  const columnList = spec.columns.map((c) => quoteIdent(c.name)).join(', ');
  const batches: string[] = [];

  // Chunked because a single INSERT ... VALUES accepts at most 1000 rows.
  for (let i = 0; i < spec.rows.length; i += MAX_ROWS_PER_INSERT) {
    const chunk = spec.rows.slice(i, i + MAX_ROWS_PER_INSERT);
    const values = chunk.map((row) => `    (${row.map(quoteLiteral).join(', ')})`).join(',\n');
    batches.push(`INSERT INTO ${target} (${columnList})\nVALUES\n${values};\nGO\n`);
  }

  return batches.join('\n');
}

function renderTable(spec: TableSpec, options: SqlOptions): string {
  return `${createTable(spec, options)}\n${insertRows(spec, options)}`;
}

/* ---------------------------------- shapes ---------------------------------- */

/** One table: a row per file, multi-row fields held as JSON. */
function flatSpec(results: ExtractionResult[], options: SqlOptions): TableSpec[] {
  const flat = flattenResults(results);
  const names = uniqueNames(flat.headers);

  const columns = names.map((name, i) => {
    // Column 0 is the file name; the rest line up with the template's fields.
    const field = i === 0 ? null : results[0]?.fields[i - 1];
    const json = field ? isJsonCell(field) || field.output !== 'value' : false;
    return {
      name,
      type: inferType(
        flat.rows.map((row) => row[i]),
        json,
      ),
      json,
    };
  });

  return [{ name: options.table, columns, rows: flat.rows }];
}

const SOURCE_COLUMN = 'SourceFile';

/**
 * Relational shape: each table field becomes its own table keyed by source file,
 * each list field a table of positioned values, and all single values collapse
 * into one row-per-file table.
 */
function normalizedSpecs(results: ExtractionResult[], options: SqlOptions): TableSpec[] {
  const template = results[0]?.fields ?? [];
  const specs: TableSpec[] = [];

  const valueFields = template.filter((f) => f.output === 'value');
  if (valueFields.length > 0) {
    const names = uniqueNames([SOURCE_COLUMN, ...valueFields.map((f) => f.fieldName)]);
    const rows = results.map((result) => [
      result.fileName,
      ...valueFields.map((f) => {
        const match = result.fields.find((r) => r.fieldId === f.fieldId);
        return match && match.ok ? (match.value ?? null) : null;
      }),
    ]);
    specs.push({
      name: options.table,
      columns: names.map((name, i) => ({
        name,
        type: i === 0 ? 'NVARCHAR(400)' : inferType(rows.map((r) => r[i])),
        json: false,
      })),
      rows,
    });
  }

  for (const field of template) {
    if (field.output === 'value') continue;

    const perFile = results.map((result) => ({
      fileName: result.fileName,
      field: result.fields.find((r) => r.fieldId === field.fieldId),
    }));

    if (field.output === 'list') {
      const rows: CellValue[][] = [];
      for (const { fileName, field: f } of perFile) {
        if (!f?.ok) continue;
        f.rows.forEach((row, i) => rows.push([fileName, i + 1, row[0] ?? null]));
      }
      specs.push({
        name: `${options.table}_${field.fieldName}`,
        columns: [
          { name: SOURCE_COLUMN, type: 'NVARCHAR(400)', json: false },
          { name: 'Position', type: 'INT', json: false },
          { name: 'Value', type: inferType(rows.map((r) => r[2])), json: false },
        ],
        rows,
      });
      continue;
    }

    // table: widest header wins, so a file with a narrower region still loads.
    const widest = perFile.reduce<FieldResult | undefined>(
      (best, { field: f }) =>
        f?.ok && (!best || (f.headers?.length ?? 0) > (best.headers?.length ?? 0)) ? f : best,
      undefined,
    );
    const width = Math.max(
      widest?.headers?.length ?? 0,
      ...perFile.map(({ field: f }) => (f?.ok ? Math.max(...f.rows.map((r) => r.length), 0) : 0)),
      1,
    );
    const headers =
      widest?.headers ?? Array.from({ length: width }, (_, i) => `Column ${i + 1}`);
    const names = uniqueNames([
      SOURCE_COLUMN,
      ...Array.from({ length: width }, (_, i) => headers[i] ?? `Column ${i + 1}`),
    ]);

    const rows: CellValue[][] = [];
    for (const { fileName, field: f } of perFile) {
      if (!f?.ok) continue;
      for (const row of f.rows) {
        rows.push([fileName, ...Array.from({ length: width }, (_, i) => row[i] ?? null)]);
      }
    }

    specs.push({
      name: `${options.table}_${field.fieldName}`,
      columns: names.map((name, i) => ({
        name,
        type: i === 0 ? 'NVARCHAR(400)' : inferType(rows.map((r) => r[i])),
        json: false,
      })),
      rows,
    });
  }

  return specs;
}

/* ---------------------------------- script ---------------------------------- */

function header(results: ExtractionResult[], options: SqlOptions, tables: TableSpec[]): string {
  const rowTotal = tables.reduce((n, t) => n + t.rows.length, 0);
  return [
    '-- Generated by Excel Template Extractor',
    `-- Template : ${results[0]?.templateName ?? '(none)'}`,
    `-- Files    : ${results.length}`,
    `-- Tables   : ${tables.map((t) => t.name).join(', ')}`,
    `-- Rows     : ${rowTotal}`,
    `-- Shape    : ${options.mode === 'flat' ? 'one row per file, multi-row fields as JSON' : 'one table per field'}`,
    '--',
    '-- Runs unchanged against SQL Server or LocalDB; only the server differs.',
    '--   sqlcmd -S localhost           -d YourDatabase -i extraction.sql',
    '--   sqlcmd -S "(localdb)\\MSSQLLocalDB" -d YourDatabase -i extraction.sql',
    '--',
    '-- Create the database first if it does not exist, e.g.',
    '--   sqlcmd -S "(localdb)\\MSSQLLocalDB" -Q "CREATE DATABASE YourDatabase"',
    '',
    'SET NOCOUNT ON;',
    'SET XACT_ABORT ON;',
    'GO',
    '',
    `IF SCHEMA_ID(${quoteLiteral(options.schema)}) IS NULL`,
    `    EXEC(${quoteLiteral(`CREATE SCHEMA ${quoteIdent(options.schema)}`)});`,
    'GO',
    '',
  ].join('\n');
}

export function buildSqlScript(results: ExtractionResult[], options: SqlOptions): string {
  const tables = options.mode === 'flat' ? flatSpec(results, options) : normalizedSpecs(results, options);
  const body = tables.map((spec) => renderTable(spec, options)).join('\n');
  return `${header(results, options, tables)}${body}`;
}
