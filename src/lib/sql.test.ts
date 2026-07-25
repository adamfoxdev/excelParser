import { describe, expect, it } from 'vitest';
import {
  buildSqlScript,
  DEFAULT_SQL_OPTIONS,
  inferType,
  quoteIdent,
  quoteLiteral,
  uniqueNames,
  type SqlOptions,
} from './sql';
import type { ExtractionResult, FieldResult } from './types';

const opts = (over: Partial<SqlOptions> = {}): SqlOptions => ({
  ...DEFAULT_SQL_OPTIONS,
  mode: 'flat',
  ...over,
});

describe('quoteIdent', () => {
  it('brackets names so spaces and reserved words are safe', () => {
    expect(quoteIdent('Unit Price')).toBe('[Unit Price]');
    expect(quoteIdent('Order')).toBe('[Order]');
    expect(quoteIdent('2024 Total')).toBe('[2024 Total]');
  });

  it('doubles a closing bracket, which is what closes the injection hole', () => {
    // Without this, a header of "a] ; DROP TABLE x --" would break out.
    expect(quoteIdent('a]b')).toBe('[a]]b]');
    expect(quoteIdent('x]; DROP TABLE Users --')).toBe('[x]]; DROP TABLE Users --]');
  });

  it('falls back for an empty name and truncates at the sysname limit', () => {
    expect(quoteIdent('   ')).toBe('[Column]');
    expect(quoteIdent('a'.repeat(200))).toBe(`[${'a'.repeat(128)}]`);
  });
});

describe('uniqueNames', () => {
  it('disambiguates repeated headers', () => {
    expect(uniqueNames(['Total', 'Total', 'Total'])).toEqual(['Total', 'Total_2', 'Total_3']);
  });

  it('treats case-insensitive collisions as duplicates, as SQL Server does', () => {
    expect(uniqueNames(['Qty', 'qty'])).toEqual(['Qty', 'qty_2']);
  });

  it('names blank headers', () => {
    expect(uniqueNames(['', ' '])).toEqual(['Column', 'Column_2']);
  });
});

describe('quoteLiteral', () => {
  it('doubles single quotes', () => {
    expect(quoteLiteral("O'Brien")).toBe("N'O''Brien'");
    expect(quoteLiteral("'; DROP TABLE Users; --")).toBe("N'''; DROP TABLE Users; --'");
  });

  it('emits NULL, numbers and bits unquoted', () => {
    expect(quoteLiteral(null)).toBe('NULL');
    expect(quoteLiteral(42)).toBe('42');
    expect(quoteLiteral(4.5)).toBe('4.5');
    expect(quoteLiteral(true)).toBe('1');
    expect(quoteLiteral(false)).toBe('0');
  });

  it('has no representation for non-finite numbers, so writes NULL', () => {
    expect(quoteLiteral(Number.NaN)).toBe('NULL');
    expect(quoteLiteral(Number.POSITIVE_INFINITY)).toBe('NULL');
  });

  it('keeps newlines inside the literal', () => {
    expect(quoteLiteral('a\nb')).toBe("N'a\nb'");
  });
});

describe('inferType', () => {
  it('picks the narrowest integer type that fits', () => {
    expect(inferType([1, 2, 3])).toBe('INT');
    expect(inferType([1, 9_000_000_000])).toBe('BIGINT');
  });

  it('uses BIT for booleans', () => {
    expect(inferType([true, false])).toBe('BIT');
  });

  it('uses DECIMAL for fractional numbers so money stays exact', () => {
    // FLOAT would store 128.8 as 128.80000000000001.
    expect(inferType([231.66, 128.8])).toBe('DECIMAL(13, 4)');
    expect(inferType([1, 2.5])).toBe('DECIMAL(10, 3)');
  });

  it('leaves the scale room for a later file with more decimals', () => {
    const type = inferType([1.5]);
    const [, scale] = /DECIMAL\((\d+), (\d+)\)/.exec(type)!.slice(1);
    expect(Number(scale)).toBeGreaterThan(1);
  });

  it('falls back to FLOAT when a value needs exponent notation', () => {
    expect(inferType([1e300, 2.5])).toBe('FLOAT');
  });

  it('only trusts unambiguous ISO dates', () => {
    expect(inferType(['2024-11-04', '2024-12-01'])).toBe('DATE');
    // 03/04/2024 is March or April depending on locale — stay textual.
    expect(inferType(['03/04/2024'])).toMatch(/^NVARCHAR/);
  });

  it('sizes text with headroom and falls back to MAX', () => {
    expect(inferType(['abc'])).toBe('NVARCHAR(50)');
    expect(inferType(['x'.repeat(5000)])).toBe('NVARCHAR(MAX)');
  });

  it('ignores nulls and blanks, and defaults an all-empty column', () => {
    expect(inferType([null, 5, null])).toBe('INT');
    expect(inferType([null, null])).toBe('NVARCHAR(255)');
  });

  it('forces JSON columns to NVARCHAR(MAX) regardless of content', () => {
    expect(inferType([1, 2], true)).toBe('NVARCHAR(MAX)');
  });
});

/* --------------------------------- scripts --------------------------------- */

const lineItems: FieldResult = {
  fieldId: 'a',
  fieldName: 'Line items',
  output: 'table',
  ok: true,
  resolvedRange: 'Invoice!A9:C11',
  headers: ['SKU', 'Description', 'Qty'],
  rows: [
    ['A-1001', "Widget, O'Brien", 12],
    ['A-1002', 'Bracket', 4],
  ],
};

const total: FieldResult = {
  fieldId: 'b',
  fieldName: 'Invoice total',
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
  rows: [['A-1001'], ['A-1002']],
};

const result: ExtractionResult = {
  fileName: 'invoice.xlsx',
  templateName: 'Invoice extract',
  fields: [lineItems, total, skus],
};

const second: ExtractionResult = {
  ...result,
  fileName: 'invoice-b.xlsx',
  fields: [lineItems, { ...total, value: 99.5 }, skus],
};

describe('flat script', () => {
  const sql = buildSqlScript([result, second], opts());

  it('creates one table with a column per field', () => {
    expect(sql).toContain('CREATE TABLE [dbo].[Extraction] (');
    expect(sql).toContain('[File]');
    expect(sql).toContain('[Line items] NVARCHAR(MAX)');
    expect(sql).toContain('[SKUs] NVARCHAR(MAX)');
  });

  it('constrains JSON columns but leaves scalars alone', () => {
    expect(sql).toContain('CHECK (ISJSON([Line items]) = 1)');
    expect(sql).not.toMatch(/\[Invoice total\][^\n]*ISJSON/);
  });

  it('inserts one row per file', () => {
    expect(sql).toContain("N'invoice.xlsx'");
    expect(sql).toContain("N'invoice-b.xlsx'");
    expect(sql.match(/^ {4}\(N'invoice/gm)).toHaveLength(2);
  });

  it('escapes quotes nested inside the JSON payload', () => {
    // The apostrophe is doubled inside the surrounding literal, so SQL Server
    // unescapes it back to valid JSON rather than terminating the string early.
    expect(sql).toContain(`"Description":"Widget, O''Brien"`);
    expect(sql).not.toContain(`"Description":"Widget, O'Brien"`);
  });

  it('honours dropExisting', () => {
    expect(sql).toContain('DROP TABLE IF EXISTS [dbo].[Extraction];');
    expect(buildSqlScript([result], opts({ dropExisting: false }))).not.toContain('DROP TABLE');
  });

  it('creates the schema when missing and documents both targets', () => {
    expect(sql).toContain("IF SCHEMA_ID(N'dbo') IS NULL");
    expect(sql).toContain('sqlcmd -S localhost');
    expect(sql).toContain('(localdb)\\MSSQLLocalDB');
  });

  it('quotes a schema and table the user typed, however hostile', () => {
    const nasty = buildSqlScript([result], opts({ schema: 'my]schema', table: 'a]; DROP --' }));
    expect(nasty).toContain('[my]]schema].[a]]; DROP --]');
  });
});

describe('normalized script', () => {
  const sql = buildSqlScript([result, second], opts({ mode: 'normalized' }));

  it('gives each non-scalar field its own table, keyed by source file', () => {
    expect(sql).toContain('CREATE TABLE [dbo].[Extraction_Line items] (');
    expect(sql).toContain('CREATE TABLE [dbo].[Extraction_SKUs] (');
    expect(sql).toContain('[SourceFile] NVARCHAR(400)');
  });

  it('turns spreadsheet headers into real columns', () => {
    expect(sql).toContain('[SKU] NVARCHAR(50)');
    expect(sql).toContain('[Qty] INT');
  });

  it('escapes quotes in cell values', () => {
    expect(sql).toContain("N'Widget, O''Brien'");
  });

  it('collapses every single-value field into one row-per-file table', () => {
    expect(sql).toContain('CREATE TABLE [dbo].[Extraction] (');
    expect(sql).toMatch(/\[Invoice total\] DECIMAL\(\d+, \d+\)/);
    expect(sql).toContain("(N'invoice.xlsx', 231.66)");
    expect(sql).toContain("(N'invoice-b.xlsx', 99.5)");
  });

  it('numbers list positions so order survives the round trip', () => {
    expect(sql).toContain('[Position] INT');
    expect(sql).toContain("(N'invoice.xlsx', 1, N'A-1001')");
    expect(sql).toContain("(N'invoice.xlsx', 2, N'A-1002')");
  });

  it('repeats table rows per file rather than merging them', () => {
    const inserts = sql.match(/N'invoice(-b)?\.xlsx', N'A-1001'/g) ?? [];
    expect(inserts).toHaveLength(2);
  });
});

describe('T-SQL limits and edge cases', () => {
  it('chunks inserts at the 1000-row-constructor cap', () => {
    // A single INSERT ... VALUES with 1001 rows is a hard syntax error.
    const big: FieldResult = {
      ...lineItems,
      rows: Array.from({ length: 2500 }, (_, i) => [`SKU-${i}`, 'x', i]),
    };
    const sql = buildSqlScript(
      [{ ...result, fields: [big] }],
      opts({ mode: 'normalized' }),
    );

    const statements = sql.match(/INSERT INTO/g) ?? [];
    expect(statements).toHaveLength(3); // 1000 + 1000 + 500

    for (const block of sql.split('INSERT INTO').slice(1)) {
      const rowCount = (block.match(/^ {4}\(/gm) ?? []).length;
      expect(rowCount).toBeLessThanOrEqual(1000);
    }
  });

  it('writes a comment instead of an empty INSERT when a field found nothing', () => {
    const empty: ExtractionResult = {
      ...result,
      fields: [{ ...lineItems, rows: [] }],
    };
    const sql = buildSqlScript([empty], opts({ mode: 'normalized' }));

    expect(sql).toContain('no rows');
    expect(sql).not.toMatch(/VALUES\s*;/);
  });

  it('emits NULL for a field that failed to resolve', () => {
    const failed: ExtractionResult = {
      ...result,
      fields: [lineItems, { ...total, ok: false, error: 'not found', value: undefined }, skus],
    };
    const sql = buildSqlScript([failed], opts());

    expect(sql).toContain('NULL');
  });

  it('produces a valid script for an empty batch', () => {
    const sql = buildSqlScript([], opts());
    expect(sql).toContain('CREATE TABLE [dbo].[Extraction]');
    expect(sql).toContain('-- Files    : 0');
  });

  it('pads a file whose table is narrower than another file\'s', () => {
    const narrow: ExtractionResult = {
      ...result,
      fileName: 'narrow.xlsx',
      fields: [
        { ...lineItems, headers: ['SKU'], rows: [['Z-9']] },
        total,
        skus,
      ],
    };
    const sql = buildSqlScript([result, narrow], opts({ mode: 'normalized' }));

    // The wider header set defines the table; the narrow file pads with NULL.
    expect(sql).toContain('[Description]');
    expect(sql).toContain("(N'narrow.xlsx', N'Z-9', NULL, NULL)");
  });
});
