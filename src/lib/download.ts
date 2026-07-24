import type { CellValue, ExtractionResult, FieldResult } from './types';

export function downloadFile(name: string, contents: string, mime: string): void {
  const url = URL.createObjectURL(new Blob([contents], { type: mime }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  URL.revokeObjectURL(url);
}

function csvCell(value: CellValue): string {
  if (value === null) return '';
  const text = String(value);
  return /[",\n\r]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function fieldToCsv(field: FieldResult): string {
  const lines: string[] = [];
  if (field.headers) lines.push(field.headers.map((h) => csvCell(h)).join(','));
  for (const row of field.rows) lines.push(row.map(csvCell).join(','));
  return lines.join('\r\n');
}

/** One CSV per extraction, with each field in its own labelled block. */
export function resultToCsv(result: ExtractionResult): string {
  const blocks = result.fields.map((field) => {
    const heading = `# ${field.fieldName}${field.ok ? ` (${field.resolvedRange})` : ' — FAILED'}`;
    const body = field.ok ? fieldToCsv(field) : `# ${field.error ?? 'unknown error'}`;
    return `${heading}\r\n${body}`;
  });
  return blocks.join('\r\n\r\n');
}

/** Field name → rows of objects (when headers exist) or arrays. */
export function resultToJson(result: ExtractionResult): string {
  const out: Record<string, unknown> = {};

  for (const field of result.fields) {
    if (!field.ok) {
      out[field.fieldName] = { error: field.error };
      continue;
    }
    if (field.output === 'value') {
      out[field.fieldName] = field.value ?? null;
      continue;
    }
    if (field.output === 'list') {
      out[field.fieldName] = field.rows.map((r) => r[0] ?? null);
      continue;
    }
    out[field.fieldName] = field.headers
      ? field.rows.map((row) =>
          Object.fromEntries(field.headers!.map((h, i) => [h, row[i] ?? null])),
        )
      : field.rows;
  }

  return JSON.stringify(
    { file: result.fileName, template: result.templateName, data: out },
    null,
    2,
  );
}

export function safeFileName(name: string): string {
  return name.replace(/[^A-Za-z0-9._-]+/g, '_').replace(/^_+|_+$/g, '') || 'extraction';
}
