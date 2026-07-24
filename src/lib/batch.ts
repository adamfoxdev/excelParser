import { applyTemplate } from './extract';
import { loadWorkbook } from './workbook';
import type { ExtractionResult, Template } from './types';

export interface BatchFile {
  id: string;
  name: string;
  /** Disk-backed handle. Holding this costs nothing; the parsed workbook does. */
  file: File;
}

export interface BatchFailure {
  fileId: string;
  fileName: string;
  error: string;
}

export interface BatchOutcome {
  results: ExtractionResult[];
  failures: BatchFailure[];
  /** True when the run stopped early because it was cancelled. */
  cancelled: boolean;
}

export interface BatchProgress {
  done: number;
  total: number;
  fileName: string;
}

export function toBatchFile(file: File): BatchFile {
  return { id: crypto.randomUUID(), name: file.name, file };
}

/** Lets the browser paint progress between files. */
const yieldToUi = () => new Promise((resolve) => setTimeout(resolve, 0));

/**
 * Extracts a template across many files without ever holding more than one
 * parsed workbook in memory.
 *
 * Each file is parsed, extracted, and dropped before the next one starts — the
 * extraction results are a tiny fraction of a workbook's dense cell matrix, so
 * peak memory stays flat no matter how long the batch is.
 */
export async function runBatch(
  files: BatchFile[],
  template: Template,
  options: { onProgress?: (p: BatchProgress) => void; signal?: AbortSignal } = {},
): Promise<BatchOutcome> {
  const { onProgress, signal } = options;
  const results: ExtractionResult[] = [];
  const failures: BatchFailure[] = [];

  for (const [i, entry] of files.entries()) {
    if (signal?.aborted) return { results, failures, cancelled: true };

    onProgress?.({ done: i, total: files.length, fileName: entry.name });
    await yieldToUi();

    try {
      const workbook = await loadWorkbook(entry.file);
      results.push(applyTemplate(workbook, template));
      // `workbook` goes out of scope here; only the extraction is retained.
    } catch (err) {
      failures.push({
        fileId: entry.id,
        fileName: entry.name,
        error: err instanceof Error ? err.message : String(err),
      });
    }
  }

  onProgress?.({ done: files.length, total: files.length, fileName: '' });
  return { results, failures, cancelled: false };
}

/**
 * Identifies the extraction-relevant part of a template. Used to tell whether a
 * completed batch still reflects the template on screen — renaming a template
 * or toggling flatten must not mark thousands of extracted rows stale.
 */
export function templateSignature(template: Template): string {
  return JSON.stringify(template.fields);
}
