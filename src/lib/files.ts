/** Collecting input files, including dropped folders. */

export const SUPPORTED_EXTENSIONS = ['.xlsx', '.xlsm', '.xls', '.csv', '.tsv'] as const;

/** Guards against a stray drop of a huge tree locking the tab up. */
export const MAX_BATCH_FILES = 2000;
const MAX_DEPTH = 12;

export function isSupported(name: string): boolean {
  const lower = name.toLowerCase();
  // "~$" files are Office lock files: real extensions, never real workbooks.
  if (lower.startsWith('~$') || lower.startsWith('.')) return false;
  return SUPPORTED_EXTENSIONS.some((ext) => lower.endsWith(ext));
}

export interface Collected {
  files: File[];
  /** Set when the walk stopped early at MAX_BATCH_FILES. */
  truncated: boolean;
  /** Files present in the drop that were not spreadsheets. */
  skipped: number;
}

async function readAllEntries(reader: FileSystemDirectoryReader): Promise<FileSystemEntry[]> {
  // readEntries yields at most ~100 per call and signals the end with an empty
  // batch, so it has to be drained in a loop. Awaiting each batch keeps this
  // flat — recursing per batch would grow the stack once a directory is large.
  const all: FileSystemEntry[] = [];

  for (;;) {
    const batch = await new Promise<FileSystemEntry[]>((resolve, reject) =>
      reader.readEntries(resolve, reject),
    );
    if (batch.length === 0) return all;
    all.push(...batch);
  }
}

function entryToFile(entry: FileSystemFileEntry): Promise<File> {
  return new Promise((resolve, reject) => entry.file(resolve, reject));
}

async function walk(entry: FileSystemEntry, out: Collected, depth: number): Promise<void> {
  if (out.files.length >= MAX_BATCH_FILES) {
    out.truncated = true;
    return;
  }

  if (entry.isFile) {
    try {
      const file = await entryToFile(entry as FileSystemFileEntry);
      if (isSupported(file.name)) out.files.push(file);
      else out.skipped++;
    } catch {
      out.skipped++;
    }
    return;
  }

  if (entry.isDirectory && depth < MAX_DEPTH) {
    const entries = await readAllEntries((entry as FileSystemDirectoryEntry).createReader());
    for (const child of entries) await walk(child, out, depth + 1);
  }
}

/** Walks dropped entries, descending into folders. */
export async function collectFromEntries(
  entries: readonly (FileSystemEntry | null)[],
): Promise<Collected> {
  const out: Collected = { files: [], truncated: false, skipped: 0 };
  for (const entry of entries) {
    if (entry) await walk(entry, out, 0);
  }
  return out;
}

/**
 * Files from a drop, descending into any dropped folders.
 *
 * The entry handles must be read synchronously — a DataTransfer is neutered as
 * soon as the drop handler yields — so they are captured before the first await.
 */
export async function filesFromDrop(transfer: DataTransfer): Promise<Collected> {
  const entries = [...transfer.items]
    .filter((item) => item.kind === 'file')
    .map((item) => item.webkitGetAsEntry?.() ?? null);
  const plainFiles = [...transfer.files];

  if (entries.some(Boolean)) return collectFromEntries(entries);

  // Browsers without the entries API: flat file list only.
  return collectFromList(plainFiles);
}

/** Files chosen through an <input>, including a directory picker. */
export function collectFromList(list: ArrayLike<File>): Collected {
  const out: Collected = { files: [], truncated: false, skipped: 0 };

  for (const file of Array.from(list)) {
    if (out.files.length >= MAX_BATCH_FILES) {
      out.truncated = true;
      break;
    }
    if (isSupported(file.name)) out.files.push(file);
    else out.skipped++;
  }

  return out;
}

/** Human-readable note about what a collection step dropped, if anything. */
export function describeCollection(c: Collected): string | null {
  const parts: string[] = [];
  if (c.skipped > 0) parts.push(`${c.skipped} non-spreadsheet file${c.skipped === 1 ? '' : 's'} skipped`);
  if (c.truncated) parts.push(`stopped at the ${MAX_BATCH_FILES}-file limit`);
  return parts.length ? parts.join(' · ') : null;
}
