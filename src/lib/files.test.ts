import { describe, expect, it } from 'vitest';
import {
  collectFromEntries,
  collectFromList,
  describeCollection,
  isSupported,
  MAX_BATCH_FILES,
} from './files';

const file = (name: string) => new File([''], name);

/* --- Minimal stand-ins for the FileSystem entry API, which jsdom lacks. --- */

function fileEntry(name: string): FileSystemEntry {
  return {
    isFile: true,
    isDirectory: false,
    name,
    file: (cb: (f: File) => void) => cb(file(name)),
  } as unknown as FileSystemEntry;
}

/** readEntries must yield its children then an empty array to signal the end. */
function dirEntry(name: string, children: FileSystemEntry[]): FileSystemEntry {
  return {
    isFile: false,
    isDirectory: true,
    name,
    createReader: () => {
      let drained = false;
      return {
        readEntries: (cb: (e: FileSystemEntry[]) => void) => {
          cb(drained ? [] : children);
          drained = true;
        },
      };
    },
  } as unknown as FileSystemEntry;
}

describe('isSupported', () => {
  it('accepts the spreadsheet extensions, case-insensitively', () => {
    for (const name of ['a.xlsx', 'b.XLSM', 'c.xls', 'd.csv', 'e.TSV']) {
      expect(isSupported(name)).toBe(true);
    }
  });

  it('rejects everything else', () => {
    for (const name of ['notes.txt', 'report.pdf', 'archive.zip', 'xlsx', 'data.xlsx.bak']) {
      expect(isSupported(name)).toBe(false);
    }
  });

  it('rejects Office lock files and dotfiles', () => {
    // "~$budget.xlsx" is a lock file Excel leaves beside an open workbook — a
    // real extension that never parses, and a folder drop is full of them.
    expect(isSupported('~$budget.xlsx')).toBe(false);
    expect(isSupported('.hidden.csv')).toBe(false);
  });
});

describe('collectFromList', () => {
  it('keeps spreadsheets and counts what it skipped', () => {
    const collected = collectFromList([
      file('a.xlsx'),
      file('readme.md'),
      file('b.csv'),
      file('~$a.xlsx'),
    ]);

    expect(collected.files.map((f) => f.name)).toEqual(['a.xlsx', 'b.csv']);
    expect(collected.skipped).toBe(2);
    expect(collected.truncated).toBe(false);
  });

  it('stops at the batch cap and flags it', () => {
    const many = Array.from({ length: MAX_BATCH_FILES + 25 }, (_, i) => file(`f${i}.xlsx`));
    const collected = collectFromList(many);

    expect(collected.files).toHaveLength(MAX_BATCH_FILES);
    expect(collected.truncated).toBe(true);
  });

  it('handles an empty selection', () => {
    expect(collectFromList([])).toEqual({ files: [], truncated: false, skipped: 0 });
  });
});

describe('collectFromEntries (folder drop)', () => {
  it('descends into nested folders and keeps only spreadsheets', async () => {
    const tree = [
      dirEntry('2024', [
        fileEntry('jan.xlsx'),
        fileEntry('notes.txt'),
        dirEntry('q2', [fileEntry('apr.csv'), fileEntry('may.xlsx')]),
      ]),
      fileEntry('loose.xls'),
    ];

    const collected = await collectFromEntries(tree);

    expect(collected.files.map((f) => f.name).sort()).toEqual([
      'apr.csv',
      'jan.xlsx',
      'loose.xls',
      'may.xlsx',
    ]);
    expect(collected.skipped).toBe(1);
  });

  it('ignores Office lock files a real folder is littered with', async () => {
    const collected = await collectFromEntries([
      dirEntry('books', [fileEntry('budget.xlsx'), fileEntry('~$budget.xlsx')]),
    ]);

    expect(collected.files.map((f) => f.name)).toEqual(['budget.xlsx']);
  });

  it('stops descending past the depth limit rather than recursing forever', async () => {
    // 20 levels deep, one workbook at the bottom; the limit is 12.
    let deepest = dirEntry('bottom', [fileEntry('buried.xlsx')]);
    for (let i = 0; i < 20; i++) deepest = dirEntry(`level-${i}`, [deepest]);

    const collected = await collectFromEntries([deepest]);

    expect(collected.files).toHaveLength(0);
  });

  it('skips null entries without throwing', async () => {
    const collected = await collectFromEntries([null, fileEntry('a.xlsx'), null]);
    expect(collected.files.map((f) => f.name)).toEqual(['a.xlsx']);
  });
});

describe('describeCollection', () => {
  it('says nothing when everything was accepted', () => {
    expect(describeCollection({ files: [file('a.xlsx')], truncated: false, skipped: 0 })).toBeNull();
  });

  it('reports skips and truncation together', () => {
    const note = describeCollection({ files: [], truncated: true, skipped: 3 });
    expect(note).toContain('3 non-spreadsheet files skipped');
    expect(note).toContain(`${MAX_BATCH_FILES}-file limit`);
  });
});
