import { createStore, del, get, set, values } from 'idb-keyval';
import type { Field, Selector, SqlTarget, Template } from './types';
import { ANY_SHEET } from './types';

export const DEFAULT_SQL_TARGET: SqlTarget = {
  schema: 'dbo',
  table: 'Extraction',
  dropExisting: true,
};

/**
 * Templates live behind this interface so a server-backed store can replace the
 * IndexedDB one without touching the UI.
 */
export interface TemplateStore {
  list(): Promise<Template[]>;
  get(id: string): Promise<Template | undefined>;
  save(template: Template): Promise<void>;
  remove(id: string): Promise<void>;
}

const store = createStore('excel-parser', 'templates');

/** Fills in fields added after a template was written, so older saves stay usable. */
function normalize(template: Template): Template {
  return {
    ...template,
    flatten: template.flatten ?? false,
    sql: { ...DEFAULT_SQL_TARGET, ...(template.sql ?? {}) },
  };
}

export const indexedDbStore: TemplateStore = {
  async list() {
    const all = (await values(store)) as Template[];
    return all.map(normalize).sort((a, b) => b.updatedAt - a.updatedAt);
  },
  async get(id) {
    const found = await get<Template>(id, store);
    return found && normalize(found);
  },
  save: (template) => set(template.id, template, store),
  remove: (id) => del(id, store),
};

export function newId(): string {
  return crypto.randomUUID();
}

export function emptyTemplate(): Template {
  const now = Date.now();
  return {
    id: newId(),
    name: 'Untitled template',
    description: '',
    fields: [],
    flatten: false,
    sql: { ...DEFAULT_SQL_TARGET },
    createdAt: now,
    updatedAt: now,
  };
}

/* ------------------------------- import/export ------------------------------ */

const EXPORT_VERSION = 1;

export function serializeTemplate(template: Template): string {
  return JSON.stringify({ version: EXPORT_VERSION, template }, null, 2);
}

function parseSelector(raw: unknown): Selector {
  if (typeof raw !== 'object' || raw === null) throw new Error('field is missing a selector');
  const s = raw as Record<string, unknown>;

  if (s.kind === 'fixed') {
    if (typeof s.sheet !== 'string' || typeof s.range !== 'string') {
      throw new Error('fixed selector needs a sheet and a range');
    }
    return { kind: 'fixed', sheet: s.sheet, range: s.range };
  }

  if (s.kind === 'anchor') {
    if (typeof s.anchorText !== 'string') throw new Error('anchor selector needs anchorText');
    const dim = (v: unknown): number | 'auto' =>
      v === 'auto' || typeof v !== 'number' ? 'auto' : v;
    return {
      kind: 'anchor',
      sheet: typeof s.sheet === 'string' ? s.sheet : ANY_SHEET,
      anchorText: s.anchorText,
      matchMode: s.matchMode === 'contains' || s.matchMode === 'regex' ? s.matchMode : 'exact',
      caseSensitive: s.caseSensitive === true,
      occurrence: typeof s.occurrence === 'number' && s.occurrence > 0 ? s.occurrence : 1,
      offsetRows: typeof s.offsetRows === 'number' ? s.offsetRows : 0,
      offsetCols: typeof s.offsetCols === 'number' ? s.offsetCols : 0,
      height: dim(s.height),
      width: dim(s.width),
    };
  }

  throw new Error(`unknown selector kind "${String(s.kind)}"`);
}

function parseSqlTarget(raw: unknown): SqlTarget {
  if (typeof raw !== 'object' || raw === null) return { ...DEFAULT_SQL_TARGET };
  const s = raw as Record<string, unknown>;
  return {
    schema: typeof s.schema === 'string' && s.schema.trim() ? s.schema : DEFAULT_SQL_TARGET.schema,
    table: typeof s.table === 'string' && s.table.trim() ? s.table : DEFAULT_SQL_TARGET.table,
    dropExisting: s.dropExisting !== false,
  };
}

/**
 * Validates an imported template. Imported JSON is untrusted input — a bad file
 * should surface a readable error, not corrupt the template list.
 */
export function deserializeTemplate(json: string): Template {
  let parsed: unknown;
  try {
    parsed = JSON.parse(json);
  } catch {
    throw new Error('That file is not valid JSON.');
  }

  const root = parsed as Record<string, unknown>;
  const raw = (root?.template ?? root) as Record<string, unknown>;
  if (typeof raw !== 'object' || raw === null) throw new Error('No template found in that file.');
  if (!Array.isArray(raw.fields)) throw new Error('Template has no fields array.');

  const fields: Field[] = raw.fields.map((f, i) => {
    const rf = f as Record<string, unknown>;
    try {
      return {
        id: typeof rf.id === 'string' ? rf.id : newId(),
        name: typeof rf.name === 'string' && rf.name.trim() ? rf.name : `Field ${i + 1}`,
        selector: parseSelector(rf.selector),
        headerRow: rf.headerRow === true,
        output: rf.output === 'list' || rf.output === 'value' ? rf.output : 'table',
      };
    } catch (err) {
      throw new Error(`Field ${i + 1}: ${err instanceof Error ? err.message : String(err)}`);
    }
  });

  const now = Date.now();
  return {
    // A fresh id keeps an import from silently overwriting an existing template.
    id: newId(),
    name: typeof raw.name === 'string' && raw.name.trim() ? raw.name : 'Imported template',
    description: typeof raw.description === 'string' ? raw.description : '',
    fields,
    flatten: raw.flatten === true,
    sql: parseSqlTarget(raw.sql),
    createdAt: typeof raw.createdAt === 'number' ? raw.createdAt : now,
    updatedAt: now,
  };
}
