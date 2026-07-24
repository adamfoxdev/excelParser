import { useCallback, useDeferredValue, useEffect, useMemo, useRef, useState } from 'react';
import { SheetGrid } from './components/SheetGrid';
import { FieldEditor } from './components/FieldEditor';
import { ResultsPanel } from './components/ResultsPanel';
import { FileBar } from './components/FileBar';
import { loadWorkbooks, type LoadFailure } from './lib/workbook';
import { applyTemplate, applyTemplateToAll, describeSelector, resolveSelector } from './lib/extract';
import { encodeRange, type RangeRef } from './lib/range';
import {
  deserializeTemplate,
  emptyTemplate,
  indexedDbStore,
  newId,
  serializeTemplate,
} from './lib/storage';
import { downloadFile, safeFileName } from './lib/download';
import type { Field, Template, WorkbookData } from './lib/types';
import './App.css';

interface Selection {
  sheet: string;
  range: RangeRef;
}

export default function App() {
  const [workbooks, setWorkbooks] = useState<WorkbookData[]>([]);
  const [activeFileId, setActiveFileId] = useState<string | null>(null);
  const [loadFailures, setLoadFailures] = useState<LoadFailure[]>([]);
  const [progress, setProgress] = useState<string | null>(null);
  const [activeSheet, setActiveSheet] = useState('');
  const [selection, setSelection] = useState<Selection | null>(null);
  const [template, setTemplate] = useState<Template>(emptyTemplate);
  const [saved, setSaved] = useState<Template[]>([]);
  const [activeFieldId, setActiveFieldId] = useState<string | null>(null);
  const [tab, setTab] = useState<'template' | 'results'>('template');
  const [status, setStatus] = useState<{ text: string; tone: 'ok' | 'error' } | null>(null);
  const [loading, setLoading] = useState(false);

  const fileInput = useRef<HTMLInputElement>(null);
  const importInput = useRef<HTMLInputElement>(null);

  const refreshSaved = useCallback(async () => {
    setSaved(await indexedDbStore.list());
  }, []);

  useEffect(() => {
    void refreshSaved();
  }, [refreshSaved]);

  useEffect(() => {
    if (!status) return;
    const t = setTimeout(() => setStatus(null), 4000);
    return () => clearTimeout(t);
  }, [status]);

  const workbook = useMemo(
    () => workbooks.find((w) => w.id === activeFileId) ?? null,
    [workbooks, activeFileId],
  );

  const sheet = useMemo(
    () => workbook?.sheets.find((s) => s.name === activeSheet) ?? null,
    [workbook, activeSheet],
  );

  const openFiles = async (files: File[]) => {
    if (files.length === 0) return;
    setLoading(true);
    setProgress(files.length > 1 ? `Reading 1 of ${files.length}…` : null);

    const { loaded, failed } = await loadWorkbooks(files, (done, total, fileName) => {
      if (total > 1 && done < total) setProgress(`Reading ${done + 1} of ${total} — ${fileName}`);
    });

    setProgress(null);
    setLoading(false);
    setLoadFailures((prev) => [...prev.filter((f) => !failed.some((n) => n.fileName === f.fileName)), ...failed]);

    if (loaded.length === 0) {
      setStatus({
        text:
          failed.length === 1
            ? `Could not read ${failed[0].fileName}: ${failed[0].error}`
            : `None of those ${failed.length} files could be read.`,
        tone: 'error',
      });
      return;
    }

    // Keep the file already in the grid; only jump to a new one on a fresh start.
    setWorkbooks((prev) => {
      const next = [...prev, ...loaded];
      if (prev.length === 0) {
        setActiveFileId(loaded[0].id);
        setActiveSheet(loaded[0].sheets[0].name);
        setSelection(null);
      }
      return next;
    });

    const loadedText =
      loaded.length === 1
        ? `Loaded ${loaded[0].fileName} — ${loaded[0].sheets.length} sheet(s)`
        : `Loaded ${loaded.length} files`;

    setStatus({
      text: failed.length ? `${loadedText} · ${failed.length} unreadable` : loadedText,
      tone: failed.length ? 'error' : 'ok',
    });
  };

  const removeFile = (id: string) => {
    setWorkbooks((prev) => {
      const next = prev.filter((w) => w.id !== id);
      if (id === activeFileId) {
        const fallback = next[0] ?? null;
        setActiveFileId(fallback?.id ?? null);
        setActiveSheet(fallback?.sheets[0].name ?? '');
        setSelection(null);
      }
      return next;
    });
  };

  const selectFile = (id: string) => {
    const wb = workbooks.find((w) => w.id === id);
    if (!wb) return;
    setActiveFileId(id);
    // Stay on the same sheet name across files where it exists — templates are
    // usually built against a consistent sheet.
    if (!wb.sheets.some((s) => s.name === activeSheet)) setActiveSheet(wb.sheets[0].name);
    setSelection(null);
  };

  const clearFiles = () => {
    setWorkbooks([]);
    setActiveFileId(null);
    setActiveSheet('');
    setSelection(null);
    setLoadFailures([]);
  };

  const touch = (fields: Field[]) => setTemplate((t) => ({ ...t, fields, updatedAt: Date.now() }));

  const addFieldFromSelection = () => {
    if (!selection) return;
    const range = encodeRange(selection.range);
    // Guess a name from the cell above the selection, then the top-left cell.
    const above = sheet?.cells[selection.range.start.row - 1]?.[selection.range.start.col];
    const topLeft = sheet?.cells[selection.range.start.row]?.[selection.range.start.col];
    const guess = [above, topLeft].find((v) => typeof v === 'string' && v.trim());

    const field: Field = {
      id: newId(),
      name: (guess as string | undefined)?.trim() || `Field ${template.fields.length + 1}`,
      selector: { kind: 'fixed', sheet: selection.sheet, range },
      headerRow: selection.range.end.row > selection.range.start.row,
      output: 'table',
    };
    touch([...template.fields, field]);
    setActiveFieldId(field.id);
  };

  const updateField = (updated: Field) =>
    touch(template.fields.map((f) => (f.id === updated.id ? updated : f)));

  const deleteField = (id: string) => {
    touch(template.fields.filter((f) => f.id !== id));
    setActiveFieldId(null);
  };

  const saveTemplate = async () => {
    const toSave = { ...template, updatedAt: Date.now() };
    await indexedDbStore.save(toSave);
    setTemplate(toSave);
    await refreshSaved();
    setStatus({ text: `Saved "${toSave.name}"`, tone: 'ok' });
  };

  const loadTemplate = async (id: string) => {
    const t = await indexedDbStore.get(id);
    if (!t) return;
    setTemplate(t);
    setActiveFieldId(null);
    setStatus({ text: `Loaded template "${t.name}"`, tone: 'ok' });
  };

  const deleteTemplate = async (t: Template) => {
    if (!window.confirm(`Delete the saved template "${t.name}"? This cannot be undone.`)) return;
    await indexedDbStore.remove(t.id);
    await refreshSaved();
    setStatus({ text: `Deleted "${t.name}"`, tone: 'ok' });
  };

  const importTemplate = async (file: File) => {
    try {
      const t = deserializeTemplate(await file.text());
      await indexedDbStore.save(t);
      await refreshSaved();
      setTemplate(t);
      setActiveFieldId(null);
      setStatus({ text: `Imported "${t.name}"`, tone: 'ok' });
    } catch (err) {
      setStatus({
        text: `Import failed: ${err instanceof Error ? err.message : String(err)}`,
        tone: 'error',
      });
    }
  };

  // The active file re-extracts immediately so the grid and by-field view track
  // edits. The whole batch is deferred: re-running every file on each keystroke
  // in the field editor would make typing crawl once a few workbooks are open.
  const result = useMemo(
    () => (workbook ? applyTemplate(workbook, template) : null),
    [workbook, template],
  );

  const deferredTemplate = useDeferredValue(template);
  const batchResults = useMemo(
    () => applyTemplateToAll(workbooks, deferredTemplate),
    [workbooks, deferredTemplate],
  );

  const resultsById = useMemo(
    () => new Map(workbooks.map((wb, i) => [wb.id, batchResults[i]])),
    [workbooks, batchResults],
  );

  const activeField = template.fields.find((f) => f.id === activeFieldId) ?? null;

  // Draw the active field's resolved region on the grid, but only when it lands
  // on the sheet currently in view.
  const highlight = useMemo(() => {
    if (!workbook || !activeField) return null;
    try {
      const region = resolveSelector(workbook, activeField.selector);
      return region.sheet.name === activeSheet ? region.range : null;
    } catch {
      return null;
    }
  }, [workbook, activeField, activeSheet]);

  const selectionLabel = selection ? encodeRange(selection.range) : null;

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <span className="brand-mark">▦</span>
          <span>Excel Template Extractor</span>
        </div>

        <div className="topbar-actions">
          <input
            ref={fileInput}
            type="file"
            accept=".xlsx,.xlsm,.xls,.csv,.tsv"
            multiple
            hidden
            onChange={(e) => {
              void openFiles([...(e.target.files ?? [])]);
              e.target.value = '';
            }}
          />
          <button className="btn" onClick={() => fileInput.current?.click()} disabled={loading}>
            {loading ? 'Reading…' : workbooks.length ? 'Add files' : 'Open workbooks'}
          </button>
          {progress ? (
            <span className="file-name mono">{progress}</span>
          ) : (
            workbooks.length === 1 && <span className="file-name mono">{workbooks[0].fileName}</span>
          )}
        </div>

        {status && <div className={`status status-${status.tone}`}>{status.text}</div>}
      </header>

      <main className="layout">
        <section className="grid-pane">
          {workbook && sheet ? (
            <>
              <FileBar
                workbooks={workbooks}
                activeId={activeFileId}
                failures={loadFailures}
                resultsById={resultsById}
                onSelect={selectFile}
                onRemove={removeFile}
                onDismissFailure={(name) =>
                  setLoadFailures((prev) => prev.filter((f) => f.fileName !== name))
                }
                onClear={clearFiles}
              />

              <div className="sheet-tabs">
                {workbook.sheets.map((s) => (
                  <button
                    key={s.name}
                    className={s.name === activeSheet ? 'is-on' : ''}
                    onClick={() => setActiveSheet(s.name)}
                  >
                    {s.name}
                    <span className="dim">
                      {s.rowCount}×{s.colCount}
                    </span>
                  </button>
                ))}
              </div>

              <SheetGrid
                sheet={sheet}
                selection={selection?.sheet === activeSheet ? selection.range : null}
                onSelectionChange={(range) => setSelection({ sheet: activeSheet, range })}
                highlight={highlight}
              />

              <div className="grid-footer">
                <span className="mono">
                  {selection && selection.sheet === activeSheet
                    ? `${selection.sheet}!${selectionLabel}`
                    : 'No selection — drag across cells, or shift-click to extend'}
                </span>
                <button
                  className="btn"
                  disabled={!selection || selection.sheet !== activeSheet}
                  onClick={addFieldFromSelection}
                >
                  + Add field from selection
                </button>
              </div>
            </>
          ) : (
            <div
              className="dropzone"
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => {
                e.preventDefault();
                void openFiles([...(e.dataTransfer.files ?? [])]);
              }}
            >
              <h2>Drop workbooks here</h2>
              <p>
                .xlsx, .xlsm, .xls, .csv and .tsv are supported. Drop several at once to run one
                template across all of them. Files are parsed entirely in your browser — nothing is
                uploaded.
              </p>
              <button className="btn" onClick={() => fileInput.current?.click()}>
                Choose files
              </button>
              {loadFailures.length > 0 && (
                <p className="result-error">
                  {loadFailures.map((f) => `${f.fileName}: ${f.error}`).join(' · ')}
                </p>
              )}
            </div>
          )}
        </section>

        <aside className="sidebar">
          <div className="tabs">
            <button className={tab === 'template' ? 'is-on' : ''} onClick={() => setTab('template')}>
              Template
            </button>
            <button className={tab === 'results' ? 'is-on' : ''} onClick={() => setTab('results')}>
              Results
              {result && result.fields.some((f) => !f.ok) && <span className="dot" />}
            </button>
          </div>

          {tab === 'template' ? (
            <div className="pane">
              <label className="lbl">
                Template name
                <input
                  className="inp"
                  value={template.name}
                  onChange={(e) => setTemplate({ ...template, name: e.target.value })}
                />
              </label>
              <label className="lbl">
                Description
                <input
                  className="inp"
                  value={template.description}
                  placeholder="What kind of file is this for?"
                  onChange={(e) => setTemplate({ ...template, description: e.target.value })}
                />
              </label>

              <div className="btn-row">
                <button className="btn" onClick={() => void saveTemplate()}>
                  Save
                </button>
                <button
                  className="btn-secondary"
                  onClick={() => {
                    setTemplate(emptyTemplate());
                    setActiveFieldId(null);
                  }}
                >
                  New
                </button>
                <button
                  className="btn-secondary"
                  onClick={() =>
                    downloadFile(
                      `${safeFileName(template.name)}.template.json`,
                      serializeTemplate(template),
                      'application/json',
                    )
                  }
                >
                  Export
                </button>
                <input
                  ref={importInput}
                  type="file"
                  accept=".json"
                  hidden
                  onChange={(e) => {
                    const f = e.target.files?.[0];
                    if (f) void importTemplate(f);
                    e.target.value = '';
                  }}
                />
                <button className="btn-secondary" onClick={() => importInput.current?.click()}>
                  Import
                </button>
              </div>

              <h3 className="section-title">Fields ({template.fields.length})</h3>
              {template.fields.length === 0 ? (
                <p className="empty">
                  Select a range in the grid and choose <strong>Add field from selection</strong>.
                </p>
              ) : (
                <ul className="field-list">
                  {template.fields.map((f) => {
                    const fieldResult = result?.fields.find((r) => r.fieldId === f.id);
                    return (
                      <li
                        key={f.id}
                        className={f.id === activeFieldId ? 'is-on' : ''}
                        onClick={() => setActiveFieldId(f.id === activeFieldId ? null : f.id)}
                      >
                        <div className="field-row">
                          <strong>{f.name}</strong>
                          {fieldResult && !fieldResult.ok && (
                            <span className="tag tag-error">!</span>
                          )}
                          <span className="tag">{f.selector.kind}</span>
                        </div>
                        <span className="mono dim">{describeSelector(f.selector)}</span>
                      </li>
                    );
                  })}
                </ul>
              )}

              {activeField && (
                <FieldEditor
                  field={activeField}
                  workbook={workbook}
                  selectionRange={selection ? encodeRange(selection.range) : null}
                  selectionSheet={selection?.sheet ?? null}
                  onChange={updateField}
                  onDelete={() => deleteField(activeField.id)}
                  onClose={() => setActiveFieldId(null)}
                />
              )}

              <h3 className="section-title">Saved templates ({saved.length})</h3>
              {saved.length === 0 ? (
                <p className="empty">Nothing saved yet.</p>
              ) : (
                <ul className="saved-list">
                  {saved.map((t) => (
                    <li key={t.id}>
                      <button className="link" onClick={() => void loadTemplate(t.id)}>
                        {t.name}
                        <span className="dim"> · {t.fields.length} fields</span>
                      </button>
                      <button
                        className="btn-ghost"
                        title={`Delete ${t.name}`}
                        onClick={() => void deleteTemplate(t)}
                      >
                        ✕
                      </button>
                    </li>
                  ))}
                </ul>
              )}
            </div>
          ) : (
            <div className="pane">
              {batchResults.length > 0 ? (
                <ResultsPanel
                  results={batchResults}
                  activeResult={result}
                  flatten={template.flatten}
                  stale={deferredTemplate !== template}
                  onFlattenChange={(flatten) =>
                    setTemplate((t) => ({ ...t, flatten, updatedAt: Date.now() }))
                  }
                />
              ) : (
                <p className="empty">Open a workbook to see extracted data.</p>
              )}
            </div>
          )}
        </aside>
      </main>
    </div>
  );
}
