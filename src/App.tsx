import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { SheetGrid } from './components/SheetGrid';
import { FieldEditor } from './components/FieldEditor';
import { ResultsPanel } from './components/ResultsPanel';
import { FileBar } from './components/FileBar';
import { loadWorkbook } from './lib/workbook';
import { applyTemplate, describeSelector, resolveSelector } from './lib/extract';
import {
  runBatch,
  templateSignature,
  toBatchFile,
  type BatchFailure,
  type BatchFile,
  type BatchProgress,
} from './lib/batch';
import { collectFromList, describeCollection, filesFromDrop, type Collected } from './lib/files';
import { encodeRange, type RangeRef } from './lib/range';
import {
  deserializeTemplate,
  emptyTemplate,
  indexedDbStore,
  newId,
  serializeTemplate,
} from './lib/storage';
import { downloadFile, safeFileName } from './lib/download';
import type { ExtractionResult, Field, Template, WorkbookData } from './lib/types';
import './App.css';

interface Selection {
  sheet: string;
  range: RangeRef;
}

/** A finished batch run, with the template it was produced from. */
interface BatchState {
  results: ExtractionResult[];
  failures: BatchFailure[];
  signature: string;
  cancelled: boolean;
}

export default function App() {
  const [files, setFiles] = useState<BatchFile[]>([]);
  const [activeFileId, setActiveFileId] = useState<string | null>(null);
  /** Only the file on screen is parsed; the rest stay as File handles. */
  const [workbook, setWorkbook] = useState<WorkbookData | null>(null);
  const [workbookError, setWorkbookError] = useState<string | null>(null);
  const [batch, setBatch] = useState<BatchState | null>(null);
  const [progress, setProgress] = useState<BatchProgress | null>(null);
  const abortRef = useRef<AbortController | null>(null);
  const [activeSheet, setActiveSheet] = useState('');
  const [selection, setSelection] = useState<Selection | null>(null);
  const [template, setTemplate] = useState<Template>(emptyTemplate);
  const [saved, setSaved] = useState<Template[]>([]);
  const [activeFieldId, setActiveFieldId] = useState<string | null>(null);
  const [tab, setTab] = useState<'template' | 'results'>('template');
  const [status, setStatus] = useState<{ text: string; tone: 'ok' | 'error' } | null>(null);
  const [loading, setLoading] = useState(false);

  const fileInput = useRef<HTMLInputElement>(null);
  const folderInput = useRef<HTMLInputElement>(null);
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

  const sheet = useMemo(
    () => workbook?.sheets.find((s) => s.name === activeSheet) ?? null,
    [workbook, activeSheet],
  );

  // Parse whichever file is active, and only that one. Switching files releases
  // the previous workbook, so memory tracks the largest single file rather than
  // the size of the batch.
  useEffect(() => {
    const entry = files.find((f) => f.id === activeFileId);
    if (!entry) {
      setWorkbook(null);
      setWorkbookError(null);
      return;
    }

    let cancelled = false;
    setLoading(true);
    setWorkbookError(null);

    loadWorkbook(entry.file)
      .then((wb) => {
        if (cancelled) return;
        setWorkbook(wb);
        // Hold the sheet name across files where it exists — templates are
        // usually built against a consistently named sheet.
        setActiveSheet((current) =>
          wb.sheets.some((s) => s.name === current) ? current : wb.sheets[0].name,
        );
      })
      .catch((err: unknown) => {
        if (cancelled) return;
        setWorkbook(null);
        setWorkbookError(err instanceof Error ? err.message : String(err));
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });

    return () => {
      cancelled = true;
    };
  }, [files, activeFileId]);

  const addFiles = (collected: Collected) => {
    if (collected.files.length === 0) {
      setStatus({
        text: describeCollection(collected) ?? 'No spreadsheet files in that drop.',
        tone: 'error',
      });
      return;
    }

    const added = collected.files.map(toBatchFile);
    setFiles((prev) => {
      if (prev.length === 0) {
        setActiveFileId(added[0].id);
        setSelection(null);
      }
      return [...prev, ...added];
    });

    const note = describeCollection(collected);
    setStatus({
      text: `Added ${added.length} file${added.length === 1 ? '' : 's'}${note ? ` · ${note}` : ''}`,
      tone: collected.truncated ? 'error' : 'ok',
    });
  };

  const removeFile = (id: string) => {
    setFiles((prev) => {
      const next = prev.filter((f) => f.id !== id);
      if (id === activeFileId) {
        setActiveFileId(next[0]?.id ?? null);
        setSelection(null);
      }
      return next;
    });
    setBatch((prev) =>
      prev
        ? {
            ...prev,
            results: prev.results.filter(
              (r) => r.fileName !== files.find((f) => f.id === id)?.name,
            ),
            failures: prev.failures.filter((f) => f.fileId !== id),
          }
        : prev,
    );
  };

  const selectFile = (id: string) => {
    if (!files.some((f) => f.id === id)) return;
    setActiveFileId(id);
    setSelection(null);
  };

  const clearFiles = () => {
    abortRef.current?.abort();
    setFiles([]);
    setActiveFileId(null);
    setActiveSheet('');
    setSelection(null);
    setBatch(null);
    setProgress(null);
  };

  const startBatch = async () => {
    abortRef.current?.abort();
    const controller = new AbortController();
    abortRef.current = controller;
    const signature = templateSignature(template);

    setProgress({ done: 0, total: files.length, fileName: '' });
    const outcome = await runBatch(files, template, {
      signal: controller.signal,
      onProgress: setProgress,
    });
    setProgress(null);
    if (controller !== abortRef.current) return; // superseded by a newer run

    setBatch({ ...outcome, signature });
    setStatus({
      text: outcome.cancelled
        ? `Stopped after ${outcome.results.length} of ${files.length} files`
        : `Extracted ${outcome.results.length} file${outcome.results.length === 1 ? '' : 's'}${
            outcome.failures.length ? ` · ${outcome.failures.length} unreadable` : ''
          }`,
      tone: outcome.failures.length || outcome.cancelled ? 'error' : 'ok',
    });
  };

  const cancelBatch = () => abortRef.current?.abort();

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

  // The active file re-extracts on every edit so authoring stays live. The rest
  // of the batch is only touched by an explicit run — re-parsing every file on
  // each keystroke would be unusable, and defeats the point of not holding them.
  const result = useMemo(
    () => (workbook ? applyTemplate(workbook, template) : null),
    [workbook, template],
  );

  const batchStale = batch !== null && batch.signature !== templateSignature(template);

  /** What the Results tab shows: a completed batch, else the live active file. */
  const displayedResults = useMemo(
    () => (batch ? batch.results : result ? [result] : []),
    [batch, result],
  );

  const resultsByName = useMemo(
    () => new Map((batch?.results ?? []).map((r) => [r.fileName, r])),
    [batch],
  );

  const resultsById = useMemo(() => {
    const map = new Map<string, ExtractionResult>();
    for (const entry of files) {
      const found =
        entry.id === activeFileId && result ? result : resultsByName.get(entry.name);
      if (found) map.set(entry.id, found);
    }
    return map;
  }, [files, activeFileId, result, resultsByName]);

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
              addFiles(collectFromList(e.target.files ?? []));
              e.target.value = '';
            }}
          />
          <input
            ref={folderInput}
            type="file"
            hidden
            // Not in React's typings; the directory picker needs the raw attributes.
            {...({ webkitdirectory: '', directory: '' } as Record<string, string>)}
            onChange={(e) => {
              addFiles(collectFromList(e.target.files ?? []));
              e.target.value = '';
            }}
          />
          <button className="btn" onClick={() => fileInput.current?.click()}>
            {files.length ? 'Add files' : 'Open workbooks'}
          </button>
          <button className="btn-secondary" onClick={() => folderInput.current?.click()}>
            Add folder
          </button>
          {files.length === 1 && <span className="file-name mono">{files[0].name}</span>}
        </div>

        {status && <div className={`status status-${status.tone}`}>{status.text}</div>}
      </header>

      <main className="layout">
        <section className="grid-pane">
          {files.length > 0 ? (
            <>
              <FileBar
                files={files}
                activeId={activeFileId}
                failures={batch?.failures ?? []}
                resultsById={resultsById}
                onSelect={selectFile}
                onRemove={removeFile}
                onClear={clearFiles}
              />

              {!sheet ? (
                <div className="dropzone">
                  {loading ? (
                    <h2>Reading {files.find((f) => f.id === activeFileId)?.name}…</h2>
                  ) : (
                    <>
                      <h2>Could not read this file</h2>
                      <p className="result-error">{workbookError}</p>
                      <p>Pick another file above, or remove this one.</p>
                    </>
                  )}
                </div>
              ) : (
                <>
              <div className="sheet-tabs">
                {workbook!.sheets.map((s) => (
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
              )}
            </>
          ) : (
            <div
              className="dropzone"
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => {
                e.preventDefault();
                void filesFromDrop(e.dataTransfer).then(addFiles);
              }}
            >
              <h2>Drop workbooks or a folder here</h2>
              <p>
                .xlsx, .xlsm, .xls, .csv and .tsv are supported. Drop a whole folder to run one
                template across everything in it — subfolders included. Files are parsed entirely
                in your browser, one at a time; nothing is uploaded.
              </p>
              <div className="btn-row">
                <button className="btn" onClick={() => fileInput.current?.click()}>
                  Choose files
                </button>
                <button className="btn-secondary" onClick={() => folderInput.current?.click()}>
                  Choose folder
                </button>
              </div>
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
              {files.length > 0 ? (
                <ResultsPanel
                  results={displayedResults}
                  activeResult={result}
                  fileCount={files.length}
                  batchRan={batch !== null}
                  batchStale={batchStale}
                  batchFailures={batch?.failures ?? []}
                  progress={progress}
                  flatten={template.flatten}
                  sql={template.sql}
                  onRunBatch={() => void startBatch()}
                  onCancelBatch={cancelBatch}
                  onFlattenChange={(flatten) =>
                    setTemplate((t) => ({ ...t, flatten, updatedAt: Date.now() }))
                  }
                  onSqlChange={(sql) => setTemplate((t) => ({ ...t, sql, updatedAt: Date.now() }))}
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
