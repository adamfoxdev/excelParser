import type { ExtractionResult, FieldResult, SqlTarget } from '../lib/types';
import { buildSqlScript, type SqlMode } from '../lib/sql';
import { downloadFile, resultToCsv, resultsToFlatCsv, resultsToJson, safeFileName } from '../lib/download';
import { flattenResults, isJsonCell } from '../lib/flatten';
import type { BatchFailure, BatchProgress } from '../lib/batch';

const PREVIEW_ROWS = 50;

function FieldBlock({ field }: { field: FieldResult }) {
  if (!field.ok) {
    return (
      <div className="result-block is-error">
        <div className="result-head">
          <strong>{field.fieldName}</strong>
          <span className="tag tag-error">failed</span>
        </div>
        <p className="result-error">{field.error}</p>
      </div>
    );
  }

  if (field.output === 'value') {
    return (
      <div className="result-block">
        <div className="result-head">
          <strong>{field.fieldName}</strong>
          <span className="mono dim">{field.resolvedRange}</span>
        </div>
        <p className="result-value">{field.value === null ? <em>empty</em> : String(field.value)}</p>
      </div>
    );
  }

  const shown = field.rows.slice(0, PREVIEW_ROWS);
  const hidden = field.rows.length - shown.length;

  return (
    <div className="result-block">
      <div className="result-head">
        <strong>{field.fieldName}</strong>
        <span className="mono dim">{field.resolvedRange}</span>
        <span className="tag">{field.rows.length} rows</span>
      </div>
      <div className="result-table-wrap">
        <table className="result-table">
          {field.headers && (
            <thead>
              <tr>
                {field.headers.map((h, i) => (
                  <th key={i}>{h}</th>
                ))}
              </tr>
            </thead>
          )}
          <tbody>
            {shown.map((row, r) => (
              <tr key={r}>
                {row.map((cell, c) => (
                  <td key={c} className={typeof cell === 'number' ? 'is-numeric' : ''}>
                    {cell === null ? '' : String(cell)}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {hidden > 0 && <p className="hint">+{hidden} more rows (all included in exports)</p>}
    </div>
  );
}

function FlatTableView({ results }: { results: ExtractionResult[] }) {
  const table = flattenResults(results);
  const fieldCount = table.headers.length - 1;

  return (
    <div className="result-block">
      <div className="result-head">
        <strong>Flat table</strong>
        <span className="tag">
          {table.rows.length} row{table.rows.length === 1 ? '' : 's'} × {fieldCount} field
          {fieldCount === 1 ? '' : 's'}
        </span>
      </div>
      <div className="result-table-wrap">
        <table className="result-table">
          <thead>
            <tr>
              {table.headers.map((h, i) => (
                <th key={i} className={i === 0 ? 'is-file' : ''}>
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {table.rows.map((row, r) => (
              <tr key={r}>
                {row.map((cell, c) => {
                  // Column 0 is the file name; the rest line up with the fields.
                  if (c === 0) {
                    return (
                      <td key={c} className="is-file" title={String(cell)}>
                        {String(cell)}
                      </td>
                    );
                  }
                  const field = results[r].fields[c - 1];
                  const text = cell === null ? '' : String(cell);
                  return (
                    <td
                      key={c}
                      className={[
                        typeof cell === 'number' ? 'is-numeric' : '',
                        isJsonCell(field) ? 'is-json' : '',
                        !field.ok ? 'is-failed' : '',
                      ]
                        .filter(Boolean)
                        .join(' ')}
                      title={field.ok ? text : field.error}
                    >
                      {field.ok ? text : '—'}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <p className="hint">
        One row per file. Multi-row fields are JSON-encoded into one cell; set a field's output
        to <strong>Single value</strong> for a bare value in its column.
      </p>
    </div>
  );
}

interface Props {
  /** A completed batch, or just the active file when no batch has run. */
  results: ExtractionResult[];
  /** The file shown in the grid — the one the by-field view details. */
  activeResult: ExtractionResult | null;
  fileCount: number;
  batchRan: boolean;
  batchStale: boolean;
  batchFailures: BatchFailure[];
  progress: BatchProgress | null;
  flatten: boolean;
  sql: SqlTarget;
  onRunBatch: () => void;
  onCancelBatch: () => void;
  onFlattenChange: (flatten: boolean) => void;
  onSqlChange: (sql: SqlTarget) => void;
}

function SqlOptions({
  sql,
  mode,
  tableCount,
  onChange,
}: {
  sql: SqlTarget;
  mode: SqlMode;
  tableCount: number;
  onChange: (sql: SqlTarget) => void;
}) {
  return (
    <div className="sql-options">
      <label className="lbl">
        Schema
        <input
          className="inp"
          value={sql.schema}
          onChange={(e) => onChange({ ...sql, schema: e.target.value })}
        />
      </label>
      <label className="lbl">
        Table
        <input
          className="inp"
          value={sql.table}
          onChange={(e) => onChange({ ...sql, table: e.target.value })}
        />
      </label>
      <label className="lbl chk">
        <input
          type="checkbox"
          checked={sql.dropExisting}
          onChange={(e) => onChange({ ...sql, dropExisting: e.target.checked })}
        />
        Drop if exists
      </label>
      <p className="hint sql-hint">
        {mode === 'flat'
          ? `One table (${sql.schema}.${sql.table}) — a row per file, multi-row fields as JSON.`
          : `${tableCount} table${tableCount === 1 ? '' : 's'} keyed by SourceFile.`}{' '}
        Same script for SQL Server and LocalDB.
      </p>
    </div>
  );
}

function BatchBar({
  fileCount,
  batchRan,
  batchStale,
  extracted,
  progress,
  onRun,
  onCancel,
}: {
  fileCount: number;
  batchRan: boolean;
  batchStale: boolean;
  extracted: number;
  progress: BatchProgress | null;
  onRun: () => void;
  onCancel: () => void;
}) {
  if (fileCount <= 1) return null;

  if (progress) {
    const pct = progress.total ? Math.round((progress.done / progress.total) * 100) : 0;
    return (
      <div className="batch-bar is-running">
        <div className="batch-progress" style={{ width: `${pct}%` }} />
        <span className="batch-text">
          Extracting {progress.done + 1} of {progress.total}
          {progress.fileName && <span className="dim"> · {progress.fileName}</span>}
        </span>
        <button className="btn-secondary" onClick={onCancel}>
          Cancel
        </button>
      </div>
    );
  }

  return (
    <div className={`batch-bar ${batchStale || !batchRan ? 'needs-run' : ''}`}>
      <span className="batch-text">
        {!batchRan
          ? `${fileCount} files loaded — only the open file is extracted so far`
          : batchStale
            ? `Template changed since the last run · showing ${extracted} file${extracted === 1 ? '' : 's'}`
            : `${extracted} of ${fileCount} files extracted`}
      </span>
      <button className="btn" onClick={onRun}>
        {batchRan ? 'Re-run batch' : `Run on all ${fileCount}`}
      </button>
    </div>
  );
}

export function ResultsPanel({
  results,
  activeResult,
  fileCount,
  batchRan,
  batchStale,
  batchFailures,
  progress,
  flatten,
  sql,
  onRunBatch,
  onCancelBatch,
  onFlattenChange,
  onSqlChange,
}: Props) {
  const failedFiles = results.filter((r) => r.fields.some((f) => !f.ok)).length;
  const fieldCount = results[0]?.fields.length ?? 0;
  // The SQL shape follows the view toggle, exactly as the CSV export does.
  const sqlMode: SqlMode = flatten ? 'flat' : 'normalized';
  const sqlTableCount =
    (results[0]?.fields.some((f) => f.output === 'value') ? 1 : 0) +
    (results[0]?.fields.filter((f) => f.output !== 'value').length ?? 0);
  const templateName = results[0]?.templateName ?? 'extraction';
  const base =
    results.length === 1
      ? `${safeFileName(templateName)}_${safeFileName(results[0].fileName.replace(/\.[^.]+$/, ''))}`
      : `${safeFileName(templateName)}_${results.length}_files`;

  return (
    <div className="results">
      <div className="seg">
        <button className={flatten ? '' : 'is-on'} onClick={() => onFlattenChange(false)}>
          By field
        </button>
        <button className={flatten ? 'is-on' : ''} onClick={() => onFlattenChange(true)}>
          Flat table
        </button>
      </div>

      <BatchBar
        fileCount={fileCount}
        batchRan={batchRan}
        batchStale={batchStale}
        extracted={batchRan ? results.length : 0}
        progress={progress}
        onRun={onRunBatch}
        onCancel={onCancelBatch}
      />

      {batchFailures.length > 0 && (
        <div className="result-block is-error">
          <div className="result-head">
            <strong>{batchFailures.length} file(s) could not be read</strong>
          </div>
          <ul className="failure-list">
            {batchFailures.slice(0, 10).map((f) => (
              <li key={f.fileId}>
                <span className="mono">{f.fileName}</span> — {f.error}
              </li>
            ))}
          </ul>
          {batchFailures.length > 10 && (
            <p className="hint">+{batchFailures.length - 10} more</p>
          )}
        </div>
      )}

      <div className="results-bar">
        <span>
          {results.length} file{results.length === 1 ? '' : 's'} · {fieldCount} field
          {fieldCount === 1 ? '' : 's'}
          {failedFiles > 0 && (
            <span className="tag tag-error">
              {failedFiles} file{failedFiles === 1 ? '' : 's'} with errors
            </span>
          )}
        </span>
        <div className="results-actions">
          <button
            className="btn-secondary"
            onClick={() => downloadFile(`${base}.json`, resultsToJson(results), 'application/json')}
          >
            Export JSON
          </button>
          <button
            className="btn-secondary"
            onClick={() =>
              flatten
                ? downloadFile(`${base}_flat.csv`, resultsToFlatCsv(results), 'text/csv')
                : activeResult &&
                  downloadFile(
                    `${safeFileName(templateName)}_${safeFileName(activeResult.fileName.replace(/\.[^.]+$/, ''))}.csv`,
                    resultToCsv(activeResult),
                    'text/csv',
                  )
            }
          >
            {flatten ? 'Export flat CSV' : 'Export CSV'}
          </button>
          <button
            className="btn-secondary"
            disabled={fieldCount === 0}
            onClick={() =>
              downloadFile(
                `${base}.sql`,
                buildSqlScript(results, { ...sql, mode: sqlMode }),
                'application/sql',
              )
            }
          >
            Export SQL
          </button>
        </div>
      </div>

      {fieldCount > 0 && (
        <SqlOptions
          sql={sql}
          mode={sqlMode}
          tableCount={sqlTableCount}
          onChange={onSqlChange}
        />
      )}

      {fieldCount === 0 ? (
        <p className="empty">Add a field to see extracted data here.</p>
      ) : flatten ? (
        <FlatTableView results={results} />
      ) : (
        <>
          {fileCount > 1 && (
            <p className="hint">
              Showing <strong>{activeResult?.fileName}</strong> — pick another file above, or
              switch to <strong>Flat table</strong> for the whole batch.
            </p>
          )}
          {activeResult?.fields.map((f) => <FieldBlock key={f.fieldId} field={f} />)}
        </>
      )}
    </div>
  );
}
