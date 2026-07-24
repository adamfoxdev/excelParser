import type { ExtractionResult, FieldResult } from '../lib/types';
import { downloadFile, resultToCsv, resultsToFlatCsv, resultsToJson, safeFileName } from '../lib/download';
import { flattenResults, isJsonCell } from '../lib/flatten';

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
  /** Every loaded file, extracted with the current template. */
  results: ExtractionResult[];
  /** The file shown in the grid — the one the by-field view details. */
  activeResult: ExtractionResult | null;
  flatten: boolean;
  onFlattenChange: (flatten: boolean) => void;
  stale: boolean;
}

export function ResultsPanel({ results, activeResult, flatten, onFlattenChange, stale }: Props) {
  const failedFiles = results.filter((r) => r.fields.some((f) => !f.ok)).length;
  const fieldCount = results[0]?.fields.length ?? 0;
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

      <div className="results-bar">
        <span>
          {results.length} file{results.length === 1 ? '' : 's'} · {fieldCount} field
          {fieldCount === 1 ? '' : 's'}
          {failedFiles > 0 && (
            <span className="tag tag-error">
              {failedFiles} file{failedFiles === 1 ? '' : 's'} with errors
            </span>
          )}
          {stale && <span className="tag">updating…</span>}
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
        </div>
      </div>

      {fieldCount === 0 ? (
        <p className="empty">Add a field to see extracted data here.</p>
      ) : flatten ? (
        <FlatTableView results={results} />
      ) : (
        <>
          {results.length > 1 && (
            <p className="hint">
              Showing <strong>{activeResult?.fileName}</strong> — pick another file above, or
              switch to <strong>Flat table</strong> for all {results.length} at once.
            </p>
          )}
          {activeResult?.fields.map((f) => <FieldBlock key={f.fieldId} field={f} />)}
        </>
      )}
    </div>
  );
}
