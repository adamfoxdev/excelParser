import type { ExtractionResult, FieldResult } from '../lib/types';
import { downloadFile, resultToCsv, resultToJson, safeFileName } from '../lib/download';

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

export function ResultsPanel({ result }: { result: ExtractionResult }) {
  const failed = result.fields.filter((f) => !f.ok).length;
  const base = `${safeFileName(result.templateName)}_${safeFileName(result.fileName.replace(/\.[^.]+$/, ''))}`;

  return (
    <div className="results">
      <div className="results-bar">
        <span>
          {result.fields.length} field{result.fields.length === 1 ? '' : 's'}
          {failed > 0 && <span className="tag tag-error">{failed} failed</span>}
        </span>
        <div className="results-actions">
          <button
            className="btn-secondary"
            onClick={() => downloadFile(`${base}.json`, resultToJson(result), 'application/json')}
          >
            Export JSON
          </button>
          <button
            className="btn-secondary"
            onClick={() => downloadFile(`${base}.csv`, resultToCsv(result), 'text/csv')}
          >
            Export CSV
          </button>
        </div>
      </div>

      {result.fields.length === 0 ? (
        <p className="empty">Add a field to see extracted data here.</p>
      ) : (
        result.fields.map((f) => <FieldBlock key={f.fieldId} field={f} />)
      )}
    </div>
  );
}
