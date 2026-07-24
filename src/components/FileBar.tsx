import type { ExtractionResult, WorkbookData } from '../lib/types';
import { failedFieldCount } from '../lib/extract';
import type { LoadFailure } from '../lib/workbook';

interface Props {
  workbooks: WorkbookData[];
  activeId: string | null;
  failures: LoadFailure[];
  /** Per-file extraction, keyed by workbook id, for the error badges. */
  resultsById: Map<string, ExtractionResult>;
  onSelect: (id: string) => void;
  onRemove: (id: string) => void;
  onDismissFailure: (fileName: string) => void;
  onClear: () => void;
}

export function FileBar({
  workbooks,
  activeId,
  failures,
  resultsById,
  onSelect,
  onRemove,
  onDismissFailure,
  onClear,
}: Props) {
  if (workbooks.length <= 1 && failures.length === 0) return null;

  return (
    <div className="file-bar">
      <span className="file-bar-count dim">
        {workbooks.length} file{workbooks.length === 1 ? '' : 's'}
      </span>

      <div className="file-chips">
        {workbooks.map((wb) => {
          const result = resultsById.get(wb.id);
          const failed = result ? failedFieldCount(result) : 0;
          return (
            <span
              key={wb.id}
              className={`file-chip ${wb.id === activeId ? 'is-on' : ''} ${failed ? 'has-error' : ''}`}
            >
              <button className="file-chip-name" onClick={() => onSelect(wb.id)} title={wb.fileName}>
                {wb.fileName}
                {failed > 0 && <span className="tag tag-error">{failed}</span>}
              </button>
              <button
                className="btn-ghost"
                title={`Remove ${wb.fileName}`}
                onClick={() => onRemove(wb.id)}
              >
                ✕
              </button>
            </span>
          );
        })}

        {failures.map((f) => (
          <span key={f.fileName} className="file-chip is-unreadable" title={f.error}>
            <span className="file-chip-name">
              {f.fileName}
              <span className="tag tag-error">unreadable</span>
            </span>
            <button className="btn-ghost" onClick={() => onDismissFailure(f.fileName)}>
              ✕
            </button>
          </span>
        ))}
      </div>

      {workbooks.length > 0 && (
        <button className="btn-ghost file-bar-clear" onClick={onClear}>
          Clear all
        </button>
      )}
    </div>
  );
}
