import type { ExtractionResult } from '../lib/types';
import { failedFieldCount } from '../lib/extract';
import type { BatchFailure, BatchFile } from '../lib/batch';

interface Props {
  files: BatchFile[];
  activeId: string | null;
  /** Files the last batch run could not parse. */
  failures: BatchFailure[];
  /** Extraction per file id — only populated for files already extracted. */
  resultsById: Map<string, ExtractionResult>;
  onSelect: (id: string) => void;
  onRemove: (id: string) => void;
  onClear: () => void;
}

export function FileBar({
  files,
  activeId,
  failures,
  resultsById,
  onSelect,
  onRemove,
  onClear,
}: Props) {
  if (files.length <= 1) return null;

  const failedIds = new Set(failures.map((f) => f.fileId));

  return (
    <div className="file-bar">
      <span className="file-bar-count dim">{files.length} files</span>

      <div className="file-chips">
        {files.map((entry) => {
          const result = resultsById.get(entry.id);
          const failedFields = result ? failedFieldCount(result) : 0;
          const unreadable = failedIds.has(entry.id);

          return (
            <span
              key={entry.id}
              className={[
                'file-chip',
                entry.id === activeId ? 'is-on' : '',
                unreadable ? 'is-unreadable' : failedFields ? 'has-error' : '',
              ]
                .filter(Boolean)
                .join(' ')}
              title={
                unreadable
                  ? failures.find((f) => f.fileId === entry.id)?.error
                  : entry.name
              }
            >
              <button className="file-chip-name" onClick={() => onSelect(entry.id)}>
                {entry.name}
                {unreadable && <span className="tag tag-error">unreadable</span>}
                {!unreadable && failedFields > 0 && (
                  <span className="tag tag-error">{failedFields}</span>
                )}
              </button>
              <button
                className="btn-ghost"
                title={`Remove ${entry.name}`}
                onClick={() => onRemove(entry.id)}
              >
                ✕
              </button>
            </span>
          );
        })}
      </div>

      <button className="btn-ghost file-bar-clear" onClick={onClear}>
        Clear all
      </button>
    </div>
  );
}
