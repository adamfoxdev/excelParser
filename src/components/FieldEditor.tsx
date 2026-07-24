import type { AnchorSelector, Field, FixedSelector, MatchMode, WorkbookData } from '../lib/types';
import { ANY_SHEET } from '../lib/types';

interface Props {
  field: Field;
  workbook: WorkbookData | null;
  selectionRange: string | null;
  selectionSheet: string | null;
  onChange: (field: Field) => void;
  onDelete: () => void;
  onClose: () => void;
}

const DEFAULT_ANCHOR: Omit<AnchorSelector, 'kind' | 'sheet'> = {
  anchorText: '',
  matchMode: 'exact',
  caseSensitive: false,
  occurrence: 1,
  offsetRows: 1,
  offsetCols: 0,
  height: 'auto',
  width: 'auto',
};

export function FieldEditor({
  field,
  workbook,
  selectionRange,
  selectionSheet,
  onChange,
  onDelete,
  onClose,
}: Props) {
  const set = (patch: Partial<Field>) => onChange({ ...field, ...patch });

  const setFixed = (patch: Partial<FixedSelector>) => {
    if (field.selector.kind !== 'fixed') return;
    set({ selector: { ...field.selector, ...patch } });
  };

  const setAnchor = (patch: Partial<AnchorSelector>) => {
    if (field.selector.kind !== 'anchor') return;
    set({ selector: { ...field.selector, ...patch } });
  };

  const switchKind = (kind: 'fixed' | 'anchor') => {
    if (kind === field.selector.kind) return;
    if (kind === 'fixed') {
      set({
        selector: {
          kind: 'fixed',
          sheet: selectionSheet ?? workbook?.sheets[0]?.name ?? 'Sheet1',
          range: selectionRange ?? 'A1:A1',
        },
      });
    } else {
      set({
        selector: {
          kind: 'anchor',
          sheet: field.selector.kind === 'fixed' ? field.selector.sheet : ANY_SHEET,
          ...DEFAULT_ANCHOR,
        },
      });
    }
  };

  const dimValue = (v: number | 'auto') => (v === 'auto' ? '' : String(v));
  const parseDim = (raw: string): number | 'auto' => {
    const n = parseInt(raw, 10);
    return raw.trim() === '' || Number.isNaN(n) || n < 1 ? 'auto' : n;
  };

  return (
    <div className="editor">
      <div className="editor-head">
        <h3>Edit field</h3>
        <button className="btn-ghost" onClick={onClose} title="Close editor">
          ✕
        </button>
      </div>

      <label className="lbl">
        Name
        <input
          className="inp"
          value={field.name}
          onChange={(e) => set({ name: e.target.value })}
          placeholder="e.g. Line items"
        />
      </label>

      <div className="seg">
        <button
          className={field.selector.kind === 'fixed' ? 'is-on' : ''}
          onClick={() => switchKind('fixed')}
        >
          Fixed range
        </button>
        <button
          className={field.selector.kind === 'anchor' ? 'is-on' : ''}
          onClick={() => switchKind('anchor')}
        >
          Anchored
        </button>
      </div>

      {field.selector.kind === 'fixed' ? (
        <>
          <div className="row">
            <label className="lbl">
              Sheet
              <select
                className="inp"
                value={field.selector.sheet}
                onChange={(e) => setFixed({ sheet: e.target.value })}
              >
                {workbook?.sheets.map((s) => (
                  <option key={s.name} value={s.name}>
                    {s.name}
                  </option>
                )) ?? <option value={field.selector.sheet}>{field.selector.sheet}</option>}
              </select>
            </label>
            <label className="lbl">
              Range
              <input
                className="inp mono"
                value={field.selector.range}
                onChange={(e) => setFixed({ range: e.target.value.toUpperCase() })}
                placeholder="B2:D50"
              />
            </label>
          </div>
          <button
            className="btn-secondary full"
            disabled={!selectionRange}
            onClick={() =>
              setFixed({ range: selectionRange!, sheet: selectionSheet ?? field.selector.sheet })
            }
          >
            {selectionRange ? `Use grid selection (${selectionRange})` : 'Select cells in the grid'}
          </button>
        </>
      ) : (
        <>
          <label className="lbl">
            Anchor text
            <input
              className="inp"
              value={field.selector.anchorText}
              onChange={(e) => setAnchor({ anchorText: e.target.value })}
              placeholder="e.g. Invoice Total"
            />
          </label>

          <div className="row">
            <label className="lbl">
              Match
              <select
                className="inp"
                value={field.selector.matchMode}
                onChange={(e) => setAnchor({ matchMode: e.target.value as MatchMode })}
              >
                <option value="exact">Exact</option>
                <option value="contains">Contains</option>
                <option value="regex">Regex</option>
              </select>
            </label>
            <label className="lbl">
              Sheet
              <select
                className="inp"
                value={field.selector.sheet}
                onChange={(e) => setAnchor({ sheet: e.target.value })}
              >
                <option value={ANY_SHEET}>Any sheet</option>
                {workbook?.sheets.map((s) => (
                  <option key={s.name} value={s.name}>
                    {s.name}
                  </option>
                ))}
              </select>
            </label>
          </div>

          <div className="row">
            <label className="lbl">
              Occurrence
              <input
                className="inp"
                type="number"
                min={1}
                value={field.selector.occurrence}
                onChange={(e) =>
                  setAnchor({ occurrence: Math.max(1, parseInt(e.target.value, 10) || 1) })
                }
              />
            </label>
            <label className="lbl chk">
              <input
                type="checkbox"
                checked={field.selector.caseSensitive}
                onChange={(e) => setAnchor({ caseSensitive: e.target.checked })}
              />
              Case sensitive
            </label>
          </div>

          <div className="row">
            <label className="lbl">
              Offset rows
              <input
                className="inp"
                type="number"
                value={field.selector.offsetRows}
                onChange={(e) => setAnchor({ offsetRows: parseInt(e.target.value, 10) || 0 })}
              />
            </label>
            <label className="lbl">
              Offset cols
              <input
                className="inp"
                type="number"
                value={field.selector.offsetCols}
                onChange={(e) => setAnchor({ offsetCols: parseInt(e.target.value, 10) || 0 })}
              />
            </label>
          </div>

          <div className="row">
            <label className="lbl">
              Height
              <input
                className="inp"
                type="number"
                min={1}
                value={dimValue(field.selector.height)}
                onChange={(e) => setAnchor({ height: parseDim(e.target.value) })}
                placeholder="auto"
              />
            </label>
            <label className="lbl">
              Width
              <input
                className="inp"
                type="number"
                min={1}
                value={dimValue(field.selector.width)}
                onChange={(e) => setAnchor({ width: parseDim(e.target.value) })}
                placeholder="auto"
              />
            </label>
          </div>
          <p className="hint">
            Blank height/width grows the region until a blank row or column ends it.
          </p>
        </>
      )}

      <hr className="rule" />

      <div className="row">
        <label className="lbl">
          Output as
          <select
            className="inp"
            value={field.output}
            onChange={(e) => set({ output: e.target.value as Field['output'] })}
          >
            <option value="table">Table</option>
            <option value="list">List</option>
            <option value="value">Single value</option>
          </select>
        </label>
        <label className="lbl chk">
          <input
            type="checkbox"
            checked={field.headerRow}
            onChange={(e) => set({ headerRow: e.target.checked })}
          />
          First row is headers
        </label>
      </div>

      <button className="btn-danger full" onClick={onDelete}>
        Delete field
      </button>
    </div>
  );
}
