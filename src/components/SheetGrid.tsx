import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { SheetData } from '../lib/types';
import { colToLetter, isInRange, normalizeRange, type CellRef, type RangeRef } from '../lib/range';

const ROW_H = 26;
const COL_W = 118;
const HEADER_H = 28;
const GUTTER_W = 58;
const OVERSCAN = 6;

interface Props {
  sheet: SheetData;
  selection: RangeRef | null;
  onSelectionChange: (range: RangeRef) => void;
  /** Region a selected field resolves to, drawn underneath the selection. */
  highlight?: RangeRef | null;
}

export function SheetGrid({ sheet, selection, onSelectionChange, highlight }: Props) {
  const scrollerRef = useRef<HTMLDivElement>(null);
  const [scroll, setScroll] = useState({ top: 0, left: 0 });
  const [viewport, setViewport] = useState({ height: 600, width: 800 });
  const dragAnchor = useRef<CellRef | null>(null);

  // Grow the grid past the used range so a user can point a template at cells
  // that are empty in this file but populated in the next one.
  const rowCount = Math.max(sheet.rowCount + 20, 40);
  const colCount = Math.max(sheet.colCount + 5, 12);

  useEffect(() => {
    const el = scrollerRef.current;
    if (!el) return;
    const measure = () => setViewport({ height: el.clientHeight, width: el.clientWidth });
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(el);
    return () => observer.disconnect();
  }, []);

  // Reset the scroll position when switching sheets, otherwise the new sheet
  // opens scrolled to wherever the previous one was.
  useEffect(() => {
    scrollerRef.current?.scrollTo({ top: 0, left: 0 });
    setScroll({ top: 0, left: 0 });
  }, [sheet.name]);

  const onScroll = useCallback(() => {
    const el = scrollerRef.current;
    if (el) setScroll({ top: el.scrollTop, left: el.scrollLeft });
  }, []);

  const startRow = Math.max(0, Math.floor(scroll.top / ROW_H) - OVERSCAN);
  const endRow = Math.min(rowCount, Math.ceil((scroll.top + viewport.height) / ROW_H) + OVERSCAN);
  const startCol = Math.max(0, Math.floor(scroll.left / COL_W) - OVERSCAN);
  const endCol = Math.min(colCount, Math.ceil((scroll.left + viewport.width) / COL_W) + OVERSCAN);

  const rows = useMemo(
    () => Array.from({ length: Math.max(endRow - startRow, 0) }, (_, i) => startRow + i),
    [startRow, endRow],
  );
  const cols = useMemo(
    () => Array.from({ length: Math.max(endCol - startCol, 0) }, (_, i) => startCol + i),
    [startCol, endCol],
  );

  const beginSelect = (ref: CellRef, shiftKey: boolean) => {
    if (shiftKey && selection) {
      onSelectionChange(normalizeRange({ start: selection.start, end: ref }));
      return;
    }
    dragAnchor.current = ref;
    onSelectionChange({ start: ref, end: ref });
  };

  /**
   * Only extends while the primary button is genuinely held. Trusting the drag
   * ref alone lets a stale drag (pointer capture loss, pointercancel, a
   * re-render under a resting cursor) redraw the selection on plain hover.
   */
  const extendSelect = (ref: CellRef, buttons: number) => {
    if (!dragAnchor.current) return;
    if ((buttons & 1) === 0) {
      dragAnchor.current = null;
      return;
    }
    onSelectionChange(normalizeRange({ start: dragAnchor.current, end: ref }));
  };

  useEffect(() => {
    const stop = () => {
      dragAnchor.current = null;
    };
    window.addEventListener('pointerup', stop);
    window.addEventListener('pointercancel', stop);
    return () => {
      window.removeEventListener('pointerup', stop);
      window.removeEventListener('pointercancel', stop);
    };
  }, []);

  return (
    <div className="grid-scroller" ref={scrollerRef} onScroll={onScroll}>
      <div
        className="grid-canvas"
        style={{ width: GUTTER_W + colCount * COL_W, height: HEADER_H + rowCount * ROW_H }}
      >
        {/* cells */}
        {rows.map((r) =>
          cols.map((c) => {
            const ref = { row: r, col: c };
            const selected = selection ? isInRange(ref, selection) : false;
            const inHighlight = highlight ? isInRange(ref, highlight) : false;
            const value = sheet.cells[r]?.[c] ?? null;
            const numeric = typeof value === 'number';

            return (
              <div
                key={`${r}:${c}`}
                className={[
                  'grid-cell',
                  selected ? 'is-selected' : '',
                  inHighlight ? 'is-highlighted' : '',
                  numeric ? 'is-numeric' : '',
                ]
                  .filter(Boolean)
                  .join(' ')}
                style={{
                  top: HEADER_H + r * ROW_H,
                  left: GUTTER_W + c * COL_W,
                  width: COL_W,
                  height: ROW_H,
                }}
                title={value === null ? '' : String(value)}
                onPointerDown={(e) => {
                  e.preventDefault();
                  beginSelect(ref, e.shiftKey);
                }}
                onPointerEnter={(e) => extendSelect(ref, e.buttons)}
              >
                {value === null ? '' : String(value)}
              </div>
            );
          }),
        )}

        {/* column headers */}
        <div className="grid-layer grid-col-headers" style={{ transform: `translateY(${scroll.top}px)` }}>
          {cols.map((c) => (
            <div
              key={c}
              className={`grid-head ${
                selection && c >= selection.start.col && c <= selection.end.col ? 'is-active' : ''
              }`}
              style={{ left: GUTTER_W + c * COL_W, width: COL_W, height: HEADER_H }}
            >
              {colToLetter(c)}
            </div>
          ))}
        </div>

        {/* row gutter */}
        <div className="grid-layer grid-row-gutter" style={{ transform: `translateX(${scroll.left}px)` }}>
          {rows.map((r) => (
            <div
              key={r}
              className={`grid-head ${
                selection && r >= selection.start.row && r <= selection.end.row ? 'is-active' : ''
              }`}
              style={{ top: HEADER_H + r * ROW_H, width: GUTTER_W, height: ROW_H }}
            >
              {r + 1}
            </div>
          ))}
        </div>

        <div
          className="grid-corner"
          style={{
            transform: `translate(${scroll.left}px, ${scroll.top}px)`,
            width: GUTTER_W,
            height: HEADER_H,
          }}
        />
      </div>
    </div>
  );
}
