/* Pure snapshot helpers for the Excel Add-in task pane.
 *
 * These functions are deliberately free of any Office.js / Excel / DOM
 * globals so they can be unit-tested with the plain Node test runner.
 *
 * The worksheet snapshot is rebuilt from the raw 2-D ``values`` array
 * Office.js returns. Emptiness is decided by inspecting the actual cell
 * contents, never by trusting ``getUsedRangeOrNullObject().isNullObject``
 * (which has shipped inconsistent across Office.js builds and can report a
 * non-empty worksheet as a null object).
 */

export function blankValue(value) {
  return value === null || value === undefined || value === "";
}

/** True when a 2-D values array contains no real data (all cells blank). */
export function isEmptyValues(values) {
  if (!Array.isArray(values) || values.length === 0) return true;
  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    if (!Array.isArray(row)) continue;
    for (let c = 0; c < row.length; c++) {
      if (!blankValue(row[c])) return false;
    }
  }
  return true;
}

/** Bounds (0-based) of the first/last non-blank cell, or null when empty. */
export function trimBounds(values) {
  if (!Array.isArray(values)) return null;
  const rowCount = values.length;
  if (rowCount === 0) return null;

  let colCount = 0;
  for (let r = 0; r < rowCount; r++) {
    const row = values[r];
    if (Array.isArray(row) && row.length > colCount) colCount = row.length;
  }
  if (colCount === 0) return null;

  let startRow = rowCount;
  let endRow = -1;
  let startCol = colCount;
  let endCol = -1;

  for (let r = 0; r < rowCount; r++) {
    const row = values[r] || [];
    for (let c = 0; c < row.length; c++) {
      if (!blankValue(row[c])) {
        if (r < startRow) startRow = r;
        if (r > endRow) endRow = r;
        if (c < startCol) startCol = c;
        if (c > endCol) endCol = c;
      }
    }
  }

  if (endRow < 0) return null;
  return { startRow, startCol, endRow, endCol };
}

/**
 * Build a trimmed snapshot from a raw 2-D values array.
 *
 * Options:
 *   baseRow / baseCol  absolute (1-based) worksheet location of values[0][0].
 *   maxRows / maxCols  hard caps; an oversized snapshot is flagged truncated.
 *
 * Returns null when the values contain no data, otherwise:
 *   { values, rows, cols, originRow, originCol, truncated }
 */
export function buildSnapshot(values, opts = {}) {
  const maxRows = opts.maxRows ?? 3000;
  const maxCols = opts.maxCols ?? 100;
  const baseRow = opts.baseRow ?? 1;
  const baseCol = opts.baseCol ?? 1;

  const bounds = trimBounds(values);
  if (!bounds) return null;

  const height = bounds.endRow - bounds.startRow + 1;
  const width = bounds.endCol - bounds.startCol + 1;
  const rows = Math.min(height, maxRows);
  const cols = Math.min(width, maxCols);

  const region = [];
  for (let r = 0; r < rows; r++) {
    const src = values[bounds.startRow + r] || [];
    const out = [];
    for (let c = 0; c < cols; c++) {
      const v = src[bounds.startCol + c];
      out.push(v === undefined ? null : v);
    }
    region.push(out);
  }

  return {
    values: region,
    rows: region.length,
    cols: region.length ? region[0].length : 0,
    originRow: baseRow + bounds.startRow,
    originCol: baseCol + bounds.startCol,
    truncated: height > maxRows || width > maxCols,
  };
}