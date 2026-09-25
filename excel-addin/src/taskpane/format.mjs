/* Excel AI Copilot -- pure formatting helpers for the task pane.
 *
 * Kept free of any DOM / Office.js dependencies so the behaviour is fully
 * unit-testable with `node --test`.
 */

/**
 * Add thousands separators to standalone numbers embeddd in a backend
 * message. Cell references ("B7"), formulas ("=SUM(A2:A4)") and short
 * numbers are left untouched so the readable text never gets corrupted.
 */
const STANDALONE_NUMBER_RE = /(?<![A-Za-z0-9_])(\d+)(?:\.(\d+))?(?![A-Za-z0-9_])/g;

export function groupThousands(text) {
  if (typeof text !== "string") return "";
  return text.replace(STANDALONE_NUMBER_RE, (match, intPart, decPart) => {
    if (intPart.length <= 3) return match;
    const grouped = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    return decPart !== undefined ? `${grouped}.${decPart}` : grouped;
  });
}

/**
 * Pick a short emoji for a successful backend response, based on what the
 * response actually contains (charts / result sheets / dedupe / default).
 */
export function successIcon(data) {
  const writes = data && data.writes ? data.writes : {};
  if (Array.isArray(writes.charts) && writes.charts.length) return "\u{1F4CA}";
  if (writes.sheets && Object.keys(writes.sheets).length) return "\u{1F4C4}";
  if (data && data.operation === "DEDUPE") return "\u{1F9F9}";
  return "\u2705";
}

/** Human-readable success message (never raw backend JSON). */
export function humanizeSuccess(data) {
  if (!data) return "\u2705 Done.";
  const raw = String(data.message || "Done.");
  return `${successIcon(data)} ${groupThousands(raw)}`;
}

/** Concise, user-friendly error message with an error marker. */
export function humanizeError(message) {
  const text = String(message || "Something went wrong. Try again.");
  return text.startsWith("\u274C") ? text : `\u274C ${text}`;
}