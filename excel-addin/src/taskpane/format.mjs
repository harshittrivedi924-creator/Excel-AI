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

/* ---------------------------------------------------------------------- */
/* Analysis rendering (pure helpers)                                       */
/* ---------------------------------------------------------------------- */

/** True when the text asks about the data rather than asking for a change. */
export function looksLikeAnalysisQuestion(text) {
  if ( typeof text !== "string") return false;
  return (
    /\b(analy[sz]e|analysis|analyses|report|insights?|profile|summary|trend|anomal\w*|outlier\w*|data quality|highest|lowest|growth)\b/i.test(
      text,
    )
  );
}

/** True when the text asks for a fresh full analysis of the current range. */
export function isAnalysisRequest(text) {
  if (typeof text !== "string") return false;
  return /\b(analy[sz]\w*|report|profile|insights?|summar\w*|important)\b/i.test(text);
}

const CHART_INTENT_RE =
  /\b(chart|graph|plot|visuali[sz])\b|\bbana(o|do|iye)?\b|\bbanao\b/i;

/** True when the text asks for a chart to be created. */
export function isChartRequest(text) {
  if (typeof text !== "string") return false;
  return CHART_INTENT_RE.test(text);
}

/* Phrasings that point back at the chart the analysis recommended, rather
 * than naming a column of their own. Deliberately a whitelist: anything not
 * recognised here keeps the normal command pipeline, so a chart command can
 * never be silently rerouted onto a different column. */
const RECOMMENDED_CHART_RE =
  /\b(recommend\w*|suggest\w*|propos\w+|yahi|usi|unhi|that|same|it)\b/i;

/* "sales ka chart", "chart of revenue" - an explicit subject the user chose. */
const EXPLICIT_CHART_SUBJECT_RE =
  /\bka\s+(chart|graph|plot)\b|\b(chart|graph|plot)\s+(of|for)\b/i;

/** True when a chart request refers back to the stored recommendation. */
export function refersToRecommendedChart(text) {
  if (!isChartRequest(text)) return false;
  if (EXPLICIT_CHART_SUBJECT_RE.test(text)) return false;
  return RECOMMENDED_CHART_RE.test(text);
}

/**
 * True when a typed command should be answered from the stored analysis
 * report instead of the command pipeline.
 *
 * A chart request only qualifies when it points at the recommendation and
 * names no column of its own, so "Sales ka chart bana do" still creates a
 * Sales chart (or reports the normal "no such column" error).
 */
export function shouldUseStoredAnalysis(text) {
  if (looksLikeAnalysisQuestion(text)) return true;
  return refersToRecommendedChart(text);
}

/** Format a numeric metric with thousands separators. */
export function formatMetric(value, digits = 2) {
  if (value === null || value === undefined || value === "") return "\u2014";
  if (typeof value !== "number" || !isFinite(value)) return String(value);
  const rounded = Number(value.toFixed(digits));
  const text = String(rounded);
  const [intPart, decPart] = text.split(".");
  const grouped = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return decPart ? `${grouped}.${decPart}` : grouped;
}

/** One-line profile summary shown at the top of the analysis card. */
export function analysisHeadline(report) {
  if (!report || !report.profile) return "No analysis available.";
  const { rows, columns, numeric_columns: numeric } = report.profile;
  if (!rows) return "There is no data in this range to analyze.";
  const base = `${rows} row\u00d7${columns} column range`;
  return numeric && numeric.length
    ? `${base} \u00b7 ${numeric.length} numeric column(s)`
    : `${base} \u00b7 no numeric column found`;
}

/** Rows for the statistics table: one per numeric column. */
export function analysisTableRows(report) {
  if (!report || !Array.isArray(report.statistics)) return [];
  return report.statistics
    .filter((s) => s && s.type === "numeric" && s.count)
    .map((s) => [
      s.column,
      formatMetric(s.sum),
      formatMetric(s.average),
      formatMetric(s.min),
      formatMetric(s.max),
    ]);
}

/** Human sentences for the growth section, one per numeric column. */
export function analysisGrowthLines(report) {
  if (!report || !Array.isArray(report.growth)) return [];
  return report.growth.map((g) => {
    if (!g.available) return `${g.column}: ${g.reason}`;
    if (g.percent === null || g.percent === undefined) {
      return `${g.column}: changed by ${formatMetric(g.change)}. ${g.reason || ""}`.trim();
    }
    const verb = g.change >= 0 ? "rose" : "fell";
    // Period labels come from the backend when the table has a label column,
    // so the card reads "Jan (1,000) -> May (2,000)" instead of bare numbers.
    const from = g.first_label ? `${g.first_label} (${formatMetric(g.first)})` : formatMetric(g.first);
    const to = g.last_label ? `${g.last_label} (${formatMetric(g.last)})` : formatMetric(g.last);
    return `${g.column} ${verb} ${formatMetric(Math.abs(g.percent))}% (${from} \u2192 ${to})`;
  });
}

/** Human sentences for the trend section. */
export function analysisTrendLines(report) {
  if (!report || !Array.isArray(report.trends)) return [];
  return report.trends.map((t) => t.detail);
}

/** Human sentences for the anomaly section, or an explicit "none" note. */
export function analysisAnomalyLines(report) {
  if (!report || !Array.isArray(report.anomalies)) return [];
  const lines = [];
  for (const finding of report.anomalies) {
    if (finding.count) {
      for (const item of finding.items) {
        const where = item.label ? ` at ${item.label}` : "";
        lines.push(
          `${finding.column}: ${formatMetric(item.value)}${where} (${item.methods.join(", ")})`,
        );
      }
    } else if (finding.note) {
      lines.push(finding.note);
    }
  }
  return lines;
}

/** Insight texts, capped so the card stays readable in a narrow pane. */
export function analysisInsightLines(report) {
  if (!report || !Array.isArray(report.insights)) return [];
  return report.insights.map((i) => (typeof i === "string" ? i : i.text)).filter(Boolean);
}

/**
 * Data-quality bullets: empty cells, duplicate rows, completeness and any
 * mixed-type columns. Always returns at least one line so the section never
 * silently disappears for a table that was profiled.
 */
export function analysisQualityLines(report) {
  const quality = report && report.data_quality;
  if (!quality) return [];
  const lines = [];
  const empty = quality.empty_cells || 0;
  const total = quality.total_cells || 0;
  lines.push(
    empty === 0
      ? "No empty cells found."
      : `${empty} empty cell${empty === 1 ? "" : "s"} out of ${total}.`,
  );
  const duplicates = quality.duplicate_rows || 0;
  lines.push(
    duplicates === 0
      ? "No duplicate rows."
      : `${duplicates} duplicate row${duplicates === 1 ? "" : "s"}.`,
  );
  if (quality.completeness_pct !== undefined && quality.completeness_pct !== null) {
    lines.push(`${quality.completeness_pct}% of cells are filled.`);
  }
  const mixed = quality.mixed_type_columns || [];
  if (mixed.length) {
    lines.push(`Mixed number/text values in: ${mixed.join(", ")}.`);
  }
  return lines;
}

/** Chart recommendation as display text plus the command to run it. */
export function analysisChartPlan(report) {
  const chart = report && report.chart_recommendation;
  if (!chart || !chart.available) {
    return { available: false, text: "", command: "" };
  }
  const kind = chart.type === "LineChart" ? "line" : "bar";
  return {
    available: true,
    text: chart.title || chart.columns.join(", "),
    command: `${chart.columns[0]} ka ${kind} chart bana do`,
  };
}
