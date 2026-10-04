import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import {
  groupThousands,
  successIcon,
  humanizeSuccess,
  humanizeError,
  looksLikeAnalysisQuestion,
  isAnalysisRequest,
  isChartRequest,
  refersToRecommendedChart,
  shouldUseStoredAnalysis,
  formatMetric,
  analysisHeadline,
  analysisTableRows,
  analysisGrowthLines,
  analysisTrendLines,
  analysisAnomalyLines,
  analysisInsightLines,
  analysisQualityLines,
  analysisChartPlan,
} from "../src/taskpane/format.mjs";

test("groupThousands formats standalone integers", () => {
  assert.equal(groupThousands("The result is 18000."), "The result is 18,000.");
  assert.equal(groupThousands("Sum of B2:B7 is 145577390."), "Sum of B2:B7 is 145,577,390.");
  assert.equal(groupThousands("total comes to 7500"), "total comes to 7,500");
});

test("groupThousands keeps decimals", () => {
  assert.equal(groupThousands("average is 12345.5"), "average is 12,345.5");
  assert.equal(groupThousands("percent 0.05"), "percent 0.05");
});

test("groupThousands leaves short numbers and cell references untouched", () => {
  assert.equal(groupThousands("I wrote 6 formula(s) to column D."), "I wrote 6 formula(s) to column D.");
  assert.equal(groupThousands("B7 already contains the value B7"), "B7 already contains the value B7");
  assert.equal(groupThousands("=SUM(A2:A4)"), "=SUM(A2:A4)");
  assert.equal(groupThousands("Revenue"), "Revenue");
});

test("successIcon reflects charts, sheets and dedupe", () => {
  assert.equal(successIcon({}), "✅");
  assert.equal(successIcon({ operation: "DEDUPE" }), "🧹");
  assert.equal(successIcon({ writes: { charts: [{ type: "ColumnClustered" }] } }), "📊");
  assert.equal(successIcon({ writes: { sheets: { Filtered: {} } } }), "📄");
  assert.equal(successIcon({ writes: { cells: [] } }), "✅");
});

test("humanizeSuccess wraps the readable backend message", () => {
  const text = humanizeSuccess({
    message: "Done. The result of Sum of A2:A4 is 18000.",
    operation: "SUM",
    writes: { cells: [], sheets: {}, charts: [] },
  });
  assert.ok(text.includes("✅"));
  assert.ok(text.includes("18,000"));
  assert.ok(!text.includes("{"));
});

test("humanizeError keeps the message readable and marks errors", () => {
  assert.equal(humanizeError("I couldn't find a column named 'Xyz'."), "❌ I couldn't find a column named 'Xyz'.");
  assert.equal(humanizeError("❌ already prefixed."), "❌ already prefixed.");
  assert.equal(humanizeError(""), "❌ Something went wrong. Try again.");
});

/* ------------------------------------------------------------------ */
/* Analysis helpers                                                    */
/* ------------------------------------------------------------------ */

const REPORT = {
  profile: { rows: 5, columns: 3, numeric_columns: ["Revenue", "Expense"] },
  statistics: [
    { column: "Month", type: "date", count: 0 },
    {
      column: "Revenue",
      type: "numeric",
      count: 5,
      sum: 7600,
      average: 1520,
      min: 1000,
      max: 2000,
    },
    {
      column: "Expense",
      type: "numeric",
      count: 5,
      sum: 3900,
      average: 780,
      min: 600,
      max: 950,
    },
  ],
  growth: [
    {
      column: "Revenue",
      available: true,
      first: 1000,
      last: 2000,
      change: 1000,
      percent: 100,
      reason: null,
    },
    {
      column: "Zero",
      available: true,
      first: 0,
      last: 5,
      change: 5,
      percent: null,
      reason: "Percentage growth is undefined because the first value is 0.",
    },
    { column: "One", available: false, reason: "Needs at least 2 numeric values." },
  ],
  trends: [{ column: "Revenue", direction: "mixed", detail: "Revenue moved 3 steps up." }],
  anomalies: [
    { column: "Revenue", count: 0, note: "No values in Revenue are outliers." },
    {
      column: "Expense",
      count: 1,
      items: [{ label: "May", value: 950, methods: ["IQR"] }],
    },
  ],
  insights: [
    { text: "Revenue has the largest total: 7600.", evidence: "sum=7600" },
    { text: "Expense has the smallest total: 3900.", evidence: "sum=3900" },
  ],
  chart_recommendation: {
    available: true,
    type: "LineChart",
    title: "Revenue vs Expense by Month",
    columns: ["Revenue", "Expense"],
  },
};

test("looksLikeAnalysisQuestion separates questions from commands", () => {
  assert.equal(looksLikeAnalysisQuestion("analyze this data"), true);
  assert.equal(looksLikeAnalysisQuestion("kya trend hai?"), true);
  assert.equal(looksLikeAnalysisQuestion("any anomalies?"), true);
  assert.equal(looksLikeAnalysisQuestion("Revenue ka total karo"), false);
  assert.equal(looksLikeAnalysisQuestion("aor b ka sum kar ke d mein daal do"), false);
  assert.equal(looksLikeAnalysisQuestion(undefined), false);
});

test("isChartRequest detects chart asks in English and Hinglish", () => {
  assert.equal(isChartRequest("create the recommended chart"), true);
  assert.equal(isChartRequest("Revenue ka line chart bana do"), true);
  assert.equal(isChartRequest("graph banao"), true);
  assert.equal(isChartRequest("what is the growth?"), false);
  assert.equal(isChartRequest(null), false);
});

/* ---------------------------------------------------------------------- */
/* Chart routing: a named column must never be rerouted onto the analysis  */
/* ---------------------------------------------------------------------- */

test("refersToRecommendedChart accepts only recommendation pointers", () => {
  assert.equal(refersToRecommendedChart("Create the recommended chart"), true);
  assert.equal(refersToRecommendedChart("Create that chart"), true);
  assert.equal(refersToRecommendedChart("Make the recommended chart"), true);
  assert.equal(refersToRecommendedChart("Recommended chart bana do"), true);
  assert.equal(refersToRecommendedChart("Create the chart you recommended"), true);
  assert.equal(refersToRecommendedChart("make that graph"), true);
});

test("refersToRecommendedChart rejects a chart that names its own column", () => {
  assert.equal(refersToRecommendedChart("Sales ka chart bana do"), false);
  assert.equal(refersToRecommendedChart("Revenue ka chart bana do"), false);
  assert.equal(refersToRecommendedChart("Expense ka chart bana do"), false);
  assert.equal(refersToRecommendedChart("Sales ka bar chart bana do"), false);
  assert.equal(refersToRecommendedChart("chart of Revenue"), false);
});

test("shouldUseStoredAnalysis keeps a named chart column on the command path", () => {
  // Defect 1 regression: these must NOT be answered from the stored report,
  // or the requested column is silently swapped for the analysed one.
  assert.equal(shouldUseStoredAnalysis("Sales ka chart bana do"), false);
  assert.equal(shouldUseStoredAnalysis("Revenue ka chart bana do"), false);
  assert.equal(shouldUseStoredAnalysis("Expense ka chart bana do"), false);
  assert.equal(shouldUseStoredAnalysis("Sales ka bar chart bana do"), false);
  assert.equal(shouldUseStoredAnalysis("Revenue ka line chart bana do"), false);
});

test("shouldUseStoredAnalysis allows the recommended-chart follow-ups", () => {
  assert.equal(shouldUseStoredAnalysis("Create the recommended chart"), true);
  assert.equal(shouldUseStoredAnalysis("Create that chart"), true);
  assert.equal(shouldUseStoredAnalysis("Recommended chart bana do"), true);
});

test("shouldUseStoredAnalysis still routes data questions to the report", () => {
  assert.equal(shouldUseStoredAnalysis("which month was highest?"), true);
  assert.equal(shouldUseStoredAnalysis("Show revenue growth"), true);
  assert.equal(shouldUseStoredAnalysis("any anomalies?"), true);
  assert.equal(shouldUseStoredAnalysis("kya trend hai?"), true);
});

// "What is important in this data?" is an analysis *trigger*: isAnalysisRequest
// claims it first, so it re-runs a fresh analysis rather than reading the
// stored report. Documented in docs/COMMANDS.md.
test("an important-style phrasing triggers a fresh analysis, not the report", () => {
  assert.equal(isAnalysisRequest("What is important in this data?"), true);
  assert.equal(shouldUseStoredAnalysis("What is important in this data?"), false);
});

// Known pre-existing Phase 4 gap, pinned here so it cannot regress silently.
// looksLikeAnalysisQuestion has no "compare" keyword, so the backend's
// compare follow-up is unreachable from the task pane and these fall through
// to the command pipeline. The backend intent itself works and is tested in
// tests/test_data_analyzer.py. Fixing the trigger is a separate change.
test("compare phrasings are not yet routed to the stored report", () => {
  assert.equal(shouldUseStoredAnalysis("Compare revenue and expenses"), false);
  assert.equal(shouldUseStoredAnalysis("Revenue vs Expense"), false);
});

test("shouldUseStoredAnalysis leaves ordinary Phase 1-3 commands alone", () => {
  assert.equal(shouldUseStoredAnalysis("b ka total karo"), false);
  assert.equal(shouldUseStoredAnalysis("inka total karo"), false);
  assert.equal(shouldUseStoredAnalysis("aor b ka sum kar ke d mein daal do"), false);
  assert.equal(shouldUseStoredAnalysis("a ka maximum batao"), false);
  assert.equal(shouldUseStoredAnalysis("c ka percentage karo"), false);
  assert.equal(shouldUseStoredAnalysis(undefined), false);
});

test("runCommand reports a command exactly once", () => {
  // Defect 2 regression: handleResponse() already emits the single success
  // message, so runCommand must not echo it again.
  const src = readFileSync(
    new URL("../src/taskpane/taskpane.js", import.meta.url),
    "utf8",
  );
  const body = src.slice(src.indexOf("async function runCommand"));
  const fn = body.slice(0, body.indexOf("\n}"));
  assert.ok(!fn.includes("addMessage"), "runCommand must not add a message itself");
  assert.ok(!fn.includes("humanizeSuccess"), "runCommand must not re-derive the message");
  assert.ok(fn.includes("handleResponse"), "runCommand must still render the response");

  // No call site may add the result of runCommand back into the chat.
  assert.equal(/addMessage\(\s*"bot"\s*,\s*run\s*\)/.test(src), false);
  assert.equal(/if\s*\(\s*run\s*\)\s*addMessage/.test(src), false);
});

test("formatMetric adds separators and keeps decimals", () => {
  assert.equal(formatMetric(7600), "7,600");
  assert.equal(formatMetric(145577390), "145,577,390");
  assert.equal(formatMetric(1234.567), "1,234.57");
  assert.equal(formatMetric(0), "0");
  assert.equal(formatMetric(null), "—");
  assert.equal(formatMetric("text"), "text");
});

test("analysisHeadline summarizes the profiled range", () => {
  assert.equal(analysisHeadline(REPORT), "5 row×3 column range · 2 numeric column(s)");
  const noNumbers = { profile: { rows: 2, columns: 1, numeric_columns: [] } };
  assert.ok(analysisHeadline(noNumbers).includes("no numeric column"));
  assert.ok(analysisHeadline({ profile: { rows: 0 } }).includes("no data"));
  assert.equal(analysisHeadline(null), "No analysis available.");
});

test("analysisTableRows lists only numeric columns", () => {
  const rows = analysisTableRows(REPORT);
  assert.equal(rows.length, 2);
  assert.deepEqual(rows[0], ["Revenue", "7,600", "1,520", "1,000", "2,000"]);
  assert.equal(rows[1][0], "Expense");
  assert.deepEqual(analysisTableRows(null), []);
});

test("analysisGrowthLines handles growth, zero denominator and unavailable", () => {
  const lines = analysisGrowthLines(REPORT);
  assert.equal(lines.length, 3);
  assert.equal(lines[0], "Revenue rose 100% (1,000 → 2,000)");
  assert.ok(lines[1].includes("changed by 5"));
  assert.ok(lines[1].includes("first value is 0"));
  assert.equal(lines[2], "One: Needs at least 2 numeric values.");
});

test("analysisGrowthLines uses period labels when the report has them", () => {
  const report = {
    growth: [
      { column: "Revenue", available: true, first: 1000, last: 2000, change: 1000, percent: 100, first_label: "Jan", last_label: "May" },
      { column: "Expense", available: true, first: 600, last: 950, change: 350, percent: 58.3 },
    ],
  };
  const lines = analysisGrowthLines(report);
  assert.equal(lines[0], "Revenue rose 100% (Jan (1,000) → May (2,000))");
  assert.equal(lines[1], "Expense rose 58.3% (600 → 950)");
});

test("analysisTrendLines and analysisAnomalyLines read the report", () => {
  assert.deepEqual(analysisTrendLines(REPORT), ["Revenue moved 3 steps up."]);
  const anomalies = analysisAnomalyLines(REPORT);
  assert.equal(anomalies.length, 2);
  assert.ok(anomalies[0].includes("No values in Revenue are outliers."));
  assert.ok(anomalies[1].includes("Expense: 950 at May (IQR)"));
});

test("analysisInsightLines accepts strings and objects", () => {
  assert.deepEqual(analysisInsightLines(REPORT), [
    "Revenue has the largest total: 7600.",
    "Expense has the smallest total: 3900.",
  ]);
  assert.deepEqual(analysisInsightLines({ insights: ["plain"] }), ["plain"]);
  assert.deepEqual(analysisInsightLines(null), []);
});

test("analysisChartPlan maps a recommendation to a pipeline command", () => {
  const plan = analysisChartPlan(REPORT);
  assert.equal(plan.available, true);
  assert.equal(plan.text, "Revenue vs Expense by Month");
  assert.equal(plan.command, "Revenue ka line chart bana do");

  const bar = analysisChartPlan({
    chart_recommendation: { available: true, type: "ColumnClustered", columns: ["Sales"] },
  });
  assert.equal(bar.command, "Sales ka bar chart bana do");

  const none = analysisChartPlan({ chart_recommendation: { available: false } });
  assert.deepEqual(none, { available: false, text: "", command: "" });
  assert.deepEqual(analysisChartPlan(null), { available: false, text: "", command: "" });
});
/* ---------------------------------------------------------------- */
/* Phase 4: analysis trigger + data-quality section                  */
/* ---------------------------------------------------------------- */

const QUALITY_REPORT = {
  profile: { rows: 5, columns: 3, numeric_columns: ["Revenue"] },
  data_quality: {
    empty_cells: 2,
    total_cells: 15,
    completeness_pct: 86.7,
    duplicate_rows: 1,
    mixed_type_columns: ["Expense"],
  },
  insights: [{ text: "Revenue rose." }],
  chart_recommendation: { available: false },
};

test("isAnalysisRequest accepts the documented analysis phrasings", () => {
  for (const cmd of [
    "analyze this data",
    "Analyze the selected data",
    "Give me insights",
    "Is data ka analysis karo",
    "Is data mein kya important hai?",
    "give me a report",
    "summary dedo",
  ]) {
    assert.equal(isAnalysisRequest(cmd), true, cmd);
  }
});

test("isAnalysisRequest ignores ordinary write commands", () => {
  for (const cmd of [
    "b ka total karo",
    "inka average nikalo",
    "Revenue ka chart bana do",
    "aor b ka sum kar ke d mein daal do",
  ]) {
    assert.equal(isAnalysisRequest(cmd), false, cmd);
  }
});

test("isAnalysisRequest handles non-strings safely", () => {
  assert.equal(isAnalysisRequest(null), false);
  assert.equal(isAnalysisRequest(undefined), false);
  assert.equal(isAnalysisRequest(42), false);
});

test("looksLikeAnalysisQuestion still accepts plural insights", () => {
  assert.equal(looksLikeAnalysisQuestion("Give me insights"), true);
  assert.equal(looksLikeAnalysisQuestion("Which month was highest?"), true);
  assert.equal(looksLikeAnalysisQuestion("b ka total karo"), false);
});

test("analysisQualityLines reports empty cells, duplicates and completeness", () => {
  const lines = analysisQualityLines(QUALITY_REPORT);

  assert.equal(lines.length, 4);
  assert.ok(lines[0].includes("2 empty cells"));
  assert.ok(lines[0].includes("15"));
  assert.ok(lines[1].includes("1 duplicate row"));
  assert.ok(lines[2].includes("86.7%"));
  assert.ok(lines[3].includes("Expense"));
});

test("analysisQualityLines reports a clean table plainly", () => {
  const lines = analysisQualityLines({
    data_quality: {
      empty_cells: 0,
      total_cells: 15,
      completeness_pct: 100,
      duplicate_rows: 0,
      mixed_type_columns: [],
    },
  });

  assert.deepEqual(lines, [
    "No empty cells found.",
    "No duplicate rows.",
    "100% of cells are filled.",
  ]);
});

test("analysisQualityLines singularises a single empty cell and duplicate", () => {
  const lines = analysisQualityLines({
    data_quality: {
      empty_cells: 1,
      total_cells: 4,
      completeness_pct: 75,
      duplicate_rows: 1,
      mixed_type_columns: [],
    },
  });

  assert.equal(lines[0], "1 empty cell out of 4.");
  assert.equal(lines[1], "1 duplicate row.");
});

test("analysisQualityLines returns nothing without a quality block", () => {
  assert.deepEqual(analysisQualityLines({}), []);
  assert.deepEqual(analysisQualityLines(null), []);
});
