"""Deterministic follow-up answers for an existing analysis.

After ``POST /api/excel/analyze`` the task pane keeps the last report in
memory. The user then asks things like "which month was highest?", "what is
the growth?", "are there anomalies?" or "create the recommended chart".

Every answer is computed here from the *already returned* report, so a
follow-up costs no new analysis, needs no model call, and always agrees with
the numbers the user just saw.
"""

import re

CHART_COMMAND_RE = re.compile(r"\b(chart|graph|banao|bana|plot|visuali[sz])\b", re.IGNORECASE)
HIGHEST_RE = re.compile(r"\b(highest|maximum|max|top|best|sabse\s*(zyada|adhik|peak))\b", re.I)
LOWEST_RE = re.compile(r"\b(lowest|minimum|min|bottom|worst|sabse\s*(kam|kam\s*kam))\b", re.I)
GROWTH_RE = re.compile(r"\b(growth|grew|increased?|badh\w*|badhi|rise|up)\b", re.I)
ANOMALY_RE = re.compile(r"\b(anomal\w*|outlier\w*|odd|unusual|weird|gaitak)\b", re.I)
TREND_RE = re.compile(r"\b(trend|pattern|badhwa|dhyan|upar|neeche)\b", re.I)
QUALITY_RE = re.compile(r"\b(quality|missing|empty|blank|duplicate|null)\b", re.I)
#: "insights", "summary", "overview" and the phrasing the task pane also treats
#: as a fresh-analysis request: "what is important in this data?" /
#: "is data mein kya important hai?". Both must answer from the existing
#: report instead of falling through to "I can't answer that".
SUMMARY_RE = re.compile(r"\b(insights?|summary|summar\w*|overview|important|samavesh)\b", re.I)
TOTAL_RE = re.compile(r"\b(total|sum)\b", re.I)
#: "Compare revenue and expenses", "revenue vs expense", "kya antar hai". A
#: comparison names two columns explicitly, so it is a different question from
#: "which is highest?" and is answered from the per-column statistics.
COMPARE_RE = re.compile(
    r"\b(compare|comparison|versus|vs\.?|difference\s+between|"
    r"antar|antardhan|antra)\b",
    re.I,
)
#: Words that make "highest"/"lowest" a question about a *row* ("which month
#: was highest?") rather than about a column ("which column is biggest?").
PERIOD_RE = re.compile(
    r"\b(month|months|period|periods|row|rows|quarter|week|day|date|dates|"
    r"year|years|time|when|kab|kaun\s*sa|kaun\s*si)\b",
    re.I,
)
#: ...and the opposite: an explicit "which column" question always compares
#: columns, even when the table has a label column.
COLUMN_RE = re.compile(r"\b(column|columns|field|fields|series|metric|metrics)\b", re.I)


def classify_followup(text, column_names=None):
    """Map a follow-up question to an intent.

    Returns ``(intent, column_name)``; ``intent`` is ``"chart"``,
    ``"highest"``, ``"lowest"``, ``"growth"``, ``"anomalies"``, ``"trend"``,
    ``"quality"``, ``"summary"`` or ``"unknown"``. ``column_names`` are the
    columns of the report being questioned, so arbitrary user headers are
    recognised without a hard-coded vocabulary.
    """
    question = str(text or "").strip()
    if not question:
        return ("unknown", None)

    named = _column_in(question, column_names)
    if CHART_COMMAND_RE.search(question):
        return ("chart", named)
    if COMPARE_RE.search(question):
        return ("compare", named)
    if ANOMALY_RE.search(question):
        return ("anomalies", named)
    if QUALITY_RE.search(question):
        return ("quality", named)
    if TREND_RE.search(question):
        return ("trend", named)
    if GROWTH_RE.search(question):
        return ("growth", named)
    if HIGHEST_RE.search(question):
        return ("highest", named)
    if LOWEST_RE.search(question):
        return ("lowest", named)
    if SUMMARY_RE.search(question) or TOTAL_RE.search(question):
        return ("summary", named)
    return ("unknown", named)


def answer_followup(question, analysis):
    """Answer a follow-up from a previous analysis report.

    Returns ``{"intent", "message", "chart_command", "handled"}``. Unknown or
    unanswerable questions are reported as such so the task pane can fall back
    to the regular command pipeline instead of inventing an answer.
    """
    if not analysis:
        return {
            "intent": "unknown",
            "message": "Run an analysis first, then I can answer questions about it.",
            "chart_command": None,
            "handled": False,
        }

    column_names = [s.get("column") for s in analysis.get("statistics", []) if s.get("column")]
    intent, requested = classify_followup(question, column_names)
    numeric_names = [
        s["column"]
        for s in analysis.get("statistics", [])
        if s.get("type") == "numeric" and s.get("count")
    ]
    if requested in numeric_names:
        target = requested
    elif len(numeric_names) == 1:
        target = numeric_names[0]
    else:
        target = None

    handler = {
        "chart": _answer_chart,
        "compare": _answer_compare,
        "highest": _answer_extreme,
        "lowest": _answer_extreme,
        "growth": _answer_growth,
        "anomalies": _answer_anomalies,
        "trend": _answer_trend,
        "quality": _answer_quality,
        "summary": _answer_summary,
    }.get(intent)

    if handler is None:
        return {
            "intent": "unknown",
            "message": (
                "I can't answer that from the analysis. Try 'highest month', "
                "'growth', 'anomalies', or 'create the recommended chart'."
            ),
            "chart_command": None,
            "handled": False,
        }

    question_text = str(question or "")
    if intent == "compare":
        # A comparison needs every column the user named, not just the first,
        # so it is resolved here rather than through the single-column target.
        return {
            "intent": intent,
            "message": _answer_compare(analysis, _columns_in(question, column_names)),
            "chart_command": None,
            "handled": True,
        }

    wants_column = bool(COLUMN_RE.search(question_text))
    wants_period = bool(PERIOD_RE.search(question_text)) and not wants_column
    if intent in ("highest", "lowest"):
        # A table with a label axis makes "highest" mean the peak period unless
        # the question explicitly asked about columns.
        wants_period = wants_period or (not wants_column and _period_wanted_by_default(analysis))
    message, chart_command = handler(analysis, target, intent, wants_period)
    return {
        "intent": intent,
        "message": message,
        "chart_command": chart_command,
        "handled": True,
    }


# ---------------------------------------------------------------------- #
# Intent handlers
# ---------------------------------------------------------------------- #


def _answer_chart(analysis, _target, _intent, _wants_period=False):
    chart = analysis.get("chart_recommendation") or {}
    if not chart.get("available"):
        return (chart.get("reason", "No chart is recommended for this data."), None)

    series = chart.get("columns") or ["data"]
    primary = series[0]
    line = chart.get("type") == "LineChart"
    # The existing chart pipeline reads "line" in the command to pick the
    # chart type, and charts one column per request, so the recommendation is
    # turned into exactly the command that pipeline understands today. The
    # wording says which column is charted so nothing is implied about the rest.
    command = f"{primary} ka line chart bana do" if line else f"{primary} ka chart bana do"
    others = [name for name in series[1:]]
    extra = f" {primary} is charted first; add {', '.join(others)} separately." if others else ""
    return (
        f"Recommended: {primary}"
        + (f" by {chart['label_column']}" if chart.get("label_column") else "")
        + f" as a {'line' if line else 'bar'} chart.{extra}",
        command,
    )


def _answer_compare(analysis, names):
    """Answer "compare revenue and expenses" from the per-column statistics.

    The user named at least one column, so those are compared against each
    other. With no usable pair the honest answer is the full ranking rather
    than a guess about which two were meant.
    """
    stats = {s["column"]: s for s in analysis.get("statistics", [])}
    numeric = [stats[name] for name in names if name in stats and stats[name].get("count")]

    if len(numeric) < 2:
        ranking = [
            s
            for s in analysis.get("statistics", [])
            if s.get("type") == "numeric" and s.get("count")
        ]
        if not ranking:
            return "There is no numeric column to compare yet."
        ordered = sorted(ranking, key=lambda s: s["sum"], reverse=True)
        listed = ", ".join(f"{s['column']} ({s['sum']})" for s in ordered)
        return (
            "I could not tell which columns to compare, so here is every numeric "
            f"column by total: {listed}. Name the two you want to compare."
        )

    first, second = numeric[0], numeric[1]
    gap = first["sum"] - second["sum"]
    lead = "more" if gap > 0 else "less"
    if gap == 0:
        lead = ""
    ratio = _ratio(first["sum"], second["sum"])
    line = (
        f"{first['column']} totals {first['sum']} against {second['column']} at "
        f"{second['sum']}, so {first['column']} is {_signed(gap)} {lead}".strip()
    )
    if ratio:
        line += f" ({ratio}x)."
    else:
        line += "."
    detail = (
        f" Averages: {first['column']} {first['average']} vs "
        f"{second['column']} {second['average']}."
    )
    return line + detail


def _signed(value):
    return f"{value:+g}"


def _ratio(larger, smaller):
    """Larger-over-smaller ratio, or None when it is not meaningful."""
    if not smaller or larger is None:
        return None
    return _fmt_ratio(larger / smaller)


def _fmt_ratio(value):
    rounded = round(value, 1)
    return int(rounded) if float(rounded).is_integer() else rounded


def _answer_extreme(analysis, target, intent, wants_period=False):
    """Answer "which month was highest?" (a row) or "which column is highest?"."""
    if wants_period:
        return (_answer_period_extreme(analysis, target, intent), None)
    return _answer_column_extreme(analysis, target, intent)


def _period_wanted_by_default(analysis):
    """True when the table has a label column, so "highest" means a period.

    With a Month/Date/Name axis present, "which is highest?" is far more often
    about the peak row than about which column has the largest total, and the
    period answer is the more useful one. A column-vs-column answer is still
    available by naming a column explicitly.
    """
    for entry in analysis.get("period_extremes") or []:
        if entry.get("highest", {}).get("label"):
            return True
    return False


def _answer_period_extreme(analysis, target, intent):
    extremes = analysis.get("period_extremes") or []
    if not extremes:
        return "There is no numeric period to compare yet."

    key = "highest" if intent == "highest" else "lowest"
    word = key
    label = "peak" if intent == "highest" else "trough"

    if target:
        entries = [e for e in extremes if e["column"] == target]
        if not entries:
            return f"{target} has no numeric values to compare."
    else:
        entries = extremes

    pairs = [(e, e[key]) for e in entries if e.get(key) and e[key].get("value") is not None]
    if not pairs:
        return "There is no numeric period to compare yet."

    chooser = max if intent == "highest" else min
    entry, point = chooser(pairs, key=lambda pair: pair[1]["value"])

    if target:
        return (
            f"{point['where']} was the {word} {target} period with {point['value']} "
            f"(average {_stat_value(analysis, target, 'average')}, "
            f"{_stat_value(analysis, target, 'count')} value(s))."
        )

    also = ", ".join(
        f"{other['column']} {other[key]['value']}" for other, _ in pairs if other is not entry
    )
    return (
        f"{point['where']} was the {word} period: {entry['column']} reached {point['value']} "
        f"({label} of the {len(pairs)} numeric column(s))"
        + (f", compared with {also}" if also else "")
        + "."
    )


def _answer_column_extreme(analysis, target, intent):
    stats = {s["column"]: s for s in analysis.get("statistics", [])}
    candidates = [stats[target]] if target and target in stats else list(stats.values())
    numeric = [s for s in candidates if s.get("type") == "numeric" and s.get("count")]
    if not numeric:
        # Same ``(message, chart_command)`` shape as the success path below,
        # so a text-only selection degrades to an honest answer instead of
        # raising when the dispatcher unpacks the result.
        return "There is no numeric column to compare yet.", None

    key = "max" if intent == "highest" else "min"
    pick = max(numeric, key=lambda s: s[key])
    others = [s for s in numeric if s is not pick]
    word = "highest" if intent == "highest" else "lowest"
    comparison = ""
    if others:
        listed = ", ".join(f"{s['column']} ({s[key]})" for s in others)
        comparison = f", compared with {listed}"
    period = _peak_phrase(analysis, pick["column"], key)
    return (
        f"{pick['column']} is the {word} with {pick[key]} "
        f"(average {pick['average']}, {pick['count']} value(s)){comparison}{period}.",
        None,
    )


def _peak_phrase(analysis, column, key):
    """Name the peak/trough period inside a column answer, when known."""
    for entry in analysis.get("period_extremes") or []:
        if entry.get("column") != column:
            continue
        point = entry.get("highest" if key == "max" else "lowest") or {}
        if point.get("where") and point.get("value") is not None:
            word = "peak" if key == "max" else "trough"
            return f"; {word} period {point['where']} at {point['value']}"
    return ""


def _stat_value(analysis, column, key):
    for stat in analysis.get("statistics", []):
        if stat.get("column") == column:
            return stat.get(key)
    return None


def _answer_growth(analysis, target, _intent, _wants_period=False):
    rows = {g["column"]: g for g in analysis.get("growth", [])}
    candidates = [rows[target]] if target and target in rows else list(rows.values())
    if not candidates:
        return ("I need at least two numeric values to measure growth.", None)

    lines = []
    for row in candidates:
        if not row.get("available"):
            lines.append(f"{row['column']}: {row.get('reason', 'not enough values')}")
            continue
        if row.get("percent") is None:
            lines.append(f"{row['column']}: changed by {row['change']} ({row.get('reason')})")
            continue
        word = "rose" if row["change"] >= 0 else "fell"
        span = _growth_span(row)
        lines.append(f"{row['column']} {word} {abs(row['percent'])}% {span}.")
    return (" ".join(lines), None)


def _growth_span(row):
    first_label, last_label = row.get("first_label"), row.get("last_label")
    if first_label and last_label:
        return f"from {first_label} ({row['first']}) to {last_label} ({row['last']})"
    if last_label:
        return f"from the first value ({row['first']}) to {last_label} ({row['last']})"
    if first_label:
        return f"from {first_label} ({row['first']}) to the last value ({row['last']})"
    return f"from {row['first']} to {row['last']}"


def _answer_anomalies(analysis, _target, _intent, _wants_period=False):
    findings = analysis.get("anomalies", [])
    flagged = [f for f in findings if f.get("count")]
    if not flagged:
        checked = [f["column"] for f in findings if f.get("method") == "IQR + z-score"]
        if checked:
            return (
                f"No values in {' or '.join(checked)} fall outside 1.5x IQR or "
                "2 standard deviations.",
                None,
            )
        notes = "; ".join(f["note"] for f in findings if f.get("note"))
        return (notes or "No anomalies were found in this data.", None)

    parts = []
    for finding in flagged:
        for item in finding["items"][:3]:
            where = (
                f" at {item['label']}" if item.get("label") else f" in row {item['row_index'] + 1}"
            )
            parts.append(
                f"{finding['column']}: {item['value']}{where} ({', '.join(item['methods'])})"
            )
    return ("Outliers detected - " + "; ".join(parts) + ".", None)


def _answer_trend(analysis, target, _intent, _wants_period=False):
    trends = analysis.get("trends", [])
    if target:
        trends = [t for t in trends if t["column"] == target]
    if not trends:
        return ("There is not enough data to detect a trend.", None)
    return (" ".join(t["detail"] for t in trends), None)


def _answer_quality(analysis, _target, _intent, _wants_period=False):
    quality = analysis.get("data_quality", {})
    profile = analysis.get("profile", {})
    return (
        f"{quality.get('completeness_pct')}% of cells are filled "
        f"({quality.get('empty_cells', 0)} empty of {quality.get('total_cells', 0)}), "
        f"{quality.get('duplicate_rows', 0)} duplicate row(s), "
        f"mixed-type columns: "
        f"{', '.join(quality.get('mixed_type_columns') or ['none'])}. "
        f"Columns profiled: {', '.join(profile.get('column_names') or ['none'])}.",
        None,
    )


def _answer_summary(analysis, target, _intent, _wants_period=False):
    insights = analysis.get("insights", [])
    if target:
        matching = [i for i in insights if target in i["text"]]
        if matching:
            insights = matching
    if not insights:
        return ("The analysis did not produce any insights.", None)
    return (" ".join(i["text"] for i in insights[:3]), None)


# ---------------------------------------------------------------------- #
# Helpers
# ---------------------------------------------------------------------- #


def _column_in(question, column_names=None):
    """Best-effort column name mentioned in the question.

    The report's own column names win, so any header the user has works, not
    just a fixed vocabulary. Longest name first stops "Net Revenue" from being
    read as "Revenue" when both exist. Word-ish boundaries stop "Cost" from
    matching inside "Customer".
    """
    matches = _columns_in(question, column_names)
    return matches[0] if matches else None


def _columns_in(question, column_names=None):
    """Every known column name mentioned in the question, longest first.

    Used by comparison questions, which name two columns at once. A trailing
    "s" is allowed so "customers" finds the "Customer" column, while the
    word boundaries keep "Cost" out of "Customer".
    """
    text = str(question or "").lower()
    candidates = [str(name) for name in (column_names or []) if name and str(name).strip()]
    candidates += [
        name for name in _KNOWN_COLUMNS if name.lower() not in {c.lower() for c in candidates}
    ]
    found = []
    for name in sorted(candidates, key=len, reverse=True):
        pattern = r"(?<!\w)" + re.escape(name.lower()) + r"s?(?!\w)"
        if re.search(pattern, text) and name not in found:
            found.append(name)
    return found


_KNOWN_COLUMNS = (
    "Revenue",
    "Expense",
    "Profit",
    "Sales",
    "Cost",
    "Income",
    "Quantity",
    "Price",
    "Amount",
    "Total",
    "Loss",
)
