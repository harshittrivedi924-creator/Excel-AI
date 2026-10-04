"""Read-only data profiling and analysis for the Excel Add-in.

This module is the deterministic core behind ``POST /api/excel/analyze``.
It answers "what is in this data?" without touching the user's workbook,
the in-memory command workbook, or any file on disk: every input arrives
as a plain snapshot (headers + rows, optionally narrowed to the user's
selection) and every output is JSON-friendly data.

Design constraints that shape the code:

* **Read-only.** Nothing here mutates a worksheet. Charts are *recommended*,
  never created; the Add-in turns a recommendation into a real chart through
  the existing ``/api/excel/command`` -> chart -> Office.js pipeline.
* **Deterministic.** The same snapshot always yields the same numbers and the
  same wording. There is no model call, no sampling and no randomness, so
  results are explainable and testable.
* **Honest.** Statistics that need a minimum sample are reported as
  unavailable instead of guessed, and every insight carries the figures it was
  derived from.
* **Dependency-light.** Only the standard library plus the project's existing
  ``CalculationEngine`` and ``is_blank`` helpers are used.
"""

import math
import re
from datetime import date, datetime

from backend.calculator.engine import CalculationEngine
from backend.excel.memory_handler import is_blank

#: A column needs at least this many numeric values before growth/trend
#: language is used. Below it the analysis says so instead of extrapolating
#: from one or two points.
MIN_TREND_POINTS = 3

#: Minimum numeric values before outlier detection is attempted. IQR and
#: z-score are meaningless on tiny samples and produce noisy results.
MIN_ANOMALY_POINTS = 4

#: Names that identify a row-ordering / time axis, in priority order. Used
#: only to decide trend direction and chart type, never to invent values.
_TIME_LABELS = (
    "month",
    "date",
    "day",
    "week",
    "quarter",
    "year",
    "period",
    "weekday",
    "time",
    "name",
    "product",
    "category",
    "item",
)

_MONTHS = {
    "jan": 1,
    "feb": 2,
    "mar": 3,
    "apr": 4,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "aug": 8,
    "sep": 9,
    "oct": 10,
    "nov": 11,
    "dec": 12,
    "january": 1,
    "february": 2,
    "march": 3,
    "april": 4,
    "june": 6,
    "july": 7,
    "august": 8,
    "september": 9,
    "october": 10,
    "november": 11,
    "december": 12,
}

#: ISO-ish and common textual date shapes, tried in order.
_DATE_PATTERNS = (
    (r"^\d{4}-\d{1,2}-\d{1,2}$", (1, 2, 3)),
    (r"^\d{1,2}[/-]\d{1,2}[/-]\d{4}$", (3, 2, 1)),
    (r"^\d{1,2}[/-]\d{1,2}[/-]\d{2}$", (3, 2, 1)),
    (r"^\d{1,2}-[A-Za-z]{3,}-\d{4}$", None),
    (r"^[A-Za-z]{3,}\s+\d{1,2},?\s*\d{4}$", None),
)

_NUMERIC_CLEAN_RE = re.compile(r"[,%\s]")
_MONEY_PREFIXES = ("$", "rs", "₹", "€", "£")


class DataAnalyzer:
    """Profile and analyse a worksheet snapshot."""

    # ------------------------------------------------------------------ #
    # Public API
    # ------------------------------------------------------------------ #

    @classmethod
    def analyze(cls, headers=None, rows=None, selection=None, sheet_name=None):
        """Analyse a snapshot and return a fully JSON-friendly report."""
        table = _resolve_table(headers, rows, selection)
        headers_out, data_rows, scope = table

        columns = _build_columns(headers_out, data_rows)
        stats = [_column_stats(col) for col in columns]
        numeric_columns = [c for c in columns if c["type"] == "numeric"]
        label_column = _label_column(columns)

        growth = [_column_growth(c, data_rows, label_column) for c in numeric_columns]
        trends = _detect_trends(numeric_columns, label_column, data_rows)
        anomalies = _detect_anomalies(numeric_columns, label_column, data_rows)
        extremes = _period_extremes(numeric_columns, data_rows, label_column)
        chart = _recommend_chart(columns, stats, label_column, data_rows)
        quality = _data_quality(columns, data_rows)
        insights = _build_insights(
            stats, growth, trends, anomalies, extremes, chart, quality, scope
        )

        return {
            "success": True,
            "sheet": sheet_name,
            "scope": scope,
            "profile": {
                "rows": len(data_rows),
                "columns": len(columns),
                "column_names": [c["name"] for c in columns],
                "numeric_columns": [c["name"] for c in numeric_columns],
                "text_columns": [c["name"] for c in columns if c["type"] in ("text", "date")],
                "date_columns": [c["name"] for c in columns if c["type"] == "date"],
                "empty_columns": [c["name"] for c in columns if c["type"] == "empty"],
                "empty_cells": quality["empty_cells"],
                "total_cells": quality["total_cells"],
                "completeness_pct": quality["completeness_pct"],
                "duplicate_rows": quality["duplicate_rows"],
                "unique_values": {c["name"]: c["unique_values"] for c in columns},
            },
            "statistics": stats,
            "growth": growth,
            "trends": trends,
            "anomalies": anomalies,
            "period_extremes": extremes,
            "insights": insights,
            "chart_recommendation": chart,
            "data_quality": quality,
            "limitations": _limitations(stats, data_rows, scope),
        }


# ---------------------------------------------------------------------- #
# Snapshot / selection handling
# ---------------------------------------------------------------------- #


def _resolve_table(headers, rows, selection):
    """Pick the cells to analyse: the selection first, else the whole snapshot.

    Returns ``(headers, data_rows, scope)``. The selection wins when it
    actually carries data, which is what the user sees highlighted; an empty
    selection is not an error and simply falls back to the used range.
    """
    if selection and isinstance(selection, dict):
        sel_rows = selection.get("values")
        if (
            isinstance(sel_rows, list)
            and sel_rows
            and any(not is_blank(cell) for row in sel_rows for cell in _row_of(row))
        ):
            sel_headers, sel_data = _split_header(sel_rows)
            return (
                sel_headers,
                sel_data,
                {
                    "source": "selection",
                    "address": selection.get("address"),
                    "row_count": len(sel_data),
                    "column_count": len(sel_headers),
                },
            )

    data_rows = [list(r) for r in (rows or []) if isinstance(r, list)]
    return (
        list(headers or []),
        data_rows,
        {
            "source": "worksheet",
            "address": None,
            "row_count": len(data_rows),
            "column_count": len(headers or []),
        },
    )


def _row_of(row):
    return row if isinstance(row, list) else []


def _split_header(values):
    """Use the first row as headers when it is not made of plain numbers."""
    first = _row_of(values[0])
    if first and all(is_blank(cell) or _to_number(cell) is None for cell in first):
        headers = [_clean_header(cell, i) for i, cell in enumerate(first)]
        return headers, [list(r) for r in values[1:]]
    return [_clean_header(None, i) for i in range(len(first))], [list(r) for r in values]


def _clean_header(value, index):
    if not is_blank(value):
        text = str(value).strip()
        if text:
            return text
    return f"Column {index + 1}"


# ---------------------------------------------------------------------- #
# Column classification
# ---------------------------------------------------------------------- #


def _build_columns(headers, rows):
    width = max([len(headers)] + [len(r) for r in rows]) if rows else len(headers)
    columns = []
    for index in range(width):
        # A header cell that is present but blank (a user typed a formula into
        # a column they never headed) must fall back to the positional name
        # too, otherwise the column reaches the task pane named ``None``.
        name = _clean_header(headers[index] if index < len(headers) else None, index)
        values = [_cell(row, index) for row in rows]
        populated = [v for v in values if not is_blank(v)]
        # Row indices are kept alongside the numbers so every later reference
        # (trend step, outlier, period extreme) points at the real worksheet
        # row even when the column has gaps or mixed values.
        number_rows = [i for i, v in enumerate(values) if _to_number(v) is not None]
        numbers = [_to_number(values[i]) for i in number_rows]
        dates = [_to_date(v) for v in populated]
        date_hits = [d for d in dates if d is not None]
        text_hits = [v for v in populated if _to_number(v) is None and _to_date(v) is None]

        if not populated:
            kind = "empty"
        elif len(numbers) == len(populated):
            kind = "numeric"
        elif len(date_hits) == len(populated):
            kind = "date"
        elif len(text_hits) == len(populated):
            kind = "text"
        else:
            kind = "mixed"

        columns.append(
            {
                "index": index,
                "name": name,
                "type": kind,
                "values": values,
                "populated": populated,
                "numbers": numbers,
                "number_rows": number_rows,
                "unique_values": len({_key(v) for v in populated}),
            }
        )
    return columns


def _cell(row, index):
    return row[index] if index < len(row) else None


def _to_number(value):
    """Coerce a snapshot cell to int/float, or None when not numeric."""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return None if isinstance(value, float) and math.isnan(value) else value
    if not isinstance(value, str):
        return None
    text = value.strip().lower().replace("\u00a0", " ")
    if not text or text in ("-", "--", "n/a", "na", "null", "none"):
        return None
    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1].strip()
    for prefix in _MONEY_PREFIXES:
        if text.startswith(prefix):
            text = text[len(prefix) :].strip()
            break
    percent = text.endswith("%")
    if percent:
        text = text[:-1].strip()
    text = _NUMERIC_CLEAN_RE.sub("", text)
    if not text:
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    if percent:
        number = number / 100
    if negative:
        number = -number
    return int(number) if number.is_integer() and abs(number) < 2**53 else number


def _to_date(value):
    """Parse a date-like cell into ``datetime``/``date``, else None."""
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    if not isinstance(value, str):
        return None
    text = value.strip()
    if not text:
        return None

    for pattern, order in _DATE_PATTERNS:
        if order is None:
            continue
        if not re.match(pattern, text):
            continue
        parts = [int(p) for p in re.split(r"[/-]", text)]
        try:
            year, month, day = (parts[i] for i in (order[0] - 1, order[1] - 1, order[2] - 1))
            return datetime(year, month, day)
        except ValueError:
            return None

    lowered = text.lower()
    for pattern in (r"^\d{1,2}-([a-z]{3,9})-(\d{4})$", r"^([a-z]{3,9})\s+(\d{1,2}),?\s*(\d{4})$"):
        match = re.match(pattern, lowered)
        if not match:
            continue
        for group in match.groups():
            if group in _MONTHS:
                month = _MONTHS[group]
                numbers = [int(g) for g in match.groups() if g.isdigit()]
                day = min(numbers)
                year = max(numbers)
                try:
                    return datetime(year, month, day)
                except ValueError:
                    return None

    if lowered in _MONTHS and len(lowered) <= 12:
        return _MONTHS[lowered]
    return None


def _is_time_like(column):
    name = str(column["name"]).strip().lower()
    return any(label in name for label in _TIME_LABELS)


def _label_column(columns):
    """First non-numeric column, preferring time-like names, else None."""
    for column in columns:
        if column["type"] in ("text", "date") and _is_time_like(column):
            return column
    for column in columns:
        if column["type"] in ("text", "date"):
            return column
    return None


# ---------------------------------------------------------------------- #
# Statistics
# ---------------------------------------------------------------------- #


def _column_stats(column):
    numbers = column["numbers"]
    stats = {
        "column": column["name"],
        "type": column["type"],
        "count": len(numbers),
        "empty": sum(1 for v in column["values"] if is_blank(v)),
        "unique_values": column["unique_values"],
    }
    if not numbers:
        stats.update({"sum": None, "average": None, "min": None, "max": None})
        return stats

    total = CalculationEngine.sum(numbers)
    stats.update(
        {
            "sum": _round(total),
            "average": _round(CalculationEngine.average(numbers)),
            "min": _round(CalculationEngine.minimum(numbers)),
            "max": _round(CalculationEngine.maximum(numbers)),
            "median": _round(CalculationEngine.median(numbers)),
            "range": _round(
                CalculationEngine.maximum(numbers) - CalculationEngine.minimum(numbers)
            ),
            "stdev": _round(CalculationEngine.standard_deviation(numbers)),
        }
    )
    return stats


def _column_growth(column, rows, label_column):
    """First-to-last change for a numeric column, with safe percentages."""
    numbers = column["numbers"]
    if len(numbers) < 2:
        return {
            "column": column["name"],
            "available": False,
            "reason": "Needs at least 2 numeric values.",
        }
    first, last = numbers[0], numbers[-1]
    number_rows = column["number_rows"]
    change = last - first
    if first == 0:
        percent = None
        reason = "Percentage growth is undefined because the first value is 0."
    else:
        # One decimal is enough to read and keeps the payload small; the raw
        # ratio is still derivable from first/last/change.
        percent = _round(CalculationEngine.growth(last, first), 1)
        reason = None
    return {
        "column": column["name"],
        "available": True,
        "first": _round(first),
        "last": _round(last),
        "first_label": _row_label(rows, label_column, number_rows[0]),
        "last_label": _row_label(rows, label_column, number_rows[-1]),
        "change": _round(change),
        "percent": percent,
        "reason": reason,
    }


def _detect_trends(numeric_columns, label_column, rows):
    """Describe direction per numeric column using consecutive deltas."""
    trends = []
    for column in numeric_columns:
        numbers = column["numbers"]
        name = column["name"]
        if len(numbers) < MIN_TREND_POINTS:
            trends.append(
                {
                    "column": name,
                    "direction": "insufficient_data",
                    "detail": (
                        f"{name} has {len(numbers)} numeric value(s); "
                        f"at least {MIN_TREND_POINTS} are needed to judge a trend."
                    ),
                    "rises": 0,
                    "falls": 0,
                    "flat": 0,
                    "net_change": None,
                }
            )
            continue

        deltas = [b - a for a, b in zip(numbers, numbers[1:], strict=False)]
        rises = sum(1 for d in deltas if d > 0)
        falls = sum(1 for d in deltas if d < 0)
        flat = len(deltas) - rises - falls
        net = numbers[-1] - numbers[0]
        if rises and not falls:
            direction = "increasing"
        elif falls and not rises:
            direction = "decreasing"
        elif rises or falls:
            # A wobbly series still has a direction when the steps agree with
            # the net move. Reporting "mixed" for a series that only ever
            # dipped once reads as non-committal and hides real structure, so
            # the step counts decide and ``net_change`` keeps the wobble
            # auditable. The detail text below still names every rise and fall,
            # so a weak trend is never presented as a clean one.
            net_up = net > 0
            if rises > falls and net_up:
                direction = "increasing"
            elif falls > rises and not net_up:
                direction = "decreasing"
            else:
                direction = "mixed"
        else:
            direction = "flat"

        biggest = None
        for i, delta in enumerate(deltas):
            if biggest is None or abs(delta) > abs(biggest[1]):
                biggest = (i + 1, delta)

        detail = (
            f"{name} moved {rises} step(s) up and {falls} step(s) down "
            f"across {len(numbers)} values (net {net:+g})."
        )
        if biggest and biggest[1]:
            number_rows = column["number_rows"]
            row_index = number_rows[biggest[0]] if biggest[0] < len(number_rows) else None
            label = _row_label(rows, label_column, row_index)
            detail += f" Largest change {biggest[1]:+g}" + (f" at {label}" if label else "") + "."
        trends.append(
            {
                "column": name,
                "direction": direction,
                "detail": detail,
                "rises": rises,
                "falls": falls,
                "flat": flat,
                "net_change": _round(net),
            }
        )
    return trends


def _detect_anomalies(numeric_columns, label_column, rows):
    """Flag outliers with IQR and z-score; both are deterministic."""
    findings = []
    for column in numeric_columns:
        numbers = column["numbers"]
        name = column["name"]
        if len(numbers) < MIN_ANOMALY_POINTS:
            findings.append(
                {
                    "column": name,
                    "count": 0,
                    "method": "insufficient_data",
                    "note": (
                        f"{name} has {len(numbers)} numeric value(s); "
                        f"at least {MIN_ANOMALY_POINTS} are needed for outlier detection."
                    ),
                }
            )
            continue

        flagged = {}
        number_rows = column["number_rows"]
        iqr = _iqr_bounds(numbers)
        if iqr is not None:
            low, high = iqr
            for index, value in enumerate(numbers):
                if value < low or value > high:
                    flagged.setdefault(index, []).append("IQR")
        stdev = CalculationEngine.standard_deviation(numbers)
        if stdev > 0:
            mean = CalculationEngine.average(numbers)
            for index, value in enumerate(numbers):
                if abs(value - mean) > 2 * stdev:
                    flagged.setdefault(index, []).append(f"z-score (|z| > 2, stdev {stdev:.2f})")

        items = []
        for index in sorted(flagged):
            row_index = number_rows[index] if index < len(number_rows) else index
            items.append(
                {
                    "row_index": row_index,
                    "label": _row_label(rows, label_column, row_index),
                    "value": _round(numbers[index]),
                    "methods": flagged[index],
                }
            )
        findings.append(
            {
                "column": name,
                "count": len(items),
                "method": "IQR + z-score",
                "items": items,
                "note": None
                if items
                else f"No values in {name} fall outside 1.5x IQR or 2 standard deviations.",
            }
        )
    return findings


def _iqr_bounds(numbers):
    if len(numbers) < 4:
        return None
    ordered = sorted(numbers)
    median = CalculationEngine.median(ordered)
    lower_half = [v for v in ordered if v <= median]
    upper_half = [v for v in ordered if v >= median]
    if len(lower_half) < 2 or len(upper_half) < 2:
        return None
    q1 = CalculationEngine.median(lower_half)
    q3 = CalculationEngine.median(upper_half)
    iqr = q3 - q1
    if iqr <= 0:
        return None
    return (q1 - 1.5 * iqr, q3 + 1.5 * iqr)


# ---------------------------------------------------------------------- #
# Chart recommendation (never creates a chart)
# ---------------------------------------------------------------------- #


def _recommend_chart(columns, stats, label_column, rows):
    """Suggest one visualization that the existing chart pipeline can build."""
    numeric = [c for c in columns if c["type"] == "numeric" and c["numbers"]]
    if not numeric:
        return {
            "available": False,
            "reason": "No numeric column is available to chart.",
        }

    series = [c["name"] for c in numeric]
    line = label_column is not None and len(numeric) <= 2 and len(rows) >= MIN_TREND_POINTS
    chart_type = "LineChart" if line else "ColumnClustered"
    title = _chart_title(series, label_column, line)
    return {
        "available": True,
        "type": chart_type,
        "title": title,
        "columns": series,
        "label_column": label_column["name"] if label_column else None,
        "row_count": len(rows),
        "reason": (
            f"'{title}' compares {len(series)} numeric column(s)"
            + (f" across {label_column['name']}" if label_column else "")
            + (". Values change over time, so a line chart reads best." if line else ".")
        ),
    }


def _chart_title(series, label_column, line):
    parts = " vs ".join(series[:2]) if len(series) >= 2 else series[0]
    if label_column:
        return f"{parts} by {label_column['name']}"
    return f"{parts} chart"


# ---------------------------------------------------------------------- #
# Period extremes (which month/row was highest?)
# ---------------------------------------------------------------------- #


def _period_extremes(numeric_columns, rows, label_column):
    """Highest and lowest single row per numeric column.

    Column statistics answer "which column is biggest"; this answers "which
    period was biggest", which is the question users actually ask after seeing
    a table. Keeping it in the report means follow-ups need no new analysis.
    """
    extremes = []
    for column in numeric_columns:
        numbers = column["numbers"]
        if not numbers:
            continue
        number_rows = column["number_rows"]
        best = max(range(len(numbers)), key=lambda i: numbers[i])
        worst = min(range(len(numbers)), key=lambda i: numbers[i])
        extremes.append(
            {
                "column": column["name"],
                "label_column": label_column["name"] if label_column else None,
                "highest": _period_point(numbers, number_rows, best, label_column, rows),
                "lowest": _period_point(numbers, number_rows, worst, label_column, rows),
            }
        )
    return extremes


def _period_point(numbers, number_rows, index, label_column, rows):
    row_index = number_rows[index] if index < len(number_rows) else index
    label = _row_label(rows, label_column, row_index)
    return {
        "label": label,
        "row_index": row_index,
        "value": _round(numbers[index]),
        "where": label or f"row {row_index + 1}",
    }


# ---------------------------------------------------------------------- #
# Data quality
# ---------------------------------------------------------------------- #


def _data_quality(columns, rows):
    total_cells = sum(len(r) for r in rows)
    empty_cells = sum(1 for row in rows for cell in row if is_blank(cell))
    keys = [_key(tuple(row)) for row in rows]
    duplicates = len(keys) - len(set(keys))
    completeness = _round((total_cells - empty_cells) / total_cells * 100) if total_cells else 0
    return {
        "empty_cells": empty_cells,
        "total_cells": total_cells,
        "completeness_pct": completeness,
        "duplicate_rows": duplicates,
        "mixed_type_columns": [c["name"] for c in columns if c["type"] == "mixed"],
    }


# ---------------------------------------------------------------------- #
# Insights
# ---------------------------------------------------------------------- #


def _build_insights(stats, growth, trends, anomalies, extremes, chart, quality, scope):
    """3-5 short, traceable observations, each with the numbers behind it.

    Candidates are collected in descending usefulness order (peak period, totals,
    growth, anomalies, trend, data quality) and the first five are kept, so the
    cap never silently drops the fact a user is most likely to ask about.
    """
    ranked = sorted(
        (s for s in stats if s["type"] == "numeric" and s["count"]),
        key=lambda s: s["sum"],
        reverse=True,
    )
    candidates = []

    peak = _peak_insight(extremes)
    if peak:
        candidates.append(peak)

    if ranked:
        top = ranked[0]
        candidates.append(
            {
                "text": f"{top['column']} has the largest total: {top['sum']}.",
                "evidence": f"sum={top['sum']} over {top['count']} value(s)",
            }
        )

    for growth_row in growth:
        if growth_row.get("available") and growth_row.get("percent") is not None:
            candidates.append(_growth_insight(growth_row))
            break

    for finding in anomalies:
        if finding.get("count"):
            first = finding["items"][0]
            where = first["label"] or f"row {first['row_index'] + 1}"
            candidates.append(
                {
                    "text": (
                        f"{finding['column']} has {finding['count']} possible outlier(s); "
                        f"{first['value']} at {where} is one of them."
                    ),
                    "evidence": ", ".join(first["methods"]),
                }
            )
            break

    if len(ranked) > 1:
        bottom = ranked[-1]
        candidates.append(
            {
                "text": f"{bottom['column']} has the smallest total: {bottom['sum']}.",
                "evidence": f"sum={bottom['sum']} over {bottom['count']} value(s)",
            }
        )
        ratio = _share(ranked[0], ranked[-1])
        if ratio:
            candidates.append(
                {
                    "text": (
                        f"{ranked[0]['column']} is about {ratio} times "
                        f"{ranked[-1]['column']} in total."
                    ),
                    "evidence": f"{ranked[0]['sum']} vs {ranked[-1]['sum']}",
                }
            )

    for trend in trends:
        if trend["direction"] in ("increasing", "decreasing"):
            word = "generally rising" if trend["direction"] == "increasing" else "generally falling"
            candidates.append(
                {"text": f"{trend['column']} is {word}.", "evidence": trend["detail"]}
            )
            break

    if quality["duplicate_rows"]:
        candidates.append(
            {
                "text": f"{quality['duplicate_rows']} duplicate row(s) were found.",
                "evidence": "identical row values counted exactly once per extra copy",
            }
        )
    if quality["mixed_type_columns"]:
        candidates.append(
            {
                "text": ("Mixed types found in: " + ", ".join(quality["mixed_type_columns"]) + "."),
                "evidence": "columns contain both numbers and text",
            }
        )
    if chart.get("available"):
        candidates.append(
            {
                "text": f"A {chart['title']} chart would show this data well.",
                "evidence": chart["reason"],
            }
        )

    insights = _dedupe_insights(candidates)
    if not insights:
        insights.append(
            {
                "text": "No numeric patterns were found in this range.",
                "evidence": f"scope={scope['source']}, rows={scope['row_count']}",
            }
        )
    return insights[:5]


def _peak_insight(extremes):
    """The single highest cell in the table, named by its period."""
    for entry in extremes:
        highest = entry.get("highest") or {}
        where, value = highest.get("where"), highest.get("value")
        if where and value is not None:
            return {
                "text": f"{where} was the peak {entry['column']} period with {value}.",
                "evidence": f"max single value of {entry['column']} = {value} at {where}",
            }
    return None


def _growth_insight(growth_row):
    direction = "rose" if growth_row["change"] >= 0 else "fell"
    first_label = growth_row.get("first_label")
    last_label = growth_row.get("last_label")
    if first_label and last_label:
        span = f"from {first_label} to {last_label}"
    else:
        span = "from the first to the last value"
    return {
        "text": (f"{growth_row['column']} {direction} {abs(growth_row['percent'])}% {span}."),
        "evidence": (
            f"{growth_row['first']} -> {growth_row['last']} (change {growth_row['change']})"
        ),
    }


def _dedupe_insights(candidates):
    insights = []
    seen = set()
    for insight in candidates:
        key = insight["text"]
        if key in seen:
            continue
        seen.add(key)
        insights.append(insight)
    return insights


def _share(largest, smallest):
    if not smallest.get("sum") or largest.get("sum") is None:
        return None
    ratio = largest["sum"] / smallest["sum"]
    if ratio <= 1:
        return None
    return _round(ratio, 1)


# ---------------------------------------------------------------------- #
# Limitations
# ---------------------------------------------------------------------- #


def _limitations(stats, rows, scope):
    notes = ["Statistics describe only the cells included in this snapshot."]
    if scope["source"] == "selection":
        notes.append("Scope: your current selection, not the whole sheet.")
    for stat in stats:
        if stat["type"] == "numeric" and stat["count"] < MIN_ANOMALY_POINTS:
            notes.append(f"{stat['column']}: too few numeric values for outlier detection.")
    if not any(s["type"] == "numeric" and s["count"] for s in stats):
        notes.append("No numeric column was found, so no statistics could be computed.")
    return notes


# ---------------------------------------------------------------------- #
# Small helpers
# ---------------------------------------------------------------------- #


def _row_label(rows, label_column, row_index):
    if label_column is None or row_index >= len(rows):
        return None
    value = _cell(rows[row_index], label_column["index"])
    return None if is_blank(value) else str(value)


def _key(value):
    """Hashable, type-stable key for duplicate detection."""
    if is_blank(value):
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f"n:{float(value)}"
    return f"s:{str(value).strip().lower()}"


def _round(value, digits=4):
    if value is None or isinstance(value, bool):
        return value
    if not isinstance(value, (int, float)):
        return value
    if isinstance(value, int):
        return value
    if not math.isfinite(value):
        return None
    rounded = round(value, digits)
    return int(rounded) if rounded.is_integer() else rounded


def analyze_snapshot(headers=None, rows=None, selection=None, sheet_name=None):
    """Module-level convenience wrapper around :class:`DataAnalyzer`."""
    return DataAnalyzer.analyze(
        headers=headers, rows=rows, selection=selection, sheet_name=sheet_name
    )
