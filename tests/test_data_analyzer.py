"""Tests for the read-only data analyzer and the analyze endpoint.

The analyzer is deterministic and offline: every test here feeds a snapshot
and asserts the exact figures the Add-in will show the user.
"""

import pytest

from backend.excel.analysis_followup import answer_followup, classify_followup
from backend.excel.data_analyzer import DataAnalyzer

REVENUE_ROWS = [
    ["Jan", 1000, 600],
    ["Feb", 1200, 650],
    ["Mar", 1800, 900],
    ["Apr", 1600, 800],
    ["May", 2000, 950],
]


@pytest.fixture
def client():
    from backend.api.app import create_app

    app = create_app()
    app.config["TESTING"] = True
    return app.test_client()


def revenue_payload(**overrides):
    payload = {
        "sheet_name": "Sheet1",
        "origin": {"row": 1, "column": 1},
        "headers": ["Month", "Revenue", "Expense"],
        "rows": [list(r) for r in REVENUE_ROWS],
    }
    payload.update(overrides)
    return payload


def stats_for(report, column):
    for stat in report["statistics"]:
        if stat["column"] == column:
            return stat
    raise AssertionError(f"no statistics for {column}")


# ---------------------------------------------------------------------- #
# Profiling
# ---------------------------------------------------------------------- #


def test_profile_counts_rows_columns_and_types():
    report = DataAnalyzer.analyze(headers=["Month", "Revenue", "Expense"], rows=REVENUE_ROWS)

    profile = report["profile"]
    assert profile["rows"] == 5
    assert profile["columns"] == 3
    assert profile["column_names"] == ["Month", "Revenue", "Expense"]
    assert profile["numeric_columns"] == ["Revenue", "Expense"]
    assert profile["empty_cells"] == 0
    assert profile["completeness_pct"] == 100
    assert profile["duplicate_rows"] == 0
    assert profile["unique_values"]["Revenue"] == 5


def test_mixed_text_and_blank_columns_are_classified():
    report = DataAnalyzer.analyze(
        headers=["Name", "Score", "Note"],
        rows=[["Amit", 10, ""], ["Sara", "n/a", "ok"], ["", 30, "ok"]],
    )
    types = {s["column"]: s["type"] for s in report["statistics"]}
    assert types["Name"] == "text"
    assert types["Score"] == "mixed"
    assert types["Note"] == "text"
    assert report["profile"]["empty_cells"] == 2
    assert report["profile"]["completeness_pct"] < 100


def test_duplicate_rows_and_empty_column_detected():
    report = DataAnalyzer.analyze(
        headers=["A", "B", "C"],
        rows=[[1, 2, None], [1, 2, None], [4, 5, None], [7, 8, None]],
    )
    assert report["profile"]["duplicate_rows"] == 1
    assert "C" in report["profile"]["empty_columns"]


def test_numeric_strings_and_currency_are_parsed():
    report = DataAnalyzer.analyze(
        headers=["Amount", "Label"],
        rows=[["$1,200", "a"], ["1,500", "b"], ["900", "c"]],
    )
    stat = stats_for(report, "Amount")
    assert stat["sum"] == 3600
    assert stat["min"] == 900
    assert stat["max"] == 1500


# ---------------------------------------------------------------------- #
# Statistics
# ---------------------------------------------------------------------- #


def test_numeric_statistics_are_exact():
    report = DataAnalyzer.analyze(headers=["Month", "Revenue", "Expense"], rows=REVENUE_ROWS)

    revenue = stats_for(report, "Revenue")
    assert revenue["count"] == 5
    assert revenue["sum"] == 7600
    assert revenue["average"] == 1520
    assert revenue["min"] == 1000
    assert revenue["max"] == 2000
    assert revenue["median"] == 1600
    assert revenue["range"] == 1000

    expense = stats_for(report, "Expense")
    assert expense["sum"] == 3900
    assert expense["average"] == 780


def test_statistics_are_none_without_numbers():
    report = DataAnalyzer.analyze(headers=["Name"], rows=[["Amit"], ["Sara"]])
    stat = stats_for(report, "Name")
    assert stat["type"] == "text"
    assert stat["count"] == 0
    assert stat["sum"] is None
    assert stat["average"] is None


def test_growth_first_to_last():
    report = DataAnalyzer.analyze(headers=["Month", "Revenue", "Expense"], rows=REVENUE_ROWS)
    growth = {g["column"]: g for g in report["growth"]}

    revenue = growth["Revenue"]
    assert revenue["available"] is True
    assert revenue["first"] == 1000
    assert revenue["last"] == 2000
    assert revenue["change"] == 1000
    assert revenue["percent"] == 100


def test_growth_handles_zero_denominator():
    report = DataAnalyzer.analyze(headers=["V"], rows=[[0], [10], [5]])
    growth = report["growth"][0]
    assert growth["available"] is True
    assert growth["percent"] is None
    assert growth["change"] == 5
    assert "0" in growth["reason"]


def test_growth_unavailable_with_single_value():
    report = DataAnalyzer.analyze(headers=["V"], rows=[[7]])
    growth = report["growth"][0]
    assert growth["available"] is False
    assert "2 numeric values" in growth["reason"]


# ---------------------------------------------------------------------- #
# Trends and anomalies
# ---------------------------------------------------------------------- #


def test_increasing_and_decreasing_trends():
    up = DataAnalyzer.analyze(headers=["M", "V"], rows=[["a", 1], ["b", 2], ["c", 3]])
    assert {t["column"]: t["direction"] for t in up["trends"]}["V"] == "increasing"

    down = DataAnalyzer.analyze(headers=["M", "V"], rows=[["a", 9], ["b", 5], ["c", 1]])
    assert {t["column"]: t["direction"] for t in down["trends"]}["V"] == "decreasing"


def test_trend_reports_insufficient_data_instead_of_guessing():
    report = DataAnalyzer.analyze(headers=["M", "V"], rows=[["a", 1], ["b", 2]])
    trend = {t["column"]: t for t in report["trends"]}["V"]
    assert trend["direction"] == "insufficient_data"
    assert "at least 3" in trend["detail"]


def test_outlier_is_flagged_with_reason():
    report = DataAnalyzer.analyze(
        headers=["Month", "Revenue"],
        rows=[["Jan", 100], ["Feb", 110], ["Mar", 105], ["Apr", 102], ["May", 5000]],
    )
    finding = {f["column"]: f for f in report["anomalies"]}["Revenue"]
    assert finding["count"] == 1
    assert finding["items"][0]["value"] == 5000
    assert finding["items"][0]["label"] == "May"
    assert finding["items"][0]["methods"]


def test_no_outliers_is_reported_explicitly():
    report = DataAnalyzer.analyze(headers=["Month", "Revenue", "Expense"], rows=REVENUE_ROWS)
    for finding in report["anomalies"]:
        assert finding["count"] == 0
        assert "No values" in finding["note"]


def test_anomalies_need_enough_values():
    report = DataAnalyzer.analyze(headers=["V"], rows=[[1], [2], [3]])
    assert report["anomalies"][0]["count"] == 0
    assert "at least 4" in report["anomalies"][0]["note"]


# ---------------------------------------------------------------------- #
# Insights, chart recommendation, limitations
# ---------------------------------------------------------------------- #


def test_insights_are_traceable_and_capped():
    report = DataAnalyzer.analyze(headers=["Month", "Revenue", "Expense"], rows=REVENUE_ROWS)
    insights = report["insights"]
    assert 3 <= len(insights) <= 5
    for insight in insights:
        assert insight["text"]
        assert insight["evidence"]
    assert any("Revenue has the largest total" in i["text"] for i in insights)
    assert any(i["evidence"].startswith("sum=") for i in insights)


def test_chart_recommendation_uses_existing_types_only():
    report = DataAnalyzer.analyze(headers=["Month", "Revenue", "Expense"], rows=REVENUE_ROWS)
    chart = report["chart_recommendation"]
    assert chart["available"] is True
    assert chart["type"] in ("LineChart", "ColumnClustered")
    assert chart["columns"] == ["Revenue", "Expense"]
    assert "Month" in chart["reason"]


def test_chart_recommendation_absent_without_numbers():
    report = DataAnalyzer.analyze(headers=["Name"], rows=[["Amit"], ["Sara"]])
    chart = report["chart_recommendation"]
    assert chart["available"] is False
    assert "No numeric column" in chart["reason"]


def test_empty_snapshot_is_handled_gracefully():
    report = DataAnalyzer.analyze(headers=[], rows=[])
    assert report["profile"]["rows"] == 0
    assert report["statistics"] == []
    assert report["chart_recommendation"]["available"] is False
    assert report["insights"]
    assert any("No numeric column" in note for note in report["limitations"])


def test_limitations_mention_selection_scope():
    report = DataAnalyzer.analyze(
        headers=["M", "V"],
        rows=[["a", 1], ["b", 2], ["c", 3]],
        selection={"address": "A1:B4", "values": [["a", 1], ["b", 2], ["c", 3]]},
    )
    assert report["scope"]["source"] == "selection"
    assert any("selection" in note for note in report["limitations"])


def test_large_but_uncapped_input_makes_no_truncation_claim():
    """A big range is only 'partial' when it was actually capped."""
    report = DataAnalyzer.analyze(
        headers=["M", "V"],
        rows=[[f"r{i}", i] for i in range(1200)],
    )
    assert report["profile"]["rows"] == 1200
    assert not any("Large dataset" in note for note in report["limitations"])
    assert not any("capped" in note for note in report["limitations"])


def test_analysis_is_deterministic():
    first = DataAnalyzer.analyze(headers=["Month", "Revenue", "Expense"], rows=REVENUE_ROWS)
    second = DataAnalyzer.analyze(headers=["Month", "Revenue", "Expense"], rows=REVENUE_ROWS)
    assert first == second


def test_empty_selection_falls_back_to_used_range():
    report = DataAnalyzer.analyze(
        headers=["Month", "Revenue"],
        rows=[["Jan", 1000], ["Feb", 1200]],
        selection={"address": "A1:A1", "values": [[None]]},
    )
    assert report["scope"]["source"] == "worksheet"
    assert report["profile"]["rows"] == 2


# ---------------------------------------------------------------------- #
# /api/excel/analyze
# ---------------------------------------------------------------------- #


def test_analyze_endpoint_contract(client):
    res = client.post("/api/excel/analyze", json=revenue_payload())
    assert res.status_code == 200
    data = res.get_json()

    assert data["success"] is True
    assert data["sheet"] == "Sheet1"
    assert data["profile"]["rows"] == 5
    assert data["profile"]["column_names"] == ["Month", "Revenue", "Expense"]
    assert stats_for(data, "Revenue")["sum"] == 7600
    assert data["growth"][0]["column"] in ("Revenue", "Expense")
    assert data["trends"] and data["anomalies"]
    assert 3 <= len(data["insights"]) <= 5
    assert data["chart_recommendation"]["available"] is True
    assert data["limitations"]
    assert "Analyzed 5 row(s)" in data["message"]
    # analysis must never write to the workbook
    assert data["writes"] == {"cells": [], "sheets": {}, "charts": []}


def test_analyze_endpoint_does_not_break_command_endpoint(client):
    payload = revenue_payload(command="Revenue ka total karo")
    before = client.post("/api/excel/command", json=payload).get_json()
    client.post("/api/excel/analyze", json=revenue_payload())
    after = client.post("/api/excel/command", json=payload).get_json()
    assert before == after
    assert after["result"] == 7600
    assert client.get("/api/excel/health").status_code == 200


def test_analyze_endpoint_handles_malformed_requests(client):
    assert client.post("/api/excel/analyze", json=None).status_code == 400
    assert client.post("/api/excel/analyze", json=[]).status_code == 400
    res = client.post("/api/excel/analyze", json={"headers": None, "rows": "nope"}).get_json()
    assert res["success"] is False
    assert "No data" in res["message"]


def test_analyze_endpoint_caps_large_snapshots(client):
    big_rows = [[f"r{i}", i, i * 2] for i in range(6000)]
    res = client.post("/api/excel/analyze", json=revenue_payload(rows=big_rows))
    assert res.status_code == 200
    data = res.get_json()
    assert data["truncated"] is True
    assert data["profile"]["rows"] == 5000
    assert any("capped" in note for note in data["limitations"])


def test_analyze_endpoint_handles_mixed_data(client):
    rows = [["Name", "Score", "Note"], ["Amit", 10, ""], ["Sara", "n/a", "ok"]]
    res = client.post(
        "/api/excel/analyze",
        json={"headers": ["Name", "Score", "Note"], "rows": rows},
    )
    data = res.get_json()
    assert data["success"] is True
    assert "no numeric column" in data["message"] or data["profile"]["numeric_columns"]
    assert data["data_quality"]["empty_cells"] == 1
    assert data["data_quality"]["mixed_type_columns"] == ["Score"]


def test_analyze_endpoint_uses_selection_when_present(client):
    res = client.post(
        "/api/excel/analyze",
        json=revenue_payload(
            selection={
                "address": "B1:B6",
                "values": [["Revenue"], [1000], [1200], [1800], [1600], [2000]],
                "row": 1,
                "column": 2,
                "row_count": 6,
                "column_count": 1,
            }
        ),
    )
    data = res.get_json()
    assert data["scope"]["source"] == "selection"
    assert data["scope"]["address"] == "B1:B6"
    assert stats_for(data, "Revenue")["sum"] == 7600


# ---------------------------------------------------------------------- #
# Follow-ups
# ---------------------------------------------------------------------- #


@pytest.fixture
def report():
    return DataAnalyzer.analyze(headers=["Month", "Revenue", "Expense"], rows=REVENUE_ROWS)


def test_followup_highest_and_lowest(report):
    highest = answer_followup("which month had the highest revenue?", report)
    assert highest["handled"] is True
    assert "Revenue" in highest["message"]
    assert "2000" in highest["message"]

    lowest = answer_followup("lowest expense", report)
    assert lowest["intent"] == "lowest"
    assert "600" in lowest["message"]


def test_followup_growth(report):
    answer = answer_followup("what is the growth?", report)
    assert answer["intent"] == "growth"
    assert "100%" in answer["message"]
    assert "rose" in answer["message"]


@pytest.mark.parametrize(
    "question",
    ["which is highest?", "which is lowest?", "highest column"],
)
def test_extreme_questions_do_not_raise_on_a_text_only_sheet(question):
    """A text-only selection has no numeric column to rank.

    ``_answer_column_extreme`` returns a bare message while the dispatcher
    unpacks a ``(message, chart_command)`` pair, so this used to raise a
    ValueError and surface as an HTTP 500 instead of an honest answer.
    """
    from backend.excel.data_analyzer import DataAnalyzer

    text_only = DataAnalyzer.analyze(
        headers=["Name", "Note"],
        rows=[["Asha", "ok"], ["Bilal", "ok"], ["Chetna", "late"]],
        sheet_name="TextOnly",
    )

    answer = answer_followup(question, text_only)

    assert answer["handled"] is True
    assert answer["chart_command"] is None
    assert "no numeric column" in answer["message"].lower()


def test_extreme_questions_still_answer_a_numeric_sheet():
    """The same call must keep working where a numeric column does exist."""
    from backend.excel.data_analyzer import DataAnalyzer

    numeric = DataAnalyzer.analyze(
        headers=["Month", "Revenue"],
        rows=[["Jan", 10], ["Feb", 20], ["Mar", 30]],
        sheet_name="Revenue",
    )

    answer = answer_followup("which is highest?", numeric)

    assert answer["handled"] is True
    assert "Mar" in answer["message"]
    assert "30" in answer["message"]


def test_followup_anomalies_without_outliers(report):
    answer = answer_followup("any anomalies?", report)
    assert answer["intent"] == "anomalies"
    assert "No values" in answer["message"]


def test_followup_anomalies_with_outlier():
    noisy = DataAnalyzer.analyze(
        headers=["Month", "Revenue"],
        rows=[["Jan", 100], ["Feb", 110], ["Mar", 105], ["Apr", 102], ["May", 5000]],
    )
    answer = answer_followup("show me the outliers", noisy)
    assert "5000" in answer["message"]


def test_followup_chart_returns_pipeline_command(report):
    answer = answer_followup("create the recommended chart", report)
    assert answer["intent"] == "chart"
    assert answer["chart_command"]
    assert "chart" in answer["chart_command"]


def test_followup_chart_command_creates_a_chart(client, report):
    answer = answer_followup("create the recommended chart", report)
    res = client.post("/api/excel/command", json=revenue_payload(command=answer["chart_command"]))
    data = res.get_json()
    assert data["success"] is True
    assert data["writes"]["charts"]
    assert data["writes"]["charts"][0]["type"] in ("LineChart", "ColumnClustered")


def test_followup_quality_and_summary(report):
    quality = answer_followup("how is the data quality?", report)
    assert quality["intent"] == "quality"
    assert "100%" in quality["message"]

    summary = answer_followup("give me a summary", report)
    assert summary["intent"] == "summary"
    assert summary["message"]


def test_unknown_followup_is_not_handled(report):
    answer = answer_followup("what is the meaning of life?", report)
    assert answer["handled"] is False
    assert answer["chart_command"] is None


def test_followup_without_previous_analysis():
    answer = answer_followup("which is highest?", None)
    assert answer["handled"] is False
    assert "Run an analysis first" in answer["message"]


@pytest.mark.parametrize(
    "text,expected",
    [
        ("highest month", "highest"),
        ("lowest", "lowest"),
        ("growth", "growth"),
        ("kitni badhi", "growth"),
        ("any anomaly", "anomalies"),
        ("trend batao", "trend"),
        ("missing values", "quality"),
        ("summary", "summary"),
        ("create a chart", "chart"),
        ("", "unknown"),
    ],
)
def test_followup_classification(text, expected):
    assert classify_followup(text)[0] == expected


@pytest.mark.parametrize(
    "question",
    [
        "What is important in this data?",
        "Is data mein kya important hai?",
        "important baatein batao",
    ],
)
def test_important_questions_are_answered_from_the_existing_report(client, report, question):
    """A follow-up must not fall through to "I can't answer that".

    The task pane treats "important" as a fresh-analysis request
    (``isAnalysisRequest``), so the follow-up router has to agree with it.
    """
    res = client.post("/api/excel/analyze", json={"question": question, "analysis": report})
    assert res.status_code == 200
    data = res.get_json()
    assert data["handled"] is True
    assert data["followup"]["intent"] == "summary"
    assert "can't answer" not in data["message"].lower()
    assert "2000" in data["message"]
    assert data["writes"] == {"cells": [], "sheets": {}, "charts": []}


def test_followup_via_endpoint(client, report):
    res = client.post(
        "/api/excel/analyze",
        json={"question": "which month was highest?", "analysis": report},
    )
    assert res.status_code == 200
    data = res.get_json()
    assert data["handled"] is True
    assert data["followup"]["intent"] == "highest"
    assert "2000" in data["message"]
    assert data["writes"] == {"cells": [], "sheets": {}, "charts": []}


# ---------------------------------------------------------------------- #
# Period-aware extremes (which month was highest?)
# ---------------------------------------------------------------------- #


def test_period_extremes_name_the_peak_and_trough_row():
    report = DataAnalyzer.analyze(
        headers=["Month", "Revenue", "Expense"], rows=REVENUE_ROWS, sheet_name="Data"
    )
    extremes = {e["column"]: e for e in report["period_extremes"]}
    assert extremes["Revenue"]["highest"]["label"] == "May"
    assert extremes["Revenue"]["highest"]["value"] == 2000
    assert extremes["Revenue"]["lowest"]["label"] == "Jan"
    assert extremes["Expense"]["highest"]["value"] == 950
    assert extremes["Revenue"]["label_column"] == "Month"


def test_period_extremes_fall_back_to_row_number_without_labels():
    report = DataAnalyzer.analyze(headers=["Value"], rows=[[10], [20], [30], [40], [50]])
    highest = report["period_extremes"][0]["highest"]
    assert highest["label"] is None
    assert highest["where"] == "row 5"
    assert highest["value"] == 50


def test_highest_month_followup_answers_the_period_not_the_column():
    report = DataAnalyzer.analyze(headers=["Month", "Revenue", "Expense"], rows=REVENUE_ROWS)
    answer = answer_followup("which month was highest?", report)
    assert answer["message"].startswith("May")
    assert "2000" in answer["message"]


def test_lowest_month_followup_answers_the_period():
    report = DataAnalyzer.analyze(headers=["Month", "Revenue", "Expense"], rows=REVENUE_ROWS)
    answer = answer_followup("lowest month for revenue", report)
    assert answer["intent"] == "lowest"
    assert answer["message"].startswith("Jan")
    assert "1000" in answer["message"]


def test_explicit_column_question_compares_columns_not_periods():
    report = DataAnalyzer.analyze(headers=["Month", "Revenue", "Expense"], rows=REVENUE_ROWS)
    answer = answer_followup("which column is highest?", report)
    assert "Revenue is the highest" in answer["message"]
    assert "Expense" in answer["message"]


def test_period_labels_survive_gaps_in_the_numeric_column():
    rows = [
        ["Jan", 100, 10],
        ["Feb", None, 20],
        ["Mar", 300, 30],
        ["Apr", "", 40],
        ["May", 500, 50],
    ]
    report = DataAnalyzer.analyze(headers=["Month", "Revenue", "Note"], rows=rows)
    extremes = {e["column"]: e for e in report["period_extremes"]}
    assert extremes["Revenue"]["highest"]["label"] == "May"
    assert extremes["Revenue"]["lowest"]["label"] == "Jan"
    growth = {g["column"]: g for g in report["growth"]}
    assert growth["Revenue"]["first_label"] == "Jan"
    assert growth["Revenue"]["last_label"] == "May"


def test_anomaly_row_index_points_at_the_real_worksheet_row():
    rows = [
        ["Jan", 100, 10],
        ["Feb", None, 20],
        ["Mar", 110, 30],
        ["Apr", 105, 40],
        ["May", 108, 50],
        ["Jun", 9000, 60],
    ]
    report = DataAnalyzer.analyze(headers=["Month", "Revenue", "Other"], rows=rows)
    finding = next(f for f in report["anomalies"] if f["column"] == "Revenue")
    assert finding["count"] >= 1
    assert finding["items"][0]["row_index"] == 5
    assert finding["items"][0]["label"] == "Jun"


def test_trend_change_uses_the_real_row_label():
    rows = [["Jan", 10], ["Feb", None], ["Mar", 12], ["Apr", 14], ["May", 60]]
    report = DataAnalyzer.analyze(headers=["Month", "Revenue"], rows=rows)
    trend = next(t for t in report["trends"] if t["column"] == "Revenue")
    assert "at May" in trend["detail"]


def test_followup_column_matching_uses_arbitrary_headers():
    headers = ["Region", "Net Revenue", "Headcount"]
    rows = [["North", 900, 4], ["South", 1500, 6], ["East", 1200, 5]]
    report = DataAnalyzer.analyze(headers=headers, rows=rows)
    assert classify_followup("highest net revenue", headers) == ("highest", "Net Revenue")
    answer = answer_followup("what is the growth in headcount?", report)
    assert "Headcount" in answer["message"]


def test_column_matching_prefers_the_longest_header():
    headers = ["Revenue", "Net Revenue"]
    assert classify_followup("highest net revenue", headers) == ("highest", "Net Revenue")


def test_column_matching_respects_word_boundaries():
    headers = ["Cost", "Customer"]
    assert classify_followup("highest cost", headers) == ("highest", "Cost")
    assert classify_followup("how many customers", headers) == ("unknown", "Customer")


def test_growth_percent_is_readable(report):
    growth = {g["column"]: g for g in report["growth"]}
    assert growth["Revenue"]["percent"] == 100
    assert growth["Expense"]["percent"] == 58.3


def test_peak_insight_leads_the_insight_list():
    report = DataAnalyzer.analyze(headers=["Month", "Revenue", "Expense"], rows=REVENUE_ROWS)
    assert "May" in report["insights"][0]["text"]
    assert "2000" in report["insights"][0]["text"]


def test_chart_followup_names_the_charted_column(report):
    answer = answer_followup("create the recommended chart", report)
    assert answer["message"].startswith("Recommended: Revenue")
    assert "Expense separately" in answer["message"]


def test_analyze_endpoint_exposes_period_extremes(client):
    res = client.post(
        "/api/excel/analyze",
        json={
            "headers": ["Month", "Revenue"],
            "rows": [["Jan", 10], ["Feb", 50], ["Mar", 20], ["Apr", 30], ["May", 40]],
        },
    )
    data = res.get_json()
    assert data["success"] is True
    assert data["period_extremes"][0]["highest"]["label"] == "Feb"
    assert data["writes"] == {"cells": [], "sheets": {}, "charts": []}


# ---------------------------------------------------------------------- #
# Trend direction: a wobbly series still reports its net direction
# ---------------------------------------------------------------------- #


def trend_for(report, column):
    for trend in report["trends"]:
        if trend["column"] == column:
            return trend
    raise AssertionError(f"no trend for {column}")


def test_wobbly_series_reports_net_direction_not_mixed():
    # The Phase 4 reference dataset rises three times and falls once, so the
    # series is increasing overall. The detail must still show the dip.
    report = DataAnalyzer.analyze(headers=["Month", "Revenue"], rows=[r[:2] for r in REVENUE_ROWS])
    trend = trend_for(report, "Revenue")

    assert trend["direction"] == "increasing"
    assert trend["rises"] == 3
    assert trend["falls"] == 1
    assert trend["net_change"] == 1000
    assert "3 step(s) up" in trend["detail"]
    assert "1 step(s) down" in trend["detail"]


def test_wobbly_falling_series_reports_net_direction_not_mixed():
    rows = [["a", 900], ["b", 400], ["c", 800], ["d", 200], ["e", 500], ["f", 100]]
    trend = trend_for(DataAnalyzer.analyze(headers=["M", "V"], rows=rows), "V")

    assert trend["direction"] == "decreasing"
    assert trend["net_change"] == -800


def test_balanced_whipsaw_stays_mixed():
    # Rises and falls cancel out (2 up, 2 down), so no direction is claimed.
    rows = [["a", 10], ["b", 40], ["c", 20], ["d", 30], ["e", 10]]
    trend = trend_for(DataAnalyzer.analyze(headers=["M", "V"], rows=rows), "V")

    assert trend["rises"] == 2
    assert trend["falls"] == 2
    assert trend["direction"] == "mixed"
    assert trend["net_change"] == 0


def test_rise_and_fall_with_opposite_net_stays_mixed():
    # More rises than falls, but the series ends below where it started.
    rows = [["a", 100], ["b", 300], ["c", 200], ["d", 50], ["e", 60]]
    trend = trend_for(DataAnalyzer.analyze(headers=["M", "V"], rows=rows), "V")

    assert trend["direction"] == "mixed"


def test_flat_series_reports_flat_and_zero_net():
    trend = trend_for(
        DataAnalyzer.analyze(headers=["M", "V"], rows=[["a", 5], ["b", 5], ["c", 5]]), "V"
    )

    assert trend["direction"] == "flat"
    assert trend["net_change"] == 0


def test_insufficient_trend_reports_null_net_change():
    trend = trend_for(DataAnalyzer.analyze(headers=["M", "V"], rows=[["a", 1], ["b", 2]]), "V")

    assert trend["direction"] == "insufficient_data"
    assert trend["net_change"] is None


def test_every_trend_exposes_the_same_keys():
    report = DataAnalyzer.analyze(headers=["Month", "Revenue"], rows=[r[:2] for r in REVENUE_ROWS])
    for trend in report["trends"]:
        assert {"column", "direction", "detail", "rises", "falls", "flat", "net_change"} <= set(
            trend
        )


def test_increasing_trend_reaches_the_insight_list():
    report = DataAnalyzer.analyze(headers=["Month", "Revenue"], rows=[r[:2] for r in REVENUE_ROWS])
    texts = [i["text"] for i in report["insights"]]
    assert any("generally rising" in t for t in texts)


# ---------------------------------------------------------------------- #
# Comparison follow-up
# ---------------------------------------------------------------------- #


@pytest.fixture
def revenue_report():
    return DataAnalyzer.analyze(headers=["Month", "Revenue", "Expense"], rows=REVENUE_ROWS)


@pytest.mark.parametrize(
    "question",
    [
        "Compare revenue and expenses",
        "Compare revenue vs expense",
        "What is the difference between revenue and expense",
        "revenue aur expense ka antar",
    ],
)
def test_compare_followup_is_classified_and_answered(revenue_report, question):
    answer = answer_followup(question, revenue_report)

    assert answer["intent"] == "compare"
    assert answer["handled"] is True
    assert "7600" in answer["message"]
    assert "3900" in answer["message"]
    assert answer["chart_command"] is None


def test_compare_followup_reports_the_signed_gap_and_ratio(revenue_report):
    message = answer_followup("Compare revenue and expenses", revenue_report)["message"]

    assert "+3700" in message
    assert "1.9x" in message


def test_compare_followup_uses_arbitrary_headers():
    report = DataAnalyzer.analyze(headers=["Month", "Income", "Outgo"], rows=REVENUE_ROWS)
    message = answer_followup("Compare income and Outgo", report)["message"]

    assert "Income" in message
    assert "Outgo" in message


def test_compare_followup_without_a_usable_pair_lists_every_column(revenue_report):
    message = answer_followup("Compare revenue", revenue_report)["message"]

    assert "Name the two you want to compare" in message
    assert "Revenue (7600)" in message
    assert "Expense (3900)" in message


def test_compare_followup_without_numeric_columns_is_handled_honestly():
    report = DataAnalyzer.analyze(headers=["Name", "Note"], rows=[["a", "x"], ["b", "y"]])
    message = answer_followup("Compare name and note", report)["message"]

    assert "no numeric column" in message.lower()


def test_compare_followup_survives_a_zero_column():
    report = DataAnalyzer.analyze(headers=["M", "A", "B"], rows=[["a", 0, 0], ["b", 0, 0]])
    message = answer_followup("Compare A and B", report)["message"]

    assert "A" in message and "B" in message


def test_compare_is_answered_via_the_endpoint(client, revenue_report):
    res = client.post(
        "/api/excel/analyze",
        json={"question": "Compare revenue and expenses", "analysis": revenue_report},
    )
    data = res.get_json()

    assert res.status_code == 200
    assert data["followup"]["intent"] == "compare"
    assert data["handled"] is True
    assert data["writes"] == {"cells": [], "sheets": {}, "charts": []}


def test_plural_insights_are_classified_as_summary():
    assert classify_followup("insights dedo")[0] == "summary"


def test_comparing_an_unknown_column_name_does_not_raise(revenue_report):
    # "sales" is not a column here; the answer must degrade, not blow up.
    answer = answer_followup("Compare sales and expense", revenue_report)

    assert answer["handled"] is True
    assert "Expense" in answer["message"]


def test_column_matching_returns_every_named_column():
    from backend.excel.analysis_followup import _columns_in

    assert _columns_in("compare revenue and expenses", ["Revenue", "Expense"]) == [
        "Revenue",
        "Expense",
    ]


def test_a_blank_header_cell_falls_back_to_a_positional_name():
    """A user can write a formula into a column they never headed.

    The task pane sends the whole used range, so the header row can hold a
    blank cell. That column must still be named, otherwise the report reaches
    the task pane with ``column_names`` containing ``None``.
    """
    from backend.excel.data_analyzer import DataAnalyzer

    report = DataAnalyzer.analyze(
        headers=["a", "b", "c", None],
        rows=[[10, 20, 30, 30], [40, 50, 60, 90], [70, 80, 90, 150]],
        sheet_name="ABC",
    )

    assert report["profile"]["column_names"] == ["a", "b", "c", "Column 4"]
    assert all(name for name in report["profile"]["column_names"])
    # The unnamed column is still analysed, not dropped.
    assert "Column 4" in report["profile"]["numeric_columns"]
    # Every name the task pane can render is a real string.
    for stat in report["statistics"]:
        assert isinstance(stat["column"], str) and stat["column"]


def test_an_all_blank_header_row_does_not_raise():
    from backend.excel.data_analyzer import DataAnalyzer

    report = DataAnalyzer.analyze(headers=[None, None], rows=[[1, 2], [3, 4]])

    assert all(name for name in report["profile"]["column_names"])
    assert report["success"] is True


def test_column_matching_does_not_double_count_singular_and_plural():
    from backend.excel.analysis_followup import _columns_in

    assert _columns_in("compare customers and cost", ["Customer", "Cost"]) == [
        "Customer",
        "Cost",
    ]
