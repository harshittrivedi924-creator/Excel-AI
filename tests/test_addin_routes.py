"""Tests for the Excel Add-in API (backend/api/addin_routes.py).

The Add-in sends a snapshot of the active worksheet captured through
Office.js. These tests exercise the endpoint contract end to end without
requiring a running Excel instance.
"""

import pytest


def revenue_payload(command, **overrides):
    payload = {
        "command": command,
        "sheet_name": "Sheet1",
        "origin": {"row": 1, "column": 1},
        "headers": ["Revenue", "Expense"],
        "rows": [[5000, 2000], [6000, 2500], [7000, 3000]],
    }
    payload.update(overrides)
    return payload


@pytest.fixture
def client():
    from backend.api.app import create_app

    app = create_app()
    app.config["TESTING"] = True
    return app.test_client()


def test_health(client):
    res = client.get("/api/excel/health")
    assert res.status_code == 200
    data = res.get_json()
    assert data["status"] == "ok"
    assert data["service"] == "excel-ai-copilot-addin-api"
    assert data["version"]


def test_valid_named_command_structured_response(client):
    res = client.post("/api/excel/command", json=revenue_payload("Revenue ka total karo"))
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert data["operation"] == "SUM"
    assert data["result"] == 18000
    assert data["message"]
    assert data["overwrite_needed"] is False
    assert data["writes"] == {"cells": [], "sheets": {}, "charts": []}
    assert data["parameters"] and data["parameters"].get("operation") == "SUM"


def test_average_and_max_and_min(client):
    avg = client.post(
        "/api/excel/command", json=revenue_payload("Revenue ka average nikalo")
    ).get_json()
    assert avg["success"] is True
    assert avg["operation"] == "AVERAGE"
    assert avg["result"] == 6000.0

    mx = client.post(
        "/api/excel/command", json=revenue_payload("Revenue ka maximum batao")
    ).get_json()
    assert mx["result"] == 7000

    mn = client.post(
        "/api/excel/command", json=revenue_payload("Revenue ka minimum batao")
    ).get_json()
    assert mn["result"] == 5000


def test_destination_writes_formula_cell(client):
    res = client.post(
        "/api/excel/command",
        json=revenue_payload("Revenue ka total karo aur A8 mein daal do"),
    )
    data = res.get_json()
    assert data["success"] is True
    assert data["result"] == 18000
    assert data["destination"] == "A8"
    cells = data["writes"]["cells"]
    assert {
        "row": 8,
        "column": 1,
        "formula": "=SUM(A2:A4)",
    } in cells


def test_overwrite_needed_and_retry(client):
    first = client.post(
        "/api/excel/command",
        json=revenue_payload("Revenue ka total karo aur B2 mein daal do"),
    )
    data = first.get_json()
    assert data["success"] is False
    assert data["overwrite_needed"] is True
    assert "already contains data" in data["message"]

    retry = client.post(
        "/api/excel/command",
        json=revenue_payload("Revenue ka total karo aur B2 mein daal do", allow_overwrite=True),
    )
    data2 = retry.get_json()
    assert data2["success"] is True
    assert {"row": 2, "column": 2, "formula": "=SUM(A2:A4)"} in data2["writes"]["cells"]


def test_invalid_command(client):
    res = client.post("/api/excel/command", json=revenue_payload("bjsh fhewk qw"))
    data = res.get_json()
    assert data["success"] is False
    assert data["message"]
    assert data["writes"] == {"cells": [], "sheets": {}, "charts": []}


def test_unknown_column_error_without_selection(client):
    res = client.post("/api/excel/command", json=revenue_payload("Profit ka total karo"))
    data = res.get_json()
    assert data["success"] is False
    assert "profit" in data["message"].lower()


def test_malformed_requests(client):
    bad = client.post("/api/excel/command", data="not json", content_type="text/plain")
    assert bad.status_code == 400

    missing_cmd = client.post("/api/excel/command", json={"command": "  "})
    assert missing_cmd.status_code == 400

    empty_cmd = client.post("/api/excel/command", json={})
    assert empty_cmd.status_code == 400


def test_empty_selection_error(client):
    payload = revenue_payload(
        "inka total karo",
        selection={
            "address": "D2:D4",
            "values": [[""], [""], [""]],
            "row": 2,
            "column": 4,
            "row_count": 3,
            "column_count": 1,
        },
    )
    data = client.post("/api/excel/command", json=payload).get_json()
    assert data["success"] is False
    assert "no numeric values" in data["message"]


def test_selection_fallback_sum(client):
    payload = revenue_payload(
        "inka total karo",
        selection={
            "address": "B2:B4",
            "values": [[5000], [6000], [7000]],
            "row": 2,
            "column": 2,
            "row_count": 3,
            "column_count": 1,
        },
    )
    data = client.post("/api/excel/command", json=payload).get_json()
    assert data["success"] is True
    assert data["operation"] == "SUM"
    assert data["result"] == 18000
    assert data["selection_used"] is True
    assert data["selection_auto_write"] == {"row": 5, "column": 2, "value": 18000}
    assert data["writes"] == {"cells": [], "sheets": {}, "charts": []}


def test_selection_fallback_average(client):
    payload = revenue_payload(
        "inko average karo",
        selection={
            "address": "B2:B4",
            "values": [[5000], [6000], [7000]],
            "row": 2,
            "column": 2,
            "row_count": 3,
            "column_count": 1,
        },
    )
    data = client.post("/api/excel/command", json=payload).get_json()
    assert data["success"] is True
    assert data["operation"] == "AVERAGE"
    assert data["result"] == 6000.0


def test_named_binary_writes_formula_column(client):
    payload = {
        "command": "Revenue minus Expense karke Profit nikalo",
        "sheet_name": "Sheet1",
        "origin": {"row": 1, "column": 1},
        "headers": ["Revenue", "Expense"],
        "rows": [[5000, 2000], [6000, 2500], [7000, 3000]],
    }
    data = client.post("/api/excel/command", json=payload).get_json()
    assert data["success"] is True
    assert data["operation"] == "SUBTRACT"
    cells = {(c["row"], c["column"], c.get("formula")) for c in data["writes"]["cells"]}
    assert (2, 3, "=A2-B2") in cells
    assert (3, 3, "=A3-B3") in cells
    assert (4, 3, "=A4-B4") in cells


def test_chart_writes_descriptor(client):
    payload = {
        "command": "Sales ka chart bana do",
        "sheet_name": "Sheet1",
        "origin": {"row": 1, "column": 1},
        "headers": ["Sales"],
        "rows": [[5], [6], [7], [8]],
    }
    data = client.post("/api/excel/command", json=payload).get_json()
    assert data["success"] is True
    assert data["operation"] == "CHART"
    assert len(data["writes"]["charts"]) == 1
    chart = data["writes"]["charts"][0]
    assert chart["type"] == "ColumnClustered"
    assert chart["source_range"] == "A1:A5"
    assert chart["title"]


def test_sort_rewrites_rows(client):
    payload = {
        "command": "A sort karo",
        "sheet_name": "Sheet1",
        "origin": {"row": 1, "column": 1},
        "headers": ["Count"],
        "rows": [[3], [1], [2]],
    }
    data = client.post("/api/excel/command", json=payload).get_json()
    assert data["success"] is True
    assert data["operation"] == "SORT"
    assert data["writes"]["cells"]


def test_cors_allows_configured_origin(client):
    import os

    expected = (
        (os.environ.get("EXCEL_ADDIN_ALLOWED_ORIGINS") or "https://localhost:3000")
        .split(",")[0]
        .strip()
    )
    res = client.get("/api/excel/health", headers={"Origin": expected})
    assert res.headers.get("Access-Control-Allow-Origin") == expected


def test_existing_api_unaffected(client):
    res = client.get("/api/workbook")
    assert res.status_code == 200
    data = res.get_json()
    assert data["loaded"] is False

    cmd = client.post("/api/command", json={"command": "sum karo"})
    assert cmd.status_code == 400
    assert cmd.headers.get("Access-Control-Allow-Origin") is None


# ---------------------------------------------------------------------- #
# Real-worksheet snapshot (Sheet1, used range A2:C7):
#   A2=a  B2=b  C2=d
#   A3=2555  B3=4544
#   A4=511   B4=3535
#   A5=214   B5=355
#   A6=3212135  B6=546
#   A7=35    B7=3545
# ---------------------------------------------------------------------- #


def user_sheet_payload(command, **overrides):
    payload = {
        "command": command,
        "sheet_name": "Sheet1",
        "origin": {"row": 2, "column": 1},
        "headers": ["a", "b", "d"],
        "rows": [
            [2555, 4544],
            [511, 3535],
            [214, 355],
            [3212135, 546],
            [35, 3545],
        ],
    }
    payload.update(overrides)
    return payload


def test_b_column_total_from_real_worksheet(client):
    res = client.post("/api/excel/command", json=user_sheet_payload("b ka total karo"))
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert data["operation"] == "SUM"
    assert data["result"] == 12525
    assert data["message"]
    assert data["writes"] == {"cells": [], "sheets": {}, "charts": []}


def test_add_two_columns_into_d_column(client):
    res = client.post(
        "/api/excel/command",
        json=user_sheet_payload("a aur b ka sum kar ke d mein daal do"),
    )
    data = res.get_json()
    assert data["success"] is True
    assert data["operation"] == "ADD"
    cells = data["writes"]["cells"]
    expected = [{"row": r, "column": 4, "formula": f"=A{r}+B{r}"} for r in range(3, 8)]
    assert [c for c in cells if c.get("formula", "").startswith("=A")] == expected


def test_add_two_columns_into_d_column_aor_typo(client):
    res = client.post(
        "/api/excel/command",
        json=user_sheet_payload("aor b ka sum kar ke d mein daal do"),
    )
    data = res.get_json()
    assert data["success"] is True
    assert data["operation"] == "ADD"
    formulas = sorted(c["formula"] for c in data["writes"]["cells"])
    assert formulas == [f"=A{r}+B{r}" for r in range(3, 8)]


def test_selection_sum_from_real_worksheet(client):
    payload = user_sheet_payload(
        "inka total karo",
        selection={
            "address": "Sheet1!A3:A7",
            "values": [[2555], [511], [214], [3212135], [35]],
            "row": 3,
            "column": 1,
            "row_count": 5,
            "column_count": 1,
        },
    )
    res = client.post("/api/excel/command", json=payload)
    data = res.get_json()
    assert data["success"] is True
    assert data["operation"] == "SUM"
    assert data["result"] == 3215450
    assert data["selection_used"] is True
    assert data["selection_auto_write"] == {"row": 8, "column": 1, "value": 3215450}


# ---------------------------------------------------------------------- #
# Real-Excel snapshot of the "Data" sheet (used range A1:D9):
#   A1=a  B1=b  C1=c  D1=d
#   A2:A6 / B2:B6 numbers, C2:C6 labels, D2:D8 empty, D9=42
# Office.js reports the empty cells inside the used range as "" rather
# than null, so blank padding is part of the real request payload.
# ---------------------------------------------------------------------- #


def blank_padded_payload(command, **overrides):
    payload = {
        "command": command,
        "sheet_name": "Data",
        "origin": {"row": 1, "column": 1},
        "headers": ["a", "b", "c", "d"],
        "rows": [
            [10, 100, "row1", ""],
            [20, 200, "row2", ""],
            [30, 300, "row3", ""],
            [40, 400, "row4", ""],
            [50, 500, "row5", ""],
            ["", "", "", ""],
            ["", "", "", ""],
            ["", "", "", 42],
        ],
    }
    payload.update(overrides)
    return payload


def test_blank_cells_are_not_treated_as_data(client):
    res = client.post(
        "/api/excel/command",
        json=blank_padded_payload("aor b ka sum kar ke d mein daal do"),
    )
    data = res.get_json()
    assert data["success"] is True
    formulas = [c["formula"] for c in data["writes"]["cells"] if c.get("formula")]
    # Only the rows holding data get a formula; the blank padding below the
    # table must not extend the write range.
    assert formulas == [f"=A{r}+B{r}" for r in range(2, 7)]
    # D9 holds 42 and is outside the write range, so it stays untouched and
    # blank cells are not reported as changes.
    assert not [c for c in data["writes"]["cells"] if c.get("clear")]
    assert not [c for c in data["writes"]["cells"] if c.get("row") == 9]


def test_element_wise_write_asks_before_overwriting(client):
    payload = blank_padded_payload(
        "aor b ka sum kar ke d mein daal do",
        rows=[
            [10, 100, "row1", 5],
            [20, 200, "row2", ""],
            [30, 300, "row3", ""],
            [40, 400, "row4", ""],
            [50, 500, "row5", ""],
        ],
    )
    data = client.post("/api/excel/command", json=payload).get_json()
    assert data["success"] is False
    assert data["overwrite_needed"] is True
    assert "already contains data" in data["message"]
    assert data["writes"]["cells"] == []

    data = client.post("/api/excel/command", json=dict(payload, allow_overwrite=True)).get_json()
    assert data["success"] is True
    formulas = sorted(c["formula"] for c in data["writes"]["cells"] if c.get("formula"))
    assert formulas == [f"=A{r}+B{r}" for r in range(2, 7)]


def test_percentage_column_write_asks_before_overwriting(client):
    payload = blank_padded_payload(
        "b ka percentage D2 mein daalo",
        rows=[
            [10, 100, "row1", 5],
            [20, 200, "row2", ""],
            [30, 300, "row3", ""],
        ],
    )
    data = client.post("/api/excel/command", json=payload).get_json()
    assert data["success"] is False
    assert data["overwrite_needed"] is True
    assert data["writes"]["cells"] == []

    data = client.post("/api/excel/command", json=dict(payload, allow_overwrite=True)).get_json()
    assert data["success"] is True
    assert [c["formula"] for c in data["writes"]["cells"] if c.get("formula")] == [
        "=B2/600*100",
        "=B3/600*100",
        "=B4/600*100",
    ]


def test_padded_used_range_reports_actual_data_rows(client):
    # Used range is A1:D9 because of D9=42, but column B only holds five data
    # rows. Chart and filter must report/count the 5 data rows, not 8.
    chart = client.post(
        "/api/excel/command", json=blank_padded_payload("b ka chart bana do")
    ).get_json()
    assert chart["success"] is True
    assert chart["writes"]["charts"][0]["source_range"] == "B1:B6"
    assert "5 rows of data" in chart["message"]
    assert "8 rows" not in chart["message"]

    filtered = client.post(
        "/api/excel/command",
        json=blank_padded_payload("Filter data jahan b 200 se zyada"),
    ).get_json()
    assert filtered["success"] is True
    assert filtered["result"]["matched"] == 3
    assert filtered["result"]["total"] == 5
    assert "3 of 5 rows" in filtered["message"]
    assert "of 8" not in filtered["message"]
