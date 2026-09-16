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
