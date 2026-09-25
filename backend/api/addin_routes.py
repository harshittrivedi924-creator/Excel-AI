"""Excel Add-in API endpoints.

These endpoints are consumed **only** by the Excel Add-in task pane and are
fully additive: the existing ``/api/*`` endpoints, the CLI, the Web UI and
every downstream module (parser, validator, calculator, copilot) are left
untouched.

Data flow
---------
1. The task pane reads a snapshot of the active worksheet through Office.js
   (headers, data rows and the origin coordinate) plus the user's current
   selection and the natural-language command.
2. This snapshot is posted to ``POST /api/excel/command``.
3. The snapshot is materialised into an in-memory openpyxl workbook at real
   worksheet coordinates (see ``backend/excel/memory_handler.py``) and run
   through the existing ``ExcelCopilot`` pipeline.
4. The backend never touches the user's Excel instance or any file on disk.
   What changed inside the in-memory workbook is returned as a
   machine-readable ``writes`` descriptor that the Add-in applies back to
   the open workbook with Office.js.

``writes`` shape::

    {
      "cells":  [{"row": 8, "column": 3, "formula": "=SUM(B2:B4)"}],
                 {"row": 9, "column": 3, "clear": true},
                 {"row": 5, "column": 2, "value": 15000},
      "sheets": {"Filtered": {"values": [[...], [...]]}},
      "charts": [{"type": "ColumnClustered", "title": "...",
                  "source_range": "B1:B4"}]
    }
"""

import re
from datetime import date, datetime

from flask import Blueprint, jsonify, request
from openpyxl.utils import get_column_letter

from backend.calculator.engine import CalculationEngine
from backend.copilot import ExcelCopilot
from backend.excel.memory_handler import MemoryExcelHandler, is_blank
from backend.parser.parser import CommandParser
from backend.validator.validator import Validator

addin_bp = Blueprint("excel_addin", __name__)

AGGREGATE_OPERATIONS = ("SUM", "AVERAGE", "MIN", "MAX", "COUNT")
BINARY_OPERATIONS = ("ADD", "SUBTRACT", "MULTIPLY", "DIVIDE", "DIFFERENCE", "GROWTH")

#: Messages raised by the parser when no data source could be resolved.
#: In that case the Add-in may fall back to commands that operate on the
#: user's current selection (e.g. ``"inka total karo"``).
_DATA_SOURCE_ERROR_PREFIX = "I couldn't figure out which data to use"


# ---------------------------------------------------------------------- #
# Routes
# ---------------------------------------------------------------------- #


@addin_bp.route("/health", methods=["GET"])
def health():
    """Health check used by the task pane to detect backend availability."""
    return jsonify(
        {
            "status": "ok",
            "service": "excel-ai-copilot-addin-api",
            "version": "1.0.0",
        }
    )


@addin_bp.route("/command", methods=["POST"])
def excel_command():
    """Execute a natural-language command against a workbook snapshot.

    Request body (JSON)::

        {
          "command": "Revenue ka total karo",
          "sheet_name": "Sheet1",
          "origin": {"row": 1, "column": 1},
          "headers": ["Revenue", "Expense"],
          "rows": [[5000, 2000], [6000, 2500], [7000, 3000]],
          "selection": {
            "address": "B2:B4",
            "values": [[5000], [6000], [7000]],
            "row": 2, "column": 2, "row_count": 3, "column_count": 1
          },
          "allow_overwrite": false
        }
    """
    payload = request.get_json(force=True, silent=True)
    if not payload or not isinstance(payload, dict):
        return jsonify(_error_response("Invalid request. Expected a JSON body.")), 400

    command = (payload.get("command") or "").strip()
    if not command:
        return jsonify(_error_response("No command was provided.")), 400

    sheet_name = payload.get("sheet_name") or "Sheet1"
    headers = payload.get("headers") or []
    rows = payload.get("rows") or []
    origin = payload.get("origin") or {}
    selection = payload.get("selection") or None
    allow_overwrite = bool(payload.get("allow_overwrite", False))

    origin_row = int(origin.get("row") or 1)
    origin_col = int(origin.get("column") or 1)

    handler = MemoryExcelHandler(
        headers=headers,
        rows=rows,
        origin_row=origin_row,
        origin_col=origin_col,
        sheet_name=sheet_name,
    )
    snapshot = _build_snapshot(headers, rows, origin_row, origin_col)

    copilot = ExcelCopilot()
    copilot.excel = handler
    copilot.validator = Validator(handler)

    result = copilot.process_command(
        command, sheet_name=handler.sheet_name, allow_overwrite=allow_overwrite
    )

    if result["success"]:
        writes = _compute_writes(handler, snapshot, _chart_writes(result, origin_row))
        return jsonify(_success_response(command, result, writes))

    # The parser/validation could not resolve a data source. Two cases are
    # re-routed onto the user's current selection:
    #   1. "I couldn't figure out which data to use..." (no source found)
    #   2. "I couldn't find a column named 'X'..."  where X is not a real
    #      workbook header (e.g. "inka total karo" resolves "total" as a
    #      named column that doesn't exist). When a selection is present,
    #      the user almost certainly means the selected cells.
    if selection and _command_needs_selection(result["message"]):
        operation = _detect_operation(command)
        if operation and selection:
            selection_result = _run_selection_operation(command, operation, selection)
            if selection_result is not None:
                return jsonify(selection_result)

    return jsonify(_error_response_from_copilot(command, result))


# ---------------------------------------------------------------------- #
# Helpers
# ---------------------------------------------------------------------- #


def _build_snapshot(headers, rows, origin_row, origin_col):
    """Map real worksheet cells that the snapshot populated.

    Keys are ``(row, column)`` tuples (1-based), values are the values the
    Add-in sent. Used to diff against the in-memory workbook afterwards.
    """
    origin_row = int(origin_row or 1)
    origin_col = int(origin_col or 1)
    snapshot = {}
    for i, header in enumerate(headers or []):
        snapshot[(origin_row, origin_col + i)] = header
    data_start = origin_row + (1 if headers else 0)
    for r, row in enumerate(rows or []):
        for c, value in enumerate(row):
            if is_blank(value):
                continue
            snapshot[(data_start + r, origin_col + c)] = value
    return snapshot


def _compute_writes(handler, snapshot, charts):
    """Diff the processed in-memory workbook against the snapshot and
    return the machine-readable writes the Add-in must apply to Excel.
    """
    cells = []
    sheets = {}
    ws = handler.workbook[handler.sheet_name]

    for row in ws.iter_rows():
        for cell in row:
            r, c = cell.row, cell.column
            original = snapshot.get((r, c))
            value = cell.value
            if value is None:
                if original is not None:
                    cells.append({"row": r, "column": c, "clear": True})
                continue
            if _same_value(value, original):
                continue
            if isinstance(value, str) and value.startswith("="):
                cells.append({"row": r, "column": c, "formula": value})
            else:
                cells.append({"row": r, "column": c, "value": _serialize(value)})

    for name in handler.workbook.sheetnames:
        if name == handler.sheet_name:
            continue
        values = _sheet_values(handler.workbook[name])
        if values:
            sheets[name] = {"values": values}

    return {"cells": cells, "sheets": sheets, "charts": charts}


def _same_value(a, b):
    """True when two cell values should be treated as unchanged."""
    if isinstance(a, bool) or isinstance(b, bool):
        return a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return float(a) == float(b)
    return a == b


def _serialize(value):
    """Serialize a cell value into a plain JSON-friendly value."""
    if isinstance(value, (bool, int, type(None))):
        return value
    if isinstance(value, float):
        return value
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, str):
        return value
    return str(value)


def _sheet_values(ws):
    """Dump a worksheet as a trimmed 2D list of serialized values."""
    rows_out = []
    for row in ws.iter_rows():
        values = [_serialize(cell.value) for cell in row]
        while values and values[-1] is None:
            values.pop()
        if values:
            rows_out.append(values)
    return rows_out


def _chart_writes(result, origin_row):
    """Build chart descriptors for CHART operations."""
    instruction = result.get("instruction") or {}
    if instruction.get("operation") != "CHART":
        return []
    rinfo = result.get("result") if isinstance(result.get("result"), dict) else {}
    column = rinfo.get("source") or instruction.get("column")
    if not column:
        return []
    line = bool(instruction.get("line_chart"))
    rows_count = result.get("written_count") or rinfo.get("rows") or 1
    source_range = (
        f"{column}{int(origin_row or 1)}:{column}{int(origin_row or 1) + int(rows_count)}"
    )
    return [
        {
            "type": "LineChart" if line else "ColumnClustered",
            "title": rinfo.get("chart") or f"Chart - {column}",
            "source_range": source_range,
        }
    ]


def _is_overwrite_message(message):
    return bool(message) and "already contains data" in message


def _command_needs_selection(message):
    """True when a failed command should retry against the selection."""
    if not message:
        return False
    return bool(
        message.startswith(_DATA_SOURCE_ERROR_PREFIX) or "couldn't find a column named" in message
    )


def _detect_operation(command):
    """Detect an operation keyword using the parser's public keyword table."""
    normalized = _normalize(command)
    for keyword, op in sorted(
        CommandParser.OPERATION_KEYWORDS.items(), key=lambda i: len(i[0]), reverse=True
    ):
        if op is not None and keyword in normalized:
            return op
    return None


def _normalize(command):
    normalized = command.strip().lower()
    normalized = re.sub(r"[?!.,]+", "", normalized)
    return re.sub(r"\s+", " ", normalized)


def _to_number(value):
    """Coerce a value to an int/float, or None when not numeric."""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        text = value.strip()
        try:
            number = float(text)
            return int(number) if number.is_integer() else number
        except ValueError:
            return None
    return None


def _flatten(values):
    flat = []
    for value in values or []:
        if isinstance(value, list):
            flat.extend(_flatten(value))
        else:
            flat.append(value)
    return flat


def _selection_auto_write(selection, result, operation):
    """Suggest a destination cell below a single-column selection."""
    if operation not in (*AGGREGATE_OPERATIONS, "ADD"):
        return None
    if selection.get("column_count") not in (None, 1):
        return None
    start_row = int(selection.get("row") or 1)
    row_count = int(selection.get("row_count") or 1)
    column = int(selection.get("column") or 1)
    return {"row": start_row + row_count, "column": column, "value": result}


def _run_selection_operation(command, operation, selection):
    """Compute an operation against the user's current selection.

    Returns a full response dict (same shape as the copilot results) or
    ``None`` when the selection cannot satisfy the operation.
    """
    numbers = []
    for value in _flatten(selection.get("values")):
        number = _to_number(value)
        if number is not None:
            numbers.append(number)

    if not numbers:
        return _error_response(
            "The selected range has no numeric values. "
            "Select cells that contain numbers and try again."
        )

    total = sum(numbers)
    result = None
    message = None
    auto_write = None

    if operation in ("SUM", "ADD", "AVERAGE", "MIN", "MAX", "COUNT"):
        engine = {
            "SUM": CalculationEngine.sum,
            "AVERAGE": CalculationEngine.average,
            "MIN": CalculationEngine.minimum,
            "MAX": CalculationEngine.maximum,
            "COUNT": CalculationEngine.count,
        }
        result = engine[operation](numbers)
        if operation == "ADD":
            result = total
        auto_write = _selection_auto_write(selection, result, operation)
    elif operation == "MULTIPLY":
        product = 1
        for number in numbers:
            product *= number
        result = product
    elif operation in ("SUBTRACT", "DIVIDE", "DIFFERENCE", "GROWTH"):
        if len(numbers) < 2:
            return _error_response(f"{operation} over the selection needs at least two numbers.")
        a, b = numbers[0], numbers[1]
        try:
            if operation == "SUBTRACT":
                result = CalculationEngine.subtract(a, b)
            elif operation == "DIVIDE":
                result = CalculationEngine.divide(a, b)
            elif operation == "DIFFERENCE":
                result = CalculationEngine.difference(a, b)
            else:
                result = CalculationEngine.growth(a, b)
        except ValueError as e:
            return _error_response(str(e))
    elif operation == "PERCENTAGE":
        if total == 0:
            return _error_response("The selection total is 0, so percentages cannot be calculated.")
        result = [round(value / total * 100, 2) for value in numbers]
    else:
        return _error_response(
            f"'{operation}' is not supported on a selection yet. "
            "Try a column-based command like 'Revenue ka total karo'."
        )

    message = _selection_message(operation, result, selection, auto_write)

    return {
        "success": True,
        "message": message,
        "operation": operation,
        "parameters": {
            "selection_address": selection.get("address"),
            "values": numbers,
        },
        "result": result,
        "destination": None,
        "overwrite_needed": False,
        "writes": {"cells": [], "sheets": {}, "charts": []},
        "selection_auto_write": auto_write,
        "selection_used": True,
    }


def _selection_message(operation, result, selection, auto_write):
    """Build a human-readable confirmation for selection operations."""
    count = selection.get("row_count") or 0
    unit = "cell" if count == 1 else "cells"
    verb = "totals" if count == 1 else "total"

    def fmt(value):
        if isinstance(value, float) and value.is_integer():
            return int(value)
        return value

    prefixes = {
        "SUM": f"Done. The selected {unit} {verb} to {fmt(result)}.",
        "ADD": f"Done. I added the selected values ({selection.get('address')}) = {fmt(result)}.",
        "AVERAGE": f"Done. The average of the selected values is {fmt(result)}.",
        "MIN": f"Done. The minimum of the selected values is {fmt(result)}.",
        "MAX": f"Done. The maximum of the selected values is {fmt(result)}.",
        "COUNT": f"Done. The selection contains {result} values.",
        "MULTIPLY": f"Done. The product of the selected values is {fmt(result)}.",
        "SUBTRACT": f"Done. First minus second selected value = {fmt(result)}.",
        "DIVIDE": f"Done. The division result is {fmt(result)}.",
        "DIFFERENCE": f"Done. The difference between the selected values is {fmt(result)}.",
        "GROWTH": f"Done. The growth of the selected values is {fmt(result)}%.",
    }
    base = prefixes.get(operation)
    if base is None and operation == "PERCENTAGE":
        preview = ", ".join(f"{fmt(p)}%" for p in result[:6])
        base = f"Done. Each selected value as a % of the selection total: {preview}."
        if len(result) > 6:
            base += f" and {len(result) - 6} more."
    base = base or "Done."

    if auto_write:
        ref = f"{get_column_letter(auto_write['column'])}{auto_write['row']}"
        base += f" I can write {fmt(auto_write['value'])} into {ref}."
    return base


def _success_response(command, result, writes):
    return {
        "success": True,
        "message": result["message"],
        "operation": result["instruction"].get("operation"),
        "parameters": result["instruction"],
        "result": result.get("result"),
        "destination": result.get("destination"),
        "written_count": result.get("written_count", 0),
        "overwrite_needed": False,
        "writes": writes,
        "selection_used": False,
    }


def _error_response_from_copilot(command, result):
    return {
        "success": False,
        "message": result["message"],
        "operation": (result.get("instruction") or {}).get("operation"),
        "parameters": result.get("instruction"),
        "result": None,
        "destination": None,
        "overwrite_needed": _is_overwrite_message(result["message"]),
        "writes": {"cells": [], "sheets": {}, "charts": []},
        "selection_used": False,
    }


def _error_response(message, operation=None):
    return {
        "success": False,
        "message": message,
        "operation": operation,
        "parameters": None,
        "result": None,
        "destination": None,
        "overwrite_needed": _is_overwrite_message(message),
        "writes": {"cells": [], "sheets": {}, "charts": []},
        "selection_used": False,
    }
