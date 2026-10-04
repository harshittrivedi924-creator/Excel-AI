# REST API Documentation

Base URL: `http://localhost:5000`

## Endpoints

### GET /api/workbook

Get information about the currently loaded workbook.

**Response:**
```json
{
  "loaded": true,
  "sheets": ["Sheet1", "Sales"],
  "active_sheet": "Sheet1",
  "rows": [
    ["Month", "Revenue", "Expense"],
    ["January", "10000", "8000"],
    ["February", "12000", "9000"]
  ],
  "file": "sales.xlsx"
}
```

**Not loaded:**
```json
{
  "loaded": false,
  "sheets": [],
  "rows": [],
  "file": null
}
```

### POST /api/upload

Upload an Excel workbook.

**Request:**
- Content-Type: `multipart/form-data`
- Body: `file` field with .xlsx/.xlsm/.xls file

**Response:** Same as `GET /api/workbook`

**Error:**
```json
{
  "loaded": false,
  "error": "Couldn't open file: ..."
}
```

### POST /api/command

Execute a natural language command.

**Request:**
```json
{
  "command": "Column D ka sum karo",
  "sheet": "Sheet1",
  "allow_overwrite": false
}
```

**Response (success):**
```json
{
  "success": true,
  "message": "Done. The sum of D2:D20 is 54000.",
  "result": 54000,
  "overwrite_needed": false
}
```

**Response (overwrite needed):**
```json
{
  "success": false,
  "message": "D21 already contains data. Replace it?",
  "result": null,
  "overwrite_needed": true
}
```

**Response (error):**
```json
{
  "success": false,
  "message": "I couldn't find Column X in this sheet.",
  "result": null,
  "overwrite_needed": false
}
```

### GET /api/history

Get the command history.

**Response:**
```json
[
  {
    "operation": "SUM",
    "result": 54000,
    "destination": "D21"
  },
  {
    "operation": "AVERAGE",
    "result": 3000,
    "destination": null
  }
]
```

## Error Handling

All endpoints return appropriate HTTP status codes:
- `200` - Success
- `400` - Bad request (missing file, invalid command)
- `413` - Analysis request too large (over 4MB)
- `500` - Server error

## Excel Add-in Endpoints

The `/api/excel/*` endpoints are used only by the Excel Add-in task pane. They
are registered as a Flask blueprint (`backend/api/addin_routes.py`) and are
additive — the `/api/*` endpoints, the CLI and the Web UI are unaffected.
See [EXCEL_ADDIN.md](EXCEL_ADDIN.md) for the full task-pane flow.

### GET /api/excel/health

Liveness probe used by the task pane to show its connection status.

**Response:**
```json
{ "status": "ok", "service": "excel-ai-copilot-addin-api", "version": "1.0.0" }
```

### POST /api/excel/command

Runs a natural-language command against a snapshot of the active worksheet and
returns a machine-readable `writes` descriptor for the Add-in to apply.

**Request:** (same snapshot the task pane already captures)
```json
{
  "command": "Revenue ka total karo",
  "sheet_name": "Sheet1",
  "origin": { "row": 1, "column": 1 },
  "headers": ["Month", "Revenue"],
  "rows": [["Jan", 1000], ["Feb", 1200]],
  "selection": { "address": "B2:B3", "values": [[1000], [1200]] },
  "allow_overwrite": false
}
```

**Response:**
```json
{
  "success": true,
  "message": "Revenue total = 2200.",
  "operation": "SUM",
  "writes": { "cells": [], "sheets": {}, "charts": [] }
}
```

`writes.charts` entries look like
`{ "type": "ColumnClustered", "title": "...", "source_range": "B1:B3" }`.

### POST /api/excel/analyze

Read-only data profiling and analysis (Phase 4). The request body is the same
snapshot structure as `/command`, so the task pane reuses its existing capture
logic verbatim. **This endpoint never modifies the workbook** — `writes` is
always empty and charts are a *recommendation* only.

**Request:**
```json
{
  "sheet_name": "Sheet1",
  "origin": { "row": 1, "column": 1 },
  "headers": ["Month", "Revenue", "Expense"],
  "rows": [
    ["Jan", 1000, 600], ["Feb", 1200, 650], ["Mar", 1800, 900],
    ["Apr", 1600, 800], ["May", 2000, 950]
  ],
  "selection": { "address": "A1:C6", "values": [["Jan", 1000, 600]] }
}
```

Passing `selection` narrows the analysis to the highlighted range. An empty or
absent selection is not an error — the analyzer falls back to the used range
and reports `"scope": { "source": "worksheet" }`.

**Response (abbreviated):**
```json
{
  "success": true,
  "message": "Analyzed 5 row(s) and 3 column(s), including 2 numeric column(s).",
  "scope": { "source": "selection", "address": "A1:C6", "row_count": 5, "column_count": 3 },
  "profile": {
    "rows": 5, "columns": 3,
    "column_names": ["Month", "Revenue", "Expense"],
    "numeric_columns": ["Revenue", "Expense"],
    "text_columns": ["Month"], "date_columns": ["Month"],
    "empty_cells": 0, "duplicate_rows": 0, "completeness_pct": 100
  },
  "statistics":  [{ "column": "Revenue", "type": "numeric", "count": 5, "sum": 7600, "average": 1520, "min": 1000, "max": 2000, "median": 1600, "range": 1000, "stdev": 414.7288 }],
  "growth":      [{ "column": "Revenue", "available": true, "first": 1000, "last": 2000, "first_label": "Jan", "last_label": "May", "change": 1000, "percent": 100 }],
  "trends":      [{ "column": "Revenue", "direction": "increasing", "rises": 3, "falls": 1, "flat": 0, "net_change": 1000, "detail": "Revenue moved 3 step(s) up and 1 step(s) down across 5 values (net +1000). Largest change +600 at Mar." }],
  "anomalies":   [{ "column": "Revenue", "count": 0, "method": "IQR + z-score", "note": "No values in Revenue fall outside 1.5x IQR or 2 standard deviations." }],
  "period_extremes": [{ "column": "Revenue", "label_column": "Month", "highest": { "label": "May", "where": "May", "row_index": 4, "value": 2000 }, "lowest": { "label": "Jan", "where": "Jan", "row_index": 0, "value": 1000 } }],
  "insights":    [{ "text": "May was the peak Revenue period with 2000.", "evidence": "max single value of Revenue = 2000 at May" }],
  "chart_recommendation": {
    "available": true, "type": "LineChart", "title": "Revenue vs Expense by Month",
    "columns": ["Revenue", "Expense"], "label_column": "Month", "row_count": 5,
    "reason": "'Revenue vs Expense by Month' compares 2 numeric column(s) across Month. Values change over time, so a line chart reads best."
  },
  "data_quality": { "empty_cells": 0, "total_cells": 15, "completeness_pct": 100, "duplicate_rows": 0, "mixed_type_columns": [] },
  "limitations": ["Statistics describe only the cells included in this snapshot."],
  "truncated": false,
  "writes": { "cells": [], "sheets": {}, "charts": [] }
}
```

Notes on the shape above:

- A bare month name such as `Jan` is typed as a **date**, so `Month` appears in
  both `text_columns` and `date_columns`.
- `statistics` entries for a non-numeric column omit `median`, `range` and
  `stdev` (they are `null`), so clients must treat those keys as optional.
- `chart_recommendation` is a **recommendation only**. The existing pipeline
  charts `columns[0]`; a chart follow-up says so and you add further series with
  normal chart commands.
- `writes` is always empty on this endpoint — analysis never modifies the
  workbook.

#### Follow-up questions

Pass a `question` plus the previous `analysis` report to ask a follow-up
without re-analysing. No snapshot is needed in that case.

```json
{ "question": "which month was highest?", "analysis": { "...": "previous report" } }
```

**Response:**
```json
{
  "success": true,
  "handled": true,
  "followup": { "intent": "highest", "message": "May was the highest Revenue period with 2000 ..." },
  "message": "May was the highest Revenue period with 2000 ...",
  "chart_command": null,
  "writes": { "cells": [], "sheets": {}, "charts": [] }
}
```

Supported `intent` values: `chart`, `compare`, `highest`, `lowest`, `growth`,
`anomalies`, `trend`, `quality`, `summary`, `unknown`.

`chart_command` is populated only for a chart follow-up; the task pane sends
it back through `/api/excel/command` so the existing chart pipeline creates
the real chart in Excel.

#### Analysis limits

| Limit | Value | Applies to |
|---|---|---|
| `MAX_ANALYSIS_BODY` | 4 MB (returns `413`) | whole request body |
| `MAX_ANALYSIS_ROWS` | 5000 | `headers` + `rows` |
| `MAX_ANALYSIS_COLS` | 100 | `headers` + `rows` |
| `MAX_ANALYSIS_CELLS` | 200000 | `headers` + `rows` |

When a cap is hit, `"truncated": true` is set, the analyzed row/column counts
are still reported, and a limitation explains that results cover the snapshot
portion only.

Two things the table does not imply:

- `MAX_ANALYSIS_CELLS` is checked after the column cap, so with 100 columns the
  cell budget (200000) is reached first and cuts a 5000-row payload to ~2000
  rows. The two limits are not independent.
- These caps are applied to `headers`/`rows` only. `selection.values` is passed
  through as supplied, so a direct API caller can analyse a larger selection
  than the table suggests. The Excel add-in is unaffected — it caps the
  snapshot it sends at `MAX_ROWS = 3000` rows / `MAX_COLS = 500` columns
  before posting (`excel-addin/src/taskpane/taskpane.js`).

## CORS

The web UI is served from the same origin as the API, so no CORS headers are
needed for it.

The Excel Add-in task pane runs from a different origin
(`https://localhost:3000`), so CORS is enabled for the `excel_addin` blueprint
only. Override the allowed origins with the
`EXCEL_ADDIN_ALLOWED_ORIGINS` environment variable (comma-separated); it
defaults to `https://localhost:3000`.
