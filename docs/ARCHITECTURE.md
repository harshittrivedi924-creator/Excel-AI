# Architecture

## Overview

Excel AI Copilot is a natural language spreadsheet automation agent that converts spoken or typed commands into Excel operations. The system follows a pipeline architecture:

```
User Input → Parser → Validator → Calculator/Excel Handler → Output
```

## System Components

### 1. Command Parser (`backend/parser/parser.py`)
Converts natural language commands into structured instructions. Supports:
- English and Hinglish commands
- Column references (e.g., "Column D", "D")
- Named columns (e.g., "Revenue", "Sales")
- Cell references (e.g., "B2", "C3:D10")
- Conditional clauses (e.g., "jahan Revenue 1000 se zyada")

### 2. Validator (`backend/validator/validator.py`)
Validates parsed instructions before execution:
- Checks operation support
- Validates column/range existence
- Handles overwrite confirmation
- Ensures data types are correct

### 3. Calculator Engine (`backend/calculator/engine.py`)
Pure calculation functions with no side effects:
- Aggregates: SUM, AVERAGE, MIN, MAX, COUNT
- Arithmetic: ADD, SUBTRACT, MULTIPLY, DIVIDE
- Advanced: DIFFERENCE, GROWTH, PERCENTAGE
- Conditional: SUMIF, COUNTIF, AVERAGEIF
- Statistics: Standard deviation for anomaly detection

### 4. Excel Handler (`backend/excel/handler.py`)
Reads and writes Excel workbooks using openpyxl:
- Load/save workbooks
- Read column values, ranges, cells
- Write formulas and values
- Data operations (sort, dedupe, filter)
- Data cleaning (standardize dates/names, detect invalid)
- Chart creation (Bar, Line)
- Workbook analysis with trend/anomaly detection

### 5. Copilot Orchestrator (`backend/copilot.py`)
Main entry point that ties all components together:
- Routes commands to appropriate handlers
- Manages calculation history
- Handles interactive mode and voice mode
- Builds user-friendly confirmation messages

### 6. Voice Module (`backend/voice/listener.py`)
Speech-to-text conversion using Google Web Speech API:
- Microphone input
- Ambient noise adjustment
- Configurable language (default: en-IN for Hinglish)

### 7. Web API (`backend/api/app.py`)
Flask-based REST API for the web UI:
- `GET /api/workbook` - Workbook info and preview
- `POST /api/upload` - Upload Excel file
- `POST /api/command` - Execute natural language command
- `GET /api/history` - Command history

### 8. Excel Add-in API (`backend/api/addin_routes.py`)
Flask blueprint serving the Office.js task pane, mounted at `/api/excel`:
- `GET /health` - Liveness probe for the task-pane status pill
- `POST /command` - Runs a command against a worksheet snapshot
- `POST /analyze` - Read-only profiling/statistics/insights (Phase 4)

The snapshot is materialised into an in-memory openpyxl workbook
(`backend/excel/memory_handler.py`) rather than a file, so the backend never
touches the user's Excel instance or disk. `POST /command` diffs the processed
workbook against the original snapshot and returns a machine-readable `writes`
descriptor that the task pane applies back through Office.js.

### 9. Data Analyzer (`backend/excel/data_analyzer.py`)
Phase 4's read-only analysis layer, behind `POST /api/excel/analyze`. Takes a
snapshot (headers + rows, narrowed to the selection when one is present) and
returns a JSON-friendly report: profile, per-column statistics, growth, trends,
anomalies, period extremes, data quality, ranked insights and a chart
recommendation. It reuses `CalculationEngine` and `is_blank`, imports nothing
outside the standard library, and never mutates a workbook or creates a chart.

`backend/excel/analysis_followup.py` answers follow-up questions
("which month was highest?", "compare revenue and expenses") from the *already
returned* report, so a follow-up costs no new analysis and always agrees with
the numbers the user just saw. Unknown questions return `handled: false` and the
task pane falls back to the regular command pipeline.

## Data Flow

```
1. User types/speaks command
2. Parser extracts: operation, operands, destination
3. Validator checks: column exists, data is numeric, overwrite OK
4. Calculator computes result (or Excel handler writes formula)
5. Copilot builds confirmation message
6. Result returned to user (CLI/Voice/Web)
```

### Add-in command flow (Phases 2-3)

```
Excel -> Office.js task pane -> snapshot -> POST /api/excel/command
      -> in-memory workbook -> parser/validator/calculator/copilot
      -> `writes` descriptor -> Office.js -> live Excel
```

### Add-in analysis flow (Phase 4)

```
Excel selection -> Office.js snapshot -> POST /api/excel/analyze
      -> DataAnalyzer (profile -> statistics -> trend -> anomaly -> insights
                        -> chart recommendation)
      -> structured JSON -> collapsible analysis card in the task pane
      -> "Create Chart" -> existing /command chart pipeline -> Office.js -> Excel
```

## Design Constraints

- **No external AI service.** The natural-language layer is deterministic
  regex/keyword parsing and the analyzer is pure arithmetic. The project runs
  fully offline with no API key, and the interfaces leave room for an LLM layer
  to be added later without changing the surrounding pipeline.
- **Analysis is read-only.** Only an explicit write command or "Create Chart"
  modifies the workbook.
- **Honest statistics.** Anything needing a minimum sample size is reported as
  unavailable with a reason rather than guessed, and every insight carries the
  figures it was derived from.

## File Structure

```
Excel-AI/
├── backend/
│   ├── api/
│   │   ├── app.py              # Flask web API
│   │   ├── addin_routes.py     # /api/excel/* add-in endpoints
│   │   └── uploads/            # Uploaded workbooks
│   ├── calculator/
│   │   └── engine.py           # Pure calculation functions
│   ├── excel/
│   │   ├── handler.py          # Excel read/write operations
│   │   ├── memory_handler.py   # In-memory workbook for add-in snapshots
│   │   ├── data_analyzer.py    # Read-only profiling & analysis
│   │   └── analysis_followup.py # Deterministic follow-up answers
│   ├── parser/
│   │   └── parser.py           # NL command parser
│   ├── validator/
│   │   └── validator.py        # Input validation
│   ├── voice/
│   │   └── listener.py         # Speech-to-text
│   └── copilot.py              # Main orchestrator
├── excel-addin/                # Office.js add-in
│   ├── manifest.xml
│   ├── src/taskpane/           # taskpane.{html,css,js}, format.mjs, snapshot.mjs
│   └── test/                   # node --test unit tests
├── frontend/
│   ├── index.html              # Web UI markup
│   ├── style.css               # Web UI styles
│   └── app.js                  # Web UI logic
├── tests/
│   ├── calculator/             # Calculator engine tests
│   ├── excel/                  # Excel handler tests
│   ├── parser/                 # Parser tests
│   ├── test_addin_routes.py    # Add-in API contract tests
│   ├── test_advanced.py        # Advanced feature tests
│   ├── test_conditional.py     # Conditional operation tests
│   ├── test_data_analyzer.py   # Data analyzer + analyze endpoint tests
│   ├── test_data_ops.py        # Data operation tests
│   ├── test_integration.py     # End-to-end tests
│   ├── test_voice.py           # Voice module tests
│   └── test_web.py             # Web API tests
├── docs/
│   ├── ARCHITECTURE.md         # This file
│   ├── COMMANDS.md             # Command reference
│   ├── API.md                  # REST API docs
│   └── EXCEL_ADDIN.md          # Add-in setup, contract, manual tests
├── main.py                     # CLI entry point
├── requirements.txt            # Python dependencies
└── README.md                   # Project overview
```
