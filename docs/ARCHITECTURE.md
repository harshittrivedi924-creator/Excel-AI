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

## Data Flow

```
1. User types/speaks command
2. Parser extracts: operation, operands, destination
3. Validator checks: column exists, data is numeric, overwrite OK
4. Calculator computes result (or Excel handler writes formula)
5. Copilot builds confirmation message
6. Result returned to user (CLI/Voice/Web)
```

## File Structure

```
Excel-AI/
├── backend/
│   ├── api/
│   │   ├── app.py              # Flask web API
│   │   └── uploads/            # Uploaded workbooks
│   ├── calculator/
│   │   └── engine.py           # Pure calculation functions
│   ├── excel/
│   │   └── handler.py          # Excel read/write operations
│   ├── parser/
│   │   └── parser.py           # NL command parser
│   ├── validator/
│   │   └── validator.py        # Input validation
│   ├── voice/
│   │   └── listener.py         # Speech-to-text
│   └── copilot.py              # Main orchestrator
├── frontend/
│   ├── index.html              # Web UI markup
│   ├── style.css               # Web UI styles
│   └── app.js                  # Web UI logic
├── tests/
│   ├── calculator/             # Calculator engine tests
│   ├── excel/                  # Excel handler tests
│   ├── parser/                 # Parser tests
│   ├── test_advanced.py        # Advanced feature tests
│   ├── test_conditional.py     # Conditional operation tests
│   ├── test_data_ops.py        # Data operation tests
│   ├── test_integration.py     # End-to-end tests
│   ├── test_voice.py           # Voice module tests
│   └── test_web.py             # Web API tests
├── docs/
│   ├── ARCHITECTURE.md         # This file
│   ├── COMMANDS.md             # Command reference
│   └── API.md                  # REST API docs
├── main.py                     # CLI entry point
├── requirements.txt            # Python dependencies
└── README.md                   # Project overview
```
