# Excel AI Copilot

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![Flask](https://img.shields.io/badge/Flask-2.3+-green.svg)](https://flask.palletsprojects.com/)
[![JavaScript](https://img.shields.io/badge/JavaScript-ES6+-yellow.svg)](https://developer.mozilla.org/en-US/docs/Web/JavaScript)
[![Office.js](https://img.shields.io/badge/Office.js-Add--in-blue.svg)](https://learn.microsoft.com/en-us/office/dev/add-ins/)
[![Tests](https://img.shields.io/badge/Tests-322%20Passing-brightgreen.svg)](#testing)

**AI-powered Excel automation and data analysis using natural-language and Hinglish commands.**

## Project Overview

Excel AI Copilot is a smart Excel assistant that lets you control spreadsheets using natural language commands in English or Hinglish. Built to automate repetitive Excel workflows and explore practical AI-assisted productivity tooling.

## Architecture

`
Excel → Office.js Task Pane → Flask API → Parser/Validator → Excel/Data Analysis → Structured Response → Excel
`

## Tech Stack

| Layer | Technologies |
|---|---|
| Backend | Python, Flask |
| Frontend/Add-in | JavaScript (ES6+), Office.js |
| Build Tools | Webpack |
| Testing | pytest (Python), Node.js test runner (JavaScript) |
| Code Quality | Ruff |
| Platform | Microsoft Excel (Windows, Mac, Web) |

## Example Commands

`	ext
b ka total karo                    # Calculate sum of column B
inka total karo                    # Calculate total of current selection
aor b ka sum kar ke d mein daal do # Sum A and B, put result in D
Analyze this data                  # Run full AI Data Analyst
Sales ka chart bana do            # Create chart for Sales column
Revenue ka chart bana do          # Create chart for Revenue column
Expense ka chart bana do          # Create chart for Expense column
`

## Phase 4 AI Data Analyst

The AI Data Analyst module provides comprehensive data profiling and insights:

- **Statistics**: Computes total, average, min, max, median, range, and standard deviation for numeric columns
- **Trend Detection**: Identifies increasing, decreasing, flat, or mixed trends with detailed direction analysis
- **Growth Analysis**: Calculates absolute and percentage changes between periods
- **Anomaly Detection**: Flags outliers using IQR (1.5× IQR) and Z-score (>2σ) methods
- **Data Quality**: Reports empty cells, duplicate rows, completeness percentage, and mixed-type columns
- **Chart Recommendations**: Suggests optimal chart types (line/column) based on data characteristics
- **Insights**: Generates up to 5 actionable insights synthesizing all analysis results

## Testing

All tests pass with comprehensive coverage:

- **281 Python tests passing** (backend functionality, parser, analyzer, API routes, integrations)
- **41 JavaScript tests passing** (add-in utilities, formatting, snapshot logic)
- **Ruff checks passing** (linting and code style)
- **Ruff format checks passing** (code formatting)
- **Webpack build successful** (add-in bundles correctly)
- **Manifest validation successful** (Office Add-in manifest valid)

## Project Structure

`	ext
Excel-AI/
├── backend/              # Flask backend
│   ├── api/              # API routes
│   ├── calculator/       # Calculation engine
│   ├── excel/            # Excel operations, data analyzer, analysis followup
│   ├── parser/           # Natural language parser (English/Hinglish)
│   ├── validator/        # Input validation
│   └── copilot.py        # Core copilot logic
├── excel-addin/          # Microsoft Excel Office.js Add-in
│   ├── src/taskpane/     # Task pane UI and logic
│   ├── manifest.xml      # Office Add-in manifest
│   └── package.json      # Dependencies
├── tests/                # Python test suite
└── docs/                 # Documentation
`

## Installation & Setup

### Prerequisites
- Python 3.8+
- Node.js 16+ and npm
- Microsoft Excel

### Backend Setup (Windows)

`ash
# Clone the repository
git clone https://github.com/harshittrivedi924-creator/Excel-AI.git
cd Excel-AI

# Create and activate virtual environment
python -m venv venv
venv\\Scripts\\activate

# Install dependencies
pip install -r requirements.txt
`

### Excel Add-in Setup

`ash
cd excel-addin
npm install
npm run build
`

### Running the Project

**Flask Backend:**
`ash
cd Excel-AI
venv\\Scripts\\activate
python main.py --serve
`

**Excel Add-in Development:**
`ash
cd excel-addin
npm run dev-server
`

## Why I Built This

This is a personal project built to automate repetitive Excel workflows and explore practical AI-assisted productivity tooling. I wanted to create a solution that understands natural language (especially Hinglish, commonly used in India) to perform complex Excel operations without requiring users to learn formulas or shortcuts. The focus was on building a deterministic, privacy-first tool that works entirely locally.

## Future Improvements

- Support for more chart types (pie, scatter, area)
- Advanced pivot table operations
- Enhanced voice input capabilities
- Batch processing of multiple operations
- Extended template support

## Author

**Harshit Trivedi**
Personal project - Excel AI Copilot
