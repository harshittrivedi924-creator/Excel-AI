# Excel AI Copilot

Natural Language Spreadsheet Automation Agent

> Type or speak commands in English or Hinglish, and the copilot handles the rest — calculations, data cleaning, charts, and analysis.

## Features

- **Natural language commands** for Excel operations
- **Calculations**: SUM, AVERAGE, MIN, MAX, COUNT, +, −, ×, ÷
- **Advanced**: DIFFERENCE, GROWTH (%), PERCENTAGE
- **Named-column lookups** ("Revenue ka total karo", "Expense ka average")
- **Element-wise formulas** ("Revenue minus expense karke profit nikalo")
- **Data operations**: sort, remove duplicates, filter, find empty cells
- **Data cleaning**: standardize dates/names, detect invalid values
- **Conditional formulas**: SUMIF, COUNTIF, AVERAGEIF
- **Chart generation** (Bar & Line) — "Sales ka chart bana do"
- **Workbook analysis** with trend & anomaly detection
- **Overwrite confirmation** before replacing existing cells
- **Hinglish support** with voice input
- **Web UI** with Excel-file upload and chat sidebar

### AI Data Analyst (Excel Add-in)

Select a range in Excel and ask the Copilot to analyze it. Fully local and
deterministic — no API key, no external service.

- **Data profiling** — rows, columns, column names, numeric/text/date columns,
  empty cells, duplicate rows, completeness
- **Statistics** — sum, average, min, max, median, range and standard deviation
  are returned by the API; the analysis card shows total, average, min and max
- **Growth** — first → last value, absolute and percentage change
- **Trend detection** — increasing / decreasing / flat / mixed (or
  `insufficient_data`), plus rises, falls and net change
- **Anomaly detection** — IQR (1.5×) and z-score (>2σ), with the reason
- **Key insights** — up to 5, each carrying the figures it came from
- **Chart recommendation** — line or column, with a one-click **Create Chart**
  (charts the first recommended column; add others with a normal chart command)
- **Follow-up questions** — "which month was highest?", "show growth",
  "any unusual values?", "compare revenue and expenses"

## Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/Excel-AI.git
cd Excel-AI

# Create virtual environment
python -m venv venv
.\venv\Scripts\activate   # Windows
# source venv/bin/activate  # macOS/Linux

# Install dependencies
pip install -r requirements.txt
```

## Quick Start

### Command Line

```bash
# Run one command
python main.py workbook.xlsx -c "Column D ka total karo aur D21 mein daal do"

# Interactive mode
python main.py workbook.xlsx

# Voice mode
python main.py workbook.xlsx --voice

# Use a specific sheet
python main.py workbook.xlsx --sheet "Sales Data" -c "Revenue ka average karo"
```

### Web UI

```bash
python main.py --serve
# Then open http://localhost:5000

# Or with a pre-loaded workbook
python main.py sales.xlsx --serve
```

### Docker

```bash
docker compose up --build
# Open http://localhost:5000
```

## Example Commands

| Category | Example |
|----------|---------|
| Sum a column | `Column D ka sum karo` |
| Named column | `Revenue ka total karo` |
| With destination | `Column D ka total karo aur D21 mein daal do` |
| Arithmetic | `B2 + C2` |
| Element-wise | `Revenue minus expense karke profit nikalo` |
| Growth | `February vs January kitni badhi` |
| Conditional | `B ka sum jahan A Pen hai` |
| Sort | `Column B sort karo` |
| Dedupe | `Duplicate rows hata do` |
| Filter | `Filter data jahan Revenue 100 se zyada` |
| Cleaning | `Dates ko standardize karo` |
| Chart | `Sales ka chart bana do` |
| Analysis | `Is workbook ka analysis kar do` |
| Data analysis | `Analyze this data`, `Is data mein kya important hai?` |
| Analysis follow-up | `Which month had the highest revenue?`, `Compare revenue and expenses` |

See [docs/COMMANDS.md](docs/COMMANDS.md) for the full command reference.

## Project Structure

```
Excel-AI/
├── backend/
│   ├── api/              # Flask web API + Excel Add-in API
│   ├── calculator/       # Calculation engine
│   ├── excel/            # Excel read/write + data analyzer
│   ├── parser/           # NL command parser
│   ├── validator/        # Input validation
│   ├── voice/            # Speech-to-text
│   └── copilot.py        # Main orchestrator
├── excel-addin/          # Office.js add-in (manifest, task pane, build)
├── frontend/             # Web UI (HTML/CSS/JS)
├── tests/                # 281 unit & integration tests
├── docs/                 # Architecture, commands, API, add-in
├── main.py               # CLI entry point
├── requirements.txt
├── Dockerfile
├── docker-compose.yml
└── .github/workflows/    # CI pipeline
```

## Testing

```bash
# Run all 281 Python tests
pytest tests/ -v

# Run a specific test file
pytest tests/test_integration.py -v

# Excel Add-in unit tests
cd excel-addin && npm test
```

## Documentation

- [Architecture](docs/ARCHITECTURE.md) — System design & data flow
- [Command Reference](docs/COMMANDS.md) — All commands & Hinglish keywords
- [REST API](docs/API.md) — Web API and Excel Add-in endpoints
- [Excel Add-in](docs/EXCEL_ADDIN.md) — Sideloading, API contract, manual tests

## License

MIT License
