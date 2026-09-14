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

See [docs/COMMANDS.md](docs/COMMANDS.md) for the full command reference.

## Project Structure

```
Excel-AI/
├── backend/
│   ├── api/              # Flask web API
│   ├── calculator/       # Calculation engine
│   ├── excel/            # Excel read/write
│   ├── parser/           # NL command parser
│   ├── validator/        # Input validation
│   ├── voice/            # Speech-to-text
│   └── copilot.py        # Main orchestrator
├── frontend/             # Web UI (HTML/CSS/JS)
├── tests/                # 129 unit & integration tests
├── docs/                 # Architecture, commands, API docs
├── main.py               # CLI entry point
├── requirements.txt
├── Dockerfile
├── docker-compose.yml
└── .github/workflows/    # CI pipeline
```

## Testing

```bash
# Run all 129 tests
pytest tests/ -v

# Run a specific test file
pytest tests/test_integration.py -v
```

## Documentation

- [Architecture](docs/ARCHITECTURE.md) — System design & data flow
- [Command Reference](docs/COMMANDS.md) — All commands & Hinglish keywords
- [REST API](docs/API.md) — Web API endpoints

## License

MIT License
