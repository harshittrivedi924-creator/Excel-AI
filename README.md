# Excel AI Copilot

Natural Language Spreadsheet Automation Agent

## Features

- Natural language commands for Excel operations
- Support for calculations (SUM, AVERAGE, MIN, MAX, COUNT)
- Basic arithmetic operations (+, -, *, /)
- Advanced calcs: DIFFERENCE, GROWTH (%), PERCENTAGE
- Named-column lookups ("Revenue ka total karo", "Expense ka average")
- Element-wise formulas ("Revenue minus expense karke profit nikalo")
- Data operations: sort, remove duplicates, filter, find empty cells
- Chart generation ("Sales ka chart bana do")
- Workbook analysis report ("Is workbook ka analysis kar do")
- Overwrite confirmation before replacing existing cells
- Hinglish support
- Voice command support
- Web UI (Excel-file upload + chat sidebar)

## Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/Excel-AI.git
cd Excel-AI

# Create virtual environment
python -m venv venv
.\venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Usage

### Command line

```bash
# Run one command
python main.py workbook.xlsx -c "Column D ka total karo aur D21 mein daal do"

# Interactive mode
python main.py workbook.xlsx

# Voice mode
python main.py workbook.xlsx --voice
```

### Web UI

```bash
python main.py --serve
# Then open http://localhost:5000
```

Upload an Excel file, then chat with the sidebar. You can also pass a
pre-loaded workbook: `python main.py sales.xlsx --serve`.

## Project Structure

```
Excel-AI/
├── backend/
├── frontend/
├── tests/
├── docs/
├── .gitignore
└── README.md
```

## License

MIT License