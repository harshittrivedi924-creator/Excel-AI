# Excel AI Copilot

Natural Language Spreadsheet Automation Agent

## Features

- Natural language commands for Excel operations
- Support for calculations (SUM, AVERAGE, MIN, MAX, COUNT)
- Basic arithmetic operations (+, -, *, /)
- Advanced calcs: DIFFERENCE, GROWTH (%), PERCENTAGE
- Named-column lookups ("Revenue ka total karo", "Expense ka average")
- Element-wise formulas ("Revenue minus expense karke profit nikalo")
- Overwrite confirmation before replacing existing cells
- Hinglish support
- Voice command support (planned)

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

```bash
# Run the application
python main.py
```

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