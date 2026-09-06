import re
from openpyxl.utils import column_index_from_string


class ParserError(Exception):
    """Exception for parser failures."""

    def __init__(self, message):
        self.message = message
        super().__init__(message)


class CommandParser:
    """Convert natural language commands into structured instructions."""

    OPERATION_KEYWORDS = {
        "sum": "SUM",
        "total": "SUM",
        "sigma": "SUM",
        "average": "AVERAGE",
        "avg": "AVERAGE",
        "mean": "AVERAGE",
        "minimum": "MIN",
        "min": "MIN",
        "maximum": "MAX",
        "max": "MAX",
        "count": "COUNT",
        "add": "ADD",
        "plus": "ADD",
        "add karo": "ADD",
        "add(ake)": "ADD",
        "subtract": "SUBTRACT",
        "minus": "SUBTRACT",
        "difference": "SUBTRACT",
        "subtract karo": "SUBTRACT",
        "multiply": "MULTIPLY",
        "times": "MULTIPLY",
        "product": "MULTIPLY",
        "multiply karo": "MULTIPLY",
        "guna": "MULTIPLY",
        "divide": "DIVIDE",
        "div": "DIVIDE",
        "percentage": "PERCENTAGE",
        "percent": "PERCENTAGE",
        "persent": "PERCENTAGE",
        "calculate": None,
    }

    HINGLISH_KEYWORDS = {
        "karo": "execute",
        "nikal": "calculate",
        "nikal do": "calculate",
        "batao": "report",
        "kar do": "execute",
        "daal do": "write",
        "mein": "to",
        "me": "to",
        "ka": "of",
        "ka total": "sum",
        "ka sum": "sum",
    }

    def parse(self, command):
        """Parse a natural language command into a structured instruction."""
        command = command.strip().lower()

        # Remove common punctuation
        command = re.sub(r"[?!.,]+", "", command)

        # Check for a cell reference in the form like B2
        cell_pattern = r"\b([a-z])\s*(\d+)\b"
        cells = re.findall(cell_pattern, command)

        # Check for column reference like "Column D" or just "D"
        column_pattern = r"\bcolumn\s+([a-z])\b"
        column_matches = re.findall(column_pattern, command)

        # Also look for bare column letters (single letters that aren't
        # part of other words, commonly D, E, A, B, etc.)
        bare_columns = re.findall(r"\b([a-hj-z])\b", command)
        # Exclude common words that are single letters

        # Determine operation
        operation = self._detect_operation(command)

        if operation is None:
            raise ParserError(
                "I couldn't understand the command. Please try using words like "
                "sum, total, average, min, max, count, add, subtract, multiply, divide."
            )

        # Handle arithmetic like B2 + C2
        if operation in ("ADD", "SUBTRACT", "MULTIPLY", "DIVIDE"):
            return self._parse_arithmetic(command, operation, column_matches)

        # Handle column-based calculations
        if column_matches:
            column = column_matches[0].upper()
        elif bare_columns:
            # Heuristic: pick the column that appears in context of operation
            column = bare_columns[0].upper()
        else:
            column = None

        # Handle SUM/AVERAGE etc on a column
        if column:
            # Check for destination cell
            destination = self._find_destination(command)
            source = column
            return {
                "operation": operation,
                "source": source,
                "column": source,
                "destination": destination,
            }

        # Try cell arithmetic like B2 + C2 with result destination D2
        if cells:
            return self._parse_cell_references(command, operation, cells)

        raise ParserError(
            "I couldn't figure out which data to use. Please specify a column "
            "(like Column D) or cells (like B2 and C2)."
        )

    def _detect_operation(self, command):
        """Detect the operation type from a command."""
        # Order matters: check multi-word commands first
        for keyword, op in sorted(
            self.OPERATION_KEYWORDS.items(), key=lambda x: len(x[0]), reverse=True
        ):
            if keyword in command and op is not None:
                return op

        # Check for arithmetic operators
        if "+" in command or "plus" in command:
            return "ADD"
        if "-" in command or ("minus" in command):
            return "SUBTRACT"
        if "*" in command or "times" in command:
            return "MULTIPLY"
        if "/" in command:
            return "DIVIDE"

        return None

    def _find_destination(self, command):
        """Find the destination cell in the command (e.g., D21).
        A cell is only treated as a destination if it follows or is followed
        by a marker word like 'daal do', 'mein', 'me', 'into', 'in', 'at'.
        """
        marker_patterns = [
            r"(?:daal do|dal do|mein|me|into|at)\s*([a-z])(\d+)\b",
            r"([a-z])(\d+)\s*(?:mein|me|daal do|dal do|ko)",
        ]

        for pattern in marker_patterns:
            for match in re.finditer(pattern, command):
                letter = match.group(1)
                digits = match.group(2)
                return f"{letter.upper()}{digits}"

        return None

    def _parse_arithmetic(self, command, operation, column_matches):
        """Parse arithmetic operations like B2 + C2."""
        # Find cell references
        cells = re.findall(r"\b([a-z])(\d+)\b", command)
        cell_refs = [f"{c[0].upper()}{c[1]}" for c in cells]

        # Also check for columns
        if column_matches:
            cell_refs = [c.upper() for c in column_matches]
            destination = self._find_destination(command)
        else:
            # Find destination first (marker-based), then treat the
            # remaining cell refs as inputs.
            destination = self._find_destination(command)
            if destination and destination in cell_refs:
                cell_refs.remove(destination)

        if len(cell_refs) < 2:
            raise ParserError(
                "I need two cell references for arithmetic operations, "
                "like 'B2 + C2'."
            )

        return {
            "operation": operation,
            "inputs": cell_refs[:2],
            "output": destination,
        }

    def _parse_cell_references(self, command, operation, cells):
        """Parse commands with explicit cell references."""
        cell_refs = [f"{c[0].upper()}{c[1]}" for c in cells]
        destination = self._find_destination(command)
        if destination and destination in cell_refs:
            cell_refs.remove(destination)

        return {
            "operation": operation,
            "inputs": cell_refs,
            "output": destination,
        }