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
        "jod": "ADD",
        "subtract": "SUBTRACT",
        "minus": "SUBTRACT",
        "ghatao": "SUBTRACT",
        "difference": "DIFFERENCE",
        "antar": "DIFFERENCE",
        "multiply": "MULTIPLY",
        "times": "MULTIPLY",
        "product": "MULTIPLY",
        "guna": "MULTIPLY",
        "divide": "DIVIDE",
        "div": "DIVIDE",
        "bhaag": "DIVIDE",
        "percentage": "PERCENTAGE",
        "percent": "PERCENTAGE",
        "persent": "PERCENTAGE",
        "kitna percent": "PERCENTAGE",
        "growth": "GROWTH",
        "increase": "GROWTH",
        "badhi": "GROWTH",
        "kitni badhi": "GROWTH",
        "kitni badh gayi": "GROWTH",
        "sort": "SORT",
        "sort karo": "SORT",
        "arrange": "SORT",
        "duplicates": "DEDUPE",
        "duplicate": "DEDUPE",
        "dedupe": "DEDUPE",
        "remove duplicates": "DEDUPE",
        "duplicate rows hata": "DEDUPE",
        "filter": "FILTER",
        "filter karo": "FILTER",
        "empty cells": "FIND_EMPTY",
        "khali cells": "FIND_EMPTY",
        "find empty": "FIND_EMPTY",
        "chart": "CHART",
        "chart bana": "CHART",
        "graph": "CHART",
        "analysis": "ANALYZE",
        "analyze": "ANALYZE",
        "analyse": "ANALYZE",
        "report bana": "ANALYZE",
        "calculate": None,
    }

    # Vocabulary of common data column names (English + Hinglish).
    NAMED_COLUMNS = {
        "january", "february", "march", "april", "june", "july",
        "august", "september", "october", "november", "december",
        "revenue", "expense", "profit", "sales", "cost", "income",
        "loss", "quantity", "price", "amount", "total",
        "bikri", "kharcha", "munafa", "aukhat", "sale", "bechne",
    }

    # Words indicating the whole sheet should be used.
    WHOLE_SHEET_WORDS = {"sabka", "sab", "sabhi", "all", "everything", "poora", "puri"}

    def parse(self, command):
        """Parse a natural language command into a structured instruction."""
        command = self._normalize(command)

        # Check for a cell reference in the form like B2
        cells = re.findall(r"\b([a-z])\s*(\d+)\b", command)

        # Check for a literal number like 100 or 2.5
        numbers = re.findall(r"\b(\d+(?:\.\d+)?)\b", command)

        # Check for column reference like "Column D" or just "D"
        column_matches = re.findall(r"\bcolumn\s+([a-z])\b", command)

        # Also look for bare column letters
        bare_columns = re.findall(r"\b([a-hj-z])\b", command)

        # Named data labels (revenue, expense, months, etc.)
        named = self._extract_named_operands(command)

        # Determine operation
        operation = self._detect_operation(command)

        if operation is None:
            raise ParserError(
                "I couldn't understand the command. Please try using words like "
                "sum, total, average, min, max, count, add, subtract, multiply, "
                "divide, difference, or percentage."
            )

        # Whole-sheet operations: "Sabka sum kar do"
        if operation in ("SUM", "AVERAGE", "MIN", "MAX", "COUNT", "PERCENTAGE"):
            if any(w in command for w in self.WHOLE_SHEET_WORDS) and not cells:
                destination = self._find_destination(command)
                return {
                    "operation": operation,
                    "whole_sheet": True,
                    "destination": destination,
                }

            # Named single column: "Sales ka total karo" or with a destination cell
            if named and not column_matches and not bare_columns:
                destination = self._find_destination(command)
                return {
                    "operation": operation,
                    "named_source": named[0],
                    "destination": destination,
                }

        # Data operations
        if operation in ("SORT", "DEDUPE", "FILTER", "FIND_EMPTY",
                         "CHART", "ANALYZE"):
            return self._parse_data_operation(command, operation, cells,
                                              named, column_matches, numbers)

        # Binary operations
        if operation in ("ADD", "SUBTRACT", "MULTIPLY", "DIVIDE",
                         "DIFFERENCE", "GROWTH"):
            return self._parse_binary(command, operation, cells, numbers, named)

        # Column-based calculations (SUM/AVERAGE/MIN/MAX/COUNT on a column)
        if column_matches:
            column = column_matches[0].upper()
        elif bare_columns:
            column = bare_columns[0].upper()
        else:
            column = None

        if column:
            destination = self._find_destination(command)
            return {
                "operation": operation,
                "source": column,
                "column": column,
                "destination": destination,
            }

        # Explicit cell references like "SUM(B2:C4)" style or fallback
        if cells:
            return self._parse_cell_references(command, operation, cells)

        raise ParserError(
            "I couldn't figure out which data to use. Please specify a column "
            "(like Column D), a cell range, or a name like 'Revenue'."
        )

    def _normalize(self, command):
        """Lowercase and strip punctuation/extra whitespace."""
        command = command.strip().lower()
        command = re.sub(r"[?!.,]+", "", command)
        return re.sub(r"\s+", " ", command)

    def _detect_operation(self, command):
        """Detect the operation type from a command."""
        # Order matters: check multi-word commands first
        for keyword, op in sorted(
            self.OPERATION_KEYWORDS.items(), key=lambda x: len(x[0]), reverse=True
        ):
            if keyword in command and op is not None:
                return op

        # Check for arithmetic operators
        if "+" in command or "plus" in command or "jod" in command:
            return "ADD"
        if "*" in command or "times" in command or "guna" in command:
            return "MULTIPLY"
        if "/" in command:
            return "DIVIDE"
        if "-" in command or "minus" in command or "ghatao" in command:
            return "SUBTRACT"

        return None

    def _extract_named_operands(self, command):
        """Return the ordered list of known data-label words in the command."""
        found = []
        for token in command.split():
            if token in self.NAMED_COLUMNS and token not in found:
                found.append(token)
        return found

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

    def _parse_binary(self, command, operation, cells, numbers, named):
        """Parse binary operations into structured instructions.

        Supports:
          - cell refs:      "B2 + C2", "B2 - C2"
          - cell + number:  "B2 + 100"
          - named columns:  "Revenue minus expense karke profit nikalo"
          - growth:         "February ki sales January se kitni badhi"
        """
        destination = self._find_destination(command)

        # Month growth pattern: "February ki sales January se kitni badhi"
        if operation == "GROWTH":
            months = re.findall(
                r"(\b(?:january|february|march|april|june|july|august|"
                r"september|october|november|december)\b)[^,]*?(\b(?:january|"
                r"february|march|april|june|july|august|september|october|"
                r"november|december)\b)",
                command,
            )
            if months:
                return {
                    "operation": "GROWTH",
                    "named_inputs": [months[0][0], months[0][1]],
                    "output": destination,
                }

        # Cell-based inputs
        if cells:
            cell_refs = [f"{c[0].upper()}{c[1]}" for c in cells]
            if destination in cell_refs:
                cell_refs.remove(destination)

            inputs = list(cell_refs)

            # Add literal numbers to fill up to two operands
            for num in numbers:
                if len(inputs) >= 2:
                    break
                try:
                    val = float(num)
                except ValueError:
                    continue
                if num not in cell_refs:
                    inputs.append(val)

            if len(inputs) < 2:
                raise ParserError(
                    "I need two values for this operation, "
                    "like 'B2 + C2' or 'B2 + 100'."
                )

            return {
                "operation": operation,
                "inputs": inputs[:2],
                "output": destination,
            }

        # Named-column inputs
        if named:
            named_inputs = list(named)
            named_output = None

            # If more than two labels appear, the trailing label is the output.
            if len(named_inputs) >= 3:
                named_output = named_inputs[-1]
                named_inputs = named_inputs[:-1]
            elif len(named_inputs) == 2 and operation in (
                "ADD", "SUBTRACT", "MULTIPLY", "DIVIDE"
            ):
                # Profit-like output: "Revenue minus expense karke profit nikalo"
                # if 'profit/munafa/loss' is near the end, treat as output.
                for out in ("profit", "munafa", "loss", "net", "result"):
                    if out in command:
                        named_output = out
                        break
                if named_output:
                    named_inputs = [n for n in named_inputs if n != named_output]

            return {
                "operation": operation,
                "named_inputs": named_inputs,
                "named_output": named_output,
                "output": destination,
            }

        # Column-letter operands: "B aur C ko multiply karke D mein daal do"
        bare_cols = re.findall(r"\b([a-hj-z])\b", command)
        if bare_cols:
            dest_col = self._find_column_destination(command)
            upper_cols = [c.upper() for c in bare_cols]
            if dest_col and dest_col.upper() in upper_cols:
                upper_cols.remove(dest_col.upper())

            if len(upper_cols) >= 2:
                return {
                    "operation": operation,
                    "named_inputs": upper_cols[:2],
                    "named_output": dest_col.upper() if dest_col else None,
                    "output": destination,
                }

        raise ParserError(
            "I need two values for this operation, like 'B2 + C2' "
            "or 'Revenue minus Expense'."
        )

    def _find_column_destination(self, command):
        """Find a destination column letter, e.g. 'D' in 'D mein daal do'."""
        patterns = [
            r"([a-hj-z])\s*(?:mein|me|daal do)\b",
            r"(?:mein|me|daal do)\s+([a-hj-z])\b",
        ]
        for pattern in patterns:
            for match in re.finditer(pattern, command):
                return match.group(1).upper()
        return None

    def _parse_data_operation(self, command, operation, cells, named,
                              column_matches, numbers):
        """Parse data operations (sort/dedupe/filter/find-empty/chart/analyze)."""
        destination = self._find_destination(command)
        dest_col = self._find_column_destination(command)

        # Determine the target column and value
        column = None
        if column_matches:
            column = column_matches[0].upper()
        elif named:
            column = named[0]
        else:
            bare = re.findall(r"\b([a-hj-z])\b", command)
            if bare:
                column = bare[0].upper()

        value = None
        operator = None

        # Filter conditions
        if operation == "FILTER":
            condition = self._parse_filter_condition(command, column)
            if condition:
                column, operator, value = condition

        if operation == "ANALYZE":
            return {"operation": "ANALYZE",
                    "column": column,
                    "whole_sheet": column is None}

        if operation in ("SORT", "DEDUPE", "FIND_EMPTY", "FILTER", "CHART"):
            if operation == "FILTER" and operator is None:
                raise ParserError(
                    "I couldn't find a filter condition. Try something like "
                    "'Filter data jahan Revenue 1000 se zyada'."
                )
            if operation in ("SORT", "CHART") and column is None:
                bare = re.findall(r"\b([a-hj-z])\b", command)
                if bare:
                    column = bare[0].upper()

            return {
                "operation": operation,
                "column": column,
                "whole_sheet": column is None,
                "value": value,
                "operator": operator,
                "destination": destination,
                "line_chart": "line" in command if operation == "CHART" else False,
            }

        raise ParserError(f"I couldn't understand the {operation} command.")

    def _parse_filter_condition(self, command, column):
        """Parse a filter condition into (column, operator, value).
        Examples:
          - "Revenue 1000 se zyada"      -> (Revenue, >, 1000)
          - "jahan C2 se zyada"          -> (C, >, 2)
          - "where B > 500"              -> (B, >, 500)
          - "<value> se kam"             -> (column, <, value)
        """
        # Pattern: <column> <number> se zyada / se kam
        m = re.search(r"\b([a-z])(\d+)\s+([\d.]+)\s+se\s+(zyada|kam|barabar)\b", command)
        if m:
            column = m.group(1).upper()
            value = float(m.group(3))
            operator = {"zyada": ">", "kam": "<", "barabar": "=="}[m.group(4)]
            return column, operator, value

        # Pattern: <column> <number> (with explicit operator symbols)
        m = re.search(r"\b([a-z])(\d+)\s*(>=|<=|==|=|>|<)\s*([\d.]+)\b", command)
        if m:
            column = m.group(1).upper()
            operator = m.group(2)
            value = float(m.group(4))
            return column, operator, value

        # Pattern: <word> <number> se zyada/kam (named column)
        m = re.search(r"\b(\w+)\s+([\d.]+)\s+se\s+(zyada|kam)\b", command)
        if m and column:
            value = float(m.group(2))
            operator = {"zyada": ">", "kam": "<"}[m.group(3)]
            return column, operator, value

        # Pattern: <word> se zyada/kam <number>
        m = re.search(r"\b(\w+)\s+se\s+(zyada|kam)\s+([\d.]+)\b", command)
        if m and column:
            value = float(m.group(3))
            operator = {"zyada": ">", "kam": "<"}[m.group(2)]
            return column, operator, value

        # Pattern: "jahan Revenue 1000 se zyada"
        m = re.search(r"jahan\s+(\w+)\s+([\d.]+)\s+se\s+(zyada|kam)\b", command)
        if m:
            column = m.group(1).upper()
            value = float(m.group(2))
            operator = {"zyada": ">", "kam": "<"}[m.group(3)]
            return column, operator, value

        return None

    def _parse_cell_references(self, command, operation, cells):
        """Parse commands with explicit cell references for aggregate ops."""
        cell_refs = [f"{c[0].upper()}{c[1]}" for c in cells]
        destination = self._find_destination(command)
        if destination and destination in cell_refs:
            cell_refs.remove(destination)

        return {
            "operation": operation,
            "inputs": cell_refs,
            "output": destination,
        }