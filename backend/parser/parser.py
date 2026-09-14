import re


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
        "sumif": "SUMIF",
        "sumifs": "SUMIF",
        "countif": "COUNTIF",
        "countifs": "COUNTIF",
        "averageif": "AVERAGEIF",
        "averageifs": "AVERAGEIF",
        "standardize dates": "STANDARDIZE_DATES",
        "dates ko standardize": "STANDARDIZE_DATES",
        "dates standardize": "STANDARDIZE_DATES",
        "dates normalize": "STANDARDIZE_DATES",
        "normalize dates": "STANDARDIZE_DATES",
        "standardize names": "STANDARDIZE_NAMES",
        "names ko standardize": "STANDARDIZE_NAMES",
        "names standardize": "STANDARDIZE_NAMES",
        "names normalize": "STANDARDIZE_NAMES",
        "capitalize names": "STANDARDIZE_NAMES",
        "invalid values": "DETECT_INVALID",
        "invalid value": "DETECT_INVALID",
        "invalid data": "DETECT_INVALID",
        "galt data": "DETECT_INVALID",
        "galtiyan": "DETECT_INVALID",
        "detect invalid": "DETECT_INVALID",
        "calculate": None,
    }

    # Vocabulary of common data column names (English + Hinglish).
    NAMED_COLUMNS = {
        "january",
        "february",
        "march",
        "april",
        "june",
        "july",
        "august",
        "september",
        "october",
        "november",
        "december",
        "revenue",
        "expense",
        "profit",
        "sales",
        "cost",
        "income",
        "loss",
        "quantity",
        "price",
        "amount",
        "total",
        "bikri",
        "kharcha",
        "munafa",
        "aukhat",
        "sale",
        "bechne",
    }

    # Words indicating the whole sheet should be used.
    WHOLE_SHEET_WORDS = {"sabka", "sab", "sabhi", "all", "everything", "poora", "puri"}

    # Strong write-intent markers. When one of these appears in a command but
    # no destination cell/column could be resolved, the command would silently
    # drop the requested write, so we reject it instead.
    WRITE_INTENT_MARKERS = (
        "daal do",
        "dal do",
        "daalo",
        "daaliye",
        "likho",
        "write",
        "place",
    )

    # Hindi particles that can follow "Column" without naming a column
    # (e.g. "column ka total karo" -> "column of the data, total it").
    HINDI_PARTICLES = {"ka", "ki", "ke", "ko", "se", "me", "mein", "na"}

    def __init__(self, headers=None):
        """Initialize the parser.

        ``headers`` optionally lists the current workbook's column headers
        (case-insensitive) so data operations like chart/sort/filter can use
        arbitrary column names without relying on the built-in vocabulary.
        """
        self.headers = set()
        if headers:
            self.set_headers(headers)

    def set_headers(self, headers):
        """Replace the known workbook column headers."""
        self.headers = set()
        for header in headers or []:
            if header:
                self.headers.add(header.strip().lower())

    def parse(self, command):
        """Parse a natural language command into a structured instruction."""
        command = self._normalize(command)

        # Conditional aggregates take priority: a command with
        # "jahan"/"where" plus a SUM/AVERAGE/COUNT keyword is conditional,
        # even when other words (e.g. "product") collide with op keywords.
        if "jahan" in command or "where" in command:
            agg = None
            for kw, op in self.OPERATION_KEYWORDS.items():
                if op in ("SUM", "AVERAGE", "COUNT") and kw in command:
                    agg = op
                    break
            if agg is not None:
                op = {"SUM": "SUMIF", "COUNT": "COUNTIF", "AVERAGE": "AVERAGEIF"}[agg]
                return self._parse_conditional(command, op)

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

        # Conditional calculations: "jahan"/"where" clause with an aggregate
        if operation in ("SUM", "AVERAGE", "COUNT", "SUMIF", "COUNTIF", "AVERAGEIF"):
            if (
                operation in ("SUMIF", "COUNTIF", "AVERAGEIF")
                or "jahan" in command
                or "where" in command
            ):
                return self._parse_conditional(command, operation)

        # Whole-sheet operations: "Sabka sum kar do"
        if operation in ("SUM", "AVERAGE", "MIN", "MAX", "COUNT", "PERCENTAGE"):
            if any(w in command for w in self.WHOLE_SHEET_WORDS) and not cells:
                destination = self._find_destination(command)
                self._reject_unresolved_write(command, destination)
                return {
                    "operation": operation,
                    "whole_sheet": True,
                    "destination": destination,
                }

            # Named single column: "Sales ka total karo" or with a destination
            # cell. Falls back to actual workbook headers so arbitrary column
            # names (Salary, Bonus, ...) work for aggregates and percentages
            # without being part of the built-in vocabulary.
            if not column_matches and not bare_columns:
                # Multi-letter column refs ("Column XYZ") are not real columns
                # unless this workbook actually has them; reject with a clear
                # message instead of guessing at a keyword like 'total'.
                multi = re.match(r"column\s+([a-z]{2,})\b", command)
                if multi and multi.group(1) not in self.HINDI_PARTICLES:
                    raise ParserError(
                        f"I couldn't find Column {multi.group(1).upper()} in this sheet."
                    )
                source = (
                    named[0] if named else (None if cells else self._find_header_operand(command))
                )
                if source:
                    destination = self._find_destination(command)
                    self._reject_unresolved_write(command, destination)
                    return {
                        "operation": operation,
                        "named_source": source,
                        "destination": destination,
                    }

        # Data operations
        if operation in (
            "SORT",
            "DEDUPE",
            "FILTER",
            "FIND_EMPTY",
            "CHART",
            "ANALYZE",
            "STANDARDIZE_DATES",
            "STANDARDIZE_NAMES",
            "DETECT_INVALID",
        ):
            return self._parse_data_operation(
                command, operation, cells, named, column_matches, numbers
            )

        # Binary operations
        if operation in ("ADD", "SUBTRACT", "MULTIPLY", "DIVIDE", "DIFFERENCE", "GROWTH"):
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
            self._reject_unresolved_write(command, destination)
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

    def _find_header_operand(self, command):
        """Return the first command token that matches a known workbook header.

        Enables aggregates and percentages to reference columns by their actual
        header names (e.g. 'Bonus ka percentage karo') even when they are not
        part of the built-in ``NAMED_COLUMNS`` vocabulary.
        """
        for token in command.split():
            if token in self.headers:
                return token
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

    def _parse_binary(self, command, operation, cells, numbers, named):
        """Parse binary operations into structured instructions.

        Supports:
          - cell refs:      "B2 + C2", "B2 - C2"
          - cell + number:  "B2 + 100"
          - named columns:  "Revenue minus expense karke profit nikalo"
          - growth:         "February ki sales January se kitni badhi"
        """
        destination = self._find_destination(command)
        dest_col = self._find_column_destination(command)
        self._reject_unresolved_write(command, destination, dest_col)

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
                    "I need two values for this operation, like 'B2 + C2' or 'B2 + 100'."
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
            elif len(named_inputs) == 2 and operation in ("ADD", "SUBTRACT", "MULTIPLY", "DIVIDE"):
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
            "I need two values for this operation, like 'B2 + C2' or 'Revenue minus Expense'."
        )

    def _reject_unresolved_write(self, command, destination, dest_col=None):
        """Raise when a command asks to write somewhere but no destination
        cell/column could be found, so the write is never silently dropped."""
        if not destination and not dest_col and self._has_write_intent(command):
            raise ParserError(
                "I couldn't make sense of the destination. Please use a cell "
                "like 'D21' or a column letter, for example "
                "'Revenue ka total karo aur D21 mein daal do'."
            )

    @staticmethod
    def _has_write_intent(command):
        """True when the command uses a strong write marker such as 'daal do'."""
        return any(marker in command for marker in CommandParser.WRITE_INTENT_MARKERS)

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

    def _parse_data_operation(self, command, operation, cells, named, column_matches, numbers):
        """Parse data operations (sort/dedupe/filter/find-empty/chart/analyze)."""
        destination = self._find_destination(command)

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

        # Fall back to actual workbook headers so arbitrary column names
        # (Salary, Orders, ...) work for data operations without needing to
        # be part of the built-in NAMED_COLUMNS vocabulary.
        if column is None and self.headers:
            for token in command.split():
                if token in self.headers:
                    column = token
                    break

        value = None
        operator = None

        # Filter conditions
        if operation == "FILTER":
            condition = self._parse_filter_condition(command, column)
            if condition:
                column, operator, value = condition

        descending = self._has_descending(command) if operation == "SORT" else False

        if operation == "ANALYZE":
            return {"operation": "ANALYZE", "column": column, "whole_sheet": column is None}

        if operation in (
            "SORT",
            "DEDUPE",
            "FIND_EMPTY",
            "FILTER",
            "CHART",
            "STANDARDIZE_DATES",
            "STANDARDIZE_NAMES",
            "DETECT_INVALID",
        ):
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
                "descending": descending,
                "line_chart": "line" in command if operation == "CHART" else False,
            }

        raise ParserError(f"I couldn't understand the {operation} command.")

    @staticmethod
    def _has_descending(command):
        """True when a sort command asks for descending/reverse order."""
        return any(token in ("descending", "desc", "reverse", "ulta") for token in command.split())

    def _parse_conditional(self, command, operation):
        """Parse a command with a 'jahan'/where clause into a conditional
        calculation (SUMIF / COUNTIF / AVERAGEIF).

        Examples:
          - "Column B ka sum karo jahan A Pen hai"
          - "Revenue ka average karo jahan month March hai"
          - "Count karo jahan B 100 se zyada"
          - "B ka sum karo jahan A Pen hai aur C 50 se kam"
        """
        operation = {"SUM": "SUMIF", "COUNT": "COUNTIF", "AVERAGE": "AVERAGEIF"}.get(
            operation, operation
        )
        destination = self._find_destination(command)
        target_col = None

        head = command
        tail = command
        m = re.search(r"\b(jahan|where)\b", command)
        if m:
            head = command[: m.start()]
            tail = command[m.end() :]

        # The target column comes from the head of the command.
        col_match = re.search(r"\bcolumn\s+([a-z])\b", head)
        if col_match:
            target_col = col_match.group(1).upper()
        else:
            head_bare = re.findall(r"\b([a-hj-z])\b", head)
            if head_bare:
                target_col = head_bare[0].upper()
            else:
                named = self._extract_named_operands(head)
                if named:
                    target_col = named[0]

        # For COUNTIF there may be no explicit target column.
        criteria = self._parse_condition_clauses(tail)
        if not criteria:
            raise ParserError(
                "I couldn't find a condition. Try something like "
                "'B ka sum karo jahan A Pen hai' or 'B ka sum karo "
                "jahan A 100 se zyada'."
            )

        # COUNTIF can count over the criteria column itself.
        if operation == "COUNTIF" and target_col is None:
            target_col = criteria[0][0]

        if target_col is None:
            raise ParserError(
                "Please specify which column to calculate, like "
                "'Column B ka sum karo jahan Price 100 se zyada'."
            )

        return {
            "operation": operation,
            "sum_range": target_col,
            "criteria": criteria,
            "destination": destination,
        }

    def _parse_condition_clauses(self, tail):
        """Parse one-or-more conditions from the tail of a command.
        Each condition is (column, operator, value) where column may be a
        letter or a named label, and value may be a number or text.
        """
        clauses = re.split(r"\b(?:aur|and)\b", tail)
        criteria = []
        for clause in clauses:
            cond = self._parse_condition_clause(clause)
            if cond:
                criteria.append(cond)
        return criteria

    def _parse_condition_clause(self, clause):
        """Parse a single condition clause into (column, operator, value)."""
        clause = clause.strip().strip(",. -")
        if not clause:
            return None

        # Numeric comparisons with symbols: "B>100", "Price >= 100"
        m = re.search(r"(\w+)\s*(>=|<=|==|=|>|<)\s*(-?[\d.]+)\b", clause)
        if m:
            return self._cond_operand(m.group(1)), m.group(2), float(m.group(3))

        # Hinglish: "B 100 se zyada", "Price 1000 se kam"
        m = re.search(r"(\w+)\s+([\d.]+)\s+se\s+(zyada|kam|barabar)\b", clause)
        if m:
            op = {"zyada": ">", "kam": "<", "barabar": "=="}[m.group(3)]
            return self._cond_operand(m.group(1)), op, float(m.group(2))

        # Hinglish: "B zyada 100", "Price kam 1000"
        m = re.search(r"(\w+)\s+(zyada|kam|barabar)\s+([\d.]+)\b", clause)
        if m:
            op = {"zyada": ">", "kam": "<", "barabar": "=="}[m.group(2)]
            return self._cond_operand(m.group(1)), op, float(m.group(3))

        # Text equality in Hinglish: "A Pen hai", "Product Pen hai"
        m = re.search(r"(\w+)\s+(?:ke?\s+)?(?:mein|me)\s+([\w.\-]+)\s+(?:hai|ho|hain)\b", clause)
        if m:
            return self._cond_operand(m.group(1)), "==", m.group(2)

        # Simpler Hinglish equality: "A Pen hai"
        m = re.search(r"(\w+)\s+([\w.\-]+)\s+(?:hai|ho|hain)\b", clause)
        if m:
            return self._cond_operand(m.group(1)), "==", m.group(2)

        # Plain text equality: "A = Pen", "A Pen", "A equals Pen"
        m = re.search(r"(\w+)\s*(?:==|=|equals|ke barabar)\s*([\w.\-]+)\b", clause)
        if m:
            return self._cond_operand(m.group(1)), "==", m.group(2)

        # Hinglish bare equality: "jahan Product Pen"
        tokens = clause.split()
        if len(tokens) == 2 and tokens[0] != "column":
            return self._cond_operand(tokens[0]), "==", tokens[1]

        return None

    @staticmethod
    def _cond_operand(token):
        """Trim filler words from a condition column token."""
        return token.strip(":;,").lower()

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
