import re

from backend.excel.handler import ExcelHandler
from backend.parser.parser import CommandParser, ParserError
from backend.validator.validator import Validator, ValidationError


class HistoryEntry:
    """Represents a single calculation in the history."""

    def __init__(self, command, instruction, result, destination=None):
        self.command = command
        self.instruction = instruction
        self.result = result
        self.destination = destination

    def __repr__(self):
        dest = f" to {self.destination}" if self.destination else ""
        return f"{self.instruction.get('operation')}(...){dest} = {self.result}"


class ExcelCopilot:
    """Main orchestrator that ties parser, validator, and excel together."""

    BINARY_OPERATIONS = ("ADD", "SUBTRACT", "MULTIPLY", "DIVIDE")
    AGGREGATE_OPERATIONS = ("SUM", "AVERAGE", "MIN", "MAX", "COUNT")
    OP_SYMBOLS = {
        "ADD": "+",
        "SUBTRACT": "-",
        "MULTIPLY": "*",
        "DIVIDE": "/",
    }

    def __init__(self, filepath=None):
        self.excel = ExcelHandler(filepath)
        self.parser = CommandParser()
        self.validator = Validator(self.excel)
        self.history = []
        self.interactive_mode = False

    def load_workbook(self, filepath):
        """Load a workbook for processing."""
        self.excel.load(filepath)
        return self

    def process_command(self, command, sheet_name=None, allow_overwrite=False):
        """Process a natural language command end-to-end.
        Returns a dict with result, confirmation, and history.
        """
        try:
            # Step 1: Parse into structured instruction
            instruction = self.parser.parse(command)

            # Step 2: Validate
            operation = self.validator.validate_operation(
                instruction.get("operation")
            )
            instruction["operation"] = operation

            # Step 3: Resolve data and execute
            result = None
            destination = None
            written_count = 0

            if operation in self.AGGREGATE_OPERATIONS:
                result = self._handle_aggregate(
                    instruction, sheet_name, allow_overwrite
                )
                destination = instruction.get("destination")

            elif operation in (*self.BINARY_OPERATIONS, "DIFFERENCE", "GROWTH"):
                result, destination, written_count = self._handle_binary(
                    instruction, sheet_name, allow_overwrite
                )

            elif operation == "PERCENTAGE":
                result, destination, written_count = self._handle_percentage(
                    instruction, sheet_name, allow_overwrite
                )

            # Step 4: Record history
            entry = HistoryEntry(command, instruction, result, destination)
            self.history.append(entry)

            # Step 5: Build confirmation
            message = self._build_confirmation(
                instruction, result, destination, written_count
            )

            return {
                "success": True,
                "instruction": instruction,
                "result": result,
                "destination": destination,
                "written_count": written_count,
                "message": message,
                "history": entry,
            }

        except (ParserError, ValidationError) as e:
            return self._error(e.message)
        except ZeroDivisionError:
            return self._error("Cannot divide by zero. Please check your numbers.")
        except Exception as e:
            return self._error(f"An unexpected error occurred: {e}")

    # ------------------------------------------------------------------ #
    # Operation handlers
    # ------------------------------------------------------------------ #

    def _handle_aggregate(self, instruction, sheet_name, allow_overwrite=False):
        """Handle SUM/AVERAGE/MIN/MAX/COUNT on a column, named column,
        cell inputs, or whole sheet. Returns the computed result.
        """
        operation = instruction["operation"]

        if instruction.get("whole_sheet"):
            values, used_range = self._get_whole_sheet_values(sheet_name)
            if not values:
                raise ValidationError(
                    "There are no numeric values in this sheet."
                )
            instruction["used_range"] = used_range
            instruction["values"] = values
            result = self.validator.execute(instruction)
            self._maybe_write_formula(
                operation, used_range,
                instruction.get("destination"), sheet_name,
                instruction, allow_overwrite,
            )
            return result

        # Named source: "Sales ka total karo"
        if instruction.get("named_source"):
            column = self._resolve_named_column(instruction["named_source"], sheet_name)
            numeric_values, used_range = self.validator.validate_column(
                column, sheet_name
            )
            instruction["column"] = column
            instruction["used_range"] = used_range
            instruction["values"] = numeric_values
            result = self.validator.execute(instruction)
            self._maybe_write_formula(
                operation, used_range,
                instruction.get("destination"), sheet_name,
                instruction, allow_overwrite,
            )
            return result

        # Explicit column
        if instruction.get("column"):
            numeric_values, used_range = self.validator.validate_column(
                instruction["column"], sheet_name
            )
            instruction["used_range"] = used_range
            instruction["values"] = numeric_values
            result = self.validator.execute(instruction)
            self._maybe_write_formula(
                operation, used_range,
                instruction.get("destination"), sheet_name,
                instruction, allow_overwrite,
            )
            return result

        # Cell inputs
        if instruction.get("inputs"):
            values = []
            for ref in instruction["inputs"]:
                val = self._read_operand_value(ref, sheet_name)
                if isinstance(val, (int, float)):
                    values.append(val)
            instruction["values"] = values
            result = self.validator.execute(instruction)
            self._write_binary_formula_from_inputs(
                instruction["inputs"], instruction.get("output"),
                instruction.get("operation"), sheet_name, instruction,
                allow_overwrite,
            )
            return result

        raise ParserError(
            "I couldn't figure out which data to use. Specify a column, "
            "a name like 'Revenue', or cell references."
        )

    def _handle_binary(self, instruction, sheet_name, allow_overwrite):
        """Handle ADD/SUBTRACT/MULTIPLY/DIVIDE/DIFFERENCE/GROWTH."""
        operation = instruction["operation"]

        # Named-column binary operation (element-wise or aggregate)
        if instruction.get("named_inputs"):
            return self._handle_named_binary(instruction, sheet_name, allow_overwrite)

        # Growth with cells: "A2 se B2 kitni badhi" -> new=B2, old=A2
        if operation == "GROWTH":
            inputs = instruction["inputs"]
            if len(inputs) != 2:
                raise ParserError("Growth needs two values, like 'B2 vs A2'.")
            new_val = self._read_operand_value(inputs[0], sheet_name)
            old_val = self._read_operand_value(inputs[1], sheet_name)
            instruction["new"], instruction["old"] = new_val, old_val
            result = self.validator.execute(instruction)
            self._write_operand_formula(
                inputs[0], inputs[1], "subtract_symbol_for_growth",
                instruction.get("output"), sheet_name, instruction,
            )
            return result, instruction.get("output"), 0

        # Cell-based binary: read operand values and compute
        inputs = instruction["inputs"]
        if len(inputs) != 2:
            raise ParserError(
                f"I need two values for {operation}, like 'B2 + C2'."
            )
        a = self._read_operand_value(inputs[0], sheet_name)
        b = self._read_operand_value(inputs[1], sheet_name)

        if operation == "DIFFERENCE":
            instruction["a"], instruction["b"] = a, b
        elif operation in self.BINARY_OPERATIONS:
            instruction["a"], instruction["b"] = a, b
        else:
            raise ParserError(f"Operation {operation} requires named data.")

        result = self.validator.execute(instruction)
        destination = instruction.get("output")
        written_count = 0

        # Write formula when a destination is specified and inputs are cells
        if destination and self._is_cell_ref(inputs[0]):
            self.validator.validate_destination(
                destination, sheet_name, allow_overwrite
            )
            if operation == "DIFFERENCE":
                formula = f"=ABS({inputs[0]}-{inputs[1]})"
            elif operation in self.BINARY_OPERATIONS:
                formula = (
                    f"={inputs[0]}{self.OP_SYMBOLS[operation]}{inputs[1]}"
                )
            self.excel.write_formula(destination, formula, sheet_name)
            instruction["formula"] = formula
            written_count = 1

        return result, destination, written_count

    def _handle_named_binary(self, instruction, sheet_name, allow_overwrite):
        """Handle binary operations on named columns.

        For ADD/SUBTRACT/MULTIPLY/DIVIDE with a named output column, write
        row-by-row formulas into the output column (element-wise).
        For GROWTH, compare column totals.
        """
        operation = instruction["operation"]
        named_inputs = instruction["named_inputs"]

        if operation == "GROWTH":
            if len(named_inputs) < 2:
                raise ParserError(
                    "Growth needs two named columns, like 'February vs January'."
                )
            new_col = self._resolve_named_column(named_inputs[0], sheet_name)
            old_col = self._resolve_named_column(named_inputs[1], sheet_name)
            new_total, _, _ = self.excel.get_column_values(new_col, sheet_name)
            old_total, _, _ = self.excel.get_column_values(old_col, sheet_name)
            instruction["new"] = sum(new_total)
            instruction["old"] = sum(old_total)
            result = self.validator.execute(instruction)
            instruction["used_range"] = (
                f"{self.excel.get_column_letter(self.excel.get_column_number(new_col))}"
                f" vs {self.excel.get_column_letter(self.excel.get_column_number(old_col))}"
            )
            return result, None, 0

        # Element-wise for the other binary ops
        if len(named_inputs) < 2:
            raise ParserError(
                f"{operation} needs two named columns, like 'Revenue and Expense'."
            )

        in1 = self._resolve_named_column(named_inputs[0], sheet_name)
        in2 = self._resolve_named_column(named_inputs[1], sheet_name)

        # Determine the data range (union of both columns)
        values1, max1, min1 = self.excel.get_column_values(in1, sheet_name)
        values2, max2, min2 = self.excel.get_column_values(in2, sheet_name)
        start_row = min(min1, min2) if (min1 and min2) else (min1 or min2)
        end_row = max(max1, max2)
        if not start_row or not end_row:
            raise ValidationError(
                "I couldn't find numeric data for the named columns."
            )

        # Determine output column
        named_output = instruction.get("named_output")
        if named_output:
            out_col = self._resolve_named_column(named_output, sheet_name, create=True)
            instruction["output_column"] = out_col
        elif instruction.get("output"):
            out_col = instruction["output"][:1]
        else:
            raise ParserError(
                "Please specify an output column, like 'profit calculate kar do'."
            )

        count = 0
        for row in range(start_row, end_row + 1):
            if operation in self.BINARY_OPERATIONS:
                formula = (
                    f"={in1}{row}{self.OP_SYMBOLS[operation]}{in2}{row}"
                )
            else:
                formula = f"=ABS({in1}{row}-{in2}{row})"
            self.excel.write_formula(f"{out_col}{row}", formula, sheet_name)
            count += 1

        instruction["row_count"] = count
        instruction["out_col"] = out_col
        return count, None, count

    def _handle_percentage(self, instruction, sheet_name, allow_overwrite):
        """Handle percentage operations."""
        named_source = instruction.get("named_source")
        if named_source:
            column = self._resolve_named_column(named_source, sheet_name)
            values, max_row, min_row = self.excel.get_column_values(column, sheet_name)
            if not values:
                raise ValidationError(
                    f"There are no numeric values in the '{named_source}' column."
                )
            total = sum(values)
            out_col = instruction.get("output_column") or instruction.get("output")
            if out_col:
                out_col = out_col[:1]
                for row in range(min_row, max_row + 1):
                    formula = f"={column}{row}/{total}*100"
                    self.excel.write_formula(f"{out_col}{row}", formula, sheet_name)
                return total, None, (max_row - min_row + 1)
            instruction["total"] = total
            return total, None, 0
        raise ParserError(
            "Please specify which column to calculate the percentage of, "
            "like 'Sales ka percentage karma karo'."
        )

    # ------------------------------------------------------------------ #
    # Helpers
    # ------------------------------------------------------------------ #

    def _resolve_named_column(self, name, sheet_name=None, create=False):
        """Resolve a name to a column letter via the header row.
        A single column letter (e.g. 'B') is used directly.
        """
        if isinstance(name, str) and len(name) == 1 and name.isalpha():
            column = name.upper()
            # Verify it exists when resolving an existing column
            if not create:
                sheet = self.excel.get_sheet(sheet_name)
                if self.excel.get_column_number(column) > sheet.max_column:
                    raise ValidationError(
                        f"I couldn't find Column {column} in this sheet."
                    )
            return column

        column = self.excel.get_column_by_name(name, sheet_name)
        if column is None and create:
            # Add a new column with the given name
            sheet = self.excel.get_sheet(sheet_name)
            col_idx = sheet.max_column + 1
            from openpyxl.utils import get_column_letter
            column = get_column_letter(col_idx)
            sheet.cell(row=1, column=col_idx).value = name.capitalize()
        if column is None:
            raise ValidationError(
                f"I couldn't find a column named '{name}' in this sheet."
            )
        return column

    def _get_whole_sheet_values(self, sheet_name=None):
        """Get all numeric values in the sheet, skipping the header row."""
        used_range = self.excel.get_used_range()
        if used_range is None:
            raise ValidationError("This sheet is empty.")
        sheet = self.excel.get_sheet(sheet_name)
        values = []
        for row in sheet.iter_rows():
            if row[0].row == 1:
                continue
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    values.append(cell.value)
                elif isinstance(cell.value, str):
                    try:
                        values.append(float(cell.value))
                    except ValueError:
                        pass
        return values, used_range

    def _maybe_write_formula(self, operation, used_range, destination,
                             sheet_name, instruction, allow_overwrite=False):
        """Write an aggregate formula to a destination cell if requested."""
        if destination:
            self.validator.validate_destination(destination, sheet_name, allow_overwrite)
            formula = f"={operation}({used_range})"
            self.excel.write_formula(destination, formula, sheet_name)
            instruction["formula"] = formula

    def _write_binary_formula_from_inputs(self, inputs, output, operation,
                                          sheet_name, instruction,
                                          allow_overwrite=False):
        """Write a formula for an aggregate operation over explicit cells."""
        if output and self._is_cell_ref(inputs[0]) and self._is_cell_ref(inputs[1]):
            if operation == "SUM":
                formula = f"=SUM({inputs[0]}:{inputs[1]})"
            else:
                formula = f"={operation}({inputs[0]}:{inputs[1]})"
            self.validator.validate_destination(output, sheet_name, allow_overwrite)
            self.excel.write_formula(output, formula, sheet_name)
            instruction["formula"] = formula

    def _write_operand_formula(self, a, b, symbol, destination, sheet_name,
                               instruction):
        """Write a growth/difference formula between two cell refs."""
        if destination and self._is_cell_ref(a) and self._is_cell_ref(b):
            if symbol == "subtract_symbol_for_growth":
                formula = f"=({a}-{b})/{b}*100"
            self.validator.validate_destination(destination, sheet_name, True)
            self.excel.write_formula(destination, formula, sheet_name)
            instruction["formula"] = formula

    def _read_operand_value(self, ref, sheet_name=None):
        """Read a value that may be a cell reference or a numeric literal."""
        if isinstance(ref, (int, float)):
            return ref
        if not isinstance(ref, str):
            raise ValidationError(f"Invalid operand: {ref}")
        if self._is_cell_ref(ref):
            return self._get_numeric_value(ref, sheet_name)
        try:
            return float(ref)
        except ValueError:
            raise ValidationError(
                f"I couldn't understand the value '{ref}'."
            )

    def _is_cell_ref(self, ref):
        """Check if a string looks like a cell reference (e.g. B2)."""
        return isinstance(ref, str) and bool(
            re.fullmatch(r"[A-Z]{1,2}\d+", ref)
        )

    def _get_numeric_value(self, ref, sheet_name=None):
        """Get a numeric value from a cell reference."""
        val = self.excel.get_cell_value(ref, sheet_name)
        if isinstance(val, str):
            try:
                return float(val)
            except ValueError:
                raise ValidationError(
                    f"Cell {ref} doesn't contain a number."
                )
        if not isinstance(val, (int, float)):
            raise ValidationError(
                f"Cell {ref} doesn't contain a number."
            )
        return val

    def _error(self, message):
        """Build an error result dict."""
        return {
            "success": False,
            "message": message,
            "instruction": None,
            "result": None,
            "destination": None,
            "written_count": 0,
            "history": None,
        }

    def _build_confirmation(self, instruction, result, destination, written_count=0):
        """Build a confirmation message for the user."""
        op = instruction.get("operation")
        source = instruction.get("column") or instruction.get("source")
        used_range = instruction.get("used_range")

        if source and used_range:
            location = f"{op.title()} of {used_range}"
        elif instruction.get("named_source"):
            location = f"{op.title()} of the {instruction['named_source']} column"
        elif instruction.get("whole_sheet"):
            location = f"{op.title()} of all the data"
        else:
            inputs = instruction.get("inputs", [])
            if inputs:
                op_symbol = self.OP_SYMBOLS.get(op, " + ")
                display = " and ".join(str(i) for i in inputs[:2])
                location = f"{display} ({op.title()})"
            else:
                location = "calculation"

        if instruction.get("output_column") and written_count:
            return (
                f"Done. I wrote {written_count} formula(s) to "
                f"column {instruction['output_column']}."
            )

        if destination:
            return (
                f"Done. I calculated {location} = {result} "
                f"and placed the result in {destination}."
            )
        else:
            return f"Done. The result of {location} is {result}."

    # ------------------------------------------------------------------ #
    # Interactive loop
    # ------------------------------------------------------------------ #

    def run_interactive(self, filepath=None):
        """Run an interactive command loop."""
        if filepath:
            self.load_workbook(filepath)

        print("=" * 60)
        print("Excel AI Copilot")
        print("Type a command like 'Column D ka sum karo' or 'Sales ka total karo'.")
        print("Type 'history' to see past calculations.")
        print("Type 'exit' or 'quit' to leave.")
        print("=" * 60)

        while True:
            try:
                command = input("\n> ").strip()
            except (EOFError, KeyboardInterrupt):
                print("\nGoodbye!")
                break

            if not command:
                continue

            if command.lower() in ("exit", "quit", "bye"):
                print("Goodbye!")
                break

            if command.lower() in ("history", "hist"):
                self._print_history()
                continue

            result = self.process_command(command)

            # Overwrite confirmation (Phase 6)
            if (not result["success"]
                    and "already contains data" in result["message"]
                    and self.interactive_mode):
                print("-" * 40)
                print(result["message"])
                try:
                    choice = input("Replace it? [Yes/No]: ").strip().lower()
                except (EOFError, KeyboardInterrupt):
                    print("\nGoodbye!")
                    break
                if choice in ("y", "ye", "yes", "ha", "haan"):
                    result = self.process_command(
                        command, allow_overwrite=True
                    )
                else:
                    print("OK, I didn't change anything.")

            self._print_result(result)

            if result["success"] and self.excel.filepath:
                try:
                    self.excel.save()
                except Exception as e:
                    print(f"Note: could not save workbook: {e}")

    def _print_result(self, result):
        """Print a result nicely."""
        print("-" * 40)
        print(result["message"])
        if not result["success"]:
            print("Try again with a clearer command.")
        print("-" * 40)

    def _print_history(self):
        """Print the calculation history."""
        if not self.history:
            print("No calculations yet.")
            return

        print("\nCalculation History")
        print("-" * 40)
        for entry in self.history:
            dest = f" to {entry.destination}" if entry.destination else ""
            print(f"{entry.instruction.get('operation')}(...){dest} = {entry.result}")